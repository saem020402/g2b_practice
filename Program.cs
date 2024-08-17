using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Text.Json.Serialization;
using OfficeOpenXml;

class Program
{
    public class Item //공고에 대한 정보(공고명, 공고기관, 수요기관, 입력일, 입력시)
    {
        [JsonPropertyName("bidNtceNm")]
        public string? BidNtceNm { get; set; }

        [JsonPropertyName("ntceInsttNm")]
        public string? NtceInsttNm { get; set; }

        [JsonPropertyName("dmndInsttNm")]
        public string? DmndInsttNm { get; set; }

        [JsonPropertyName("bidNtceDate")]
        public string? BidNtceDate { get; set; }

        [JsonPropertyName("bidNtceBgn")]
        public string? BidNtceBgn { get; set; }
    }

    public class Body //공고항목과 페이지정보
    {
        [JsonPropertyName("items")] //필요한 정보들을 리스트로 저장하기 위해
        public List<Item>? Items { get; set; }

        [JsonPropertyName("numOfRows")] //한번에 가져오는 row수
        public int NumOfRows { get; set; }

        [JsonPropertyName("pageNo")] //page수
        public int PageNo { get; set; }

        [JsonPropertyName("totalCount")] //데이터량
        public int TotalCount { get; set; }
    }

    public class Response //API 응답의 본체
    {
        [JsonPropertyName("body")]
        public Body? Body { get; set; }
    }

    public class Root //전체응답
    {
        [JsonPropertyName("response")]
        public Response? Response { get; set; }
    }

    static async Task Main(string[] args)
    {
        //날짜입력
        Console.WriteLine("Enter the end date (YYYYMMDD): ");
        string endDateInput = Console.ReadLine();

        //엑셀파일경로 입력
        Console.WriteLine("Enter the path to save the Excel file (e.g., C:\\path\\to\\file.xlsx): ");
        string excelFilePath = Console.ReadLine();
        string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

        //입력날짜 형식 검증
        if (!DateTime.TryParseExact(endDateInput, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out DateTime endDate))
        {
            Console.WriteLine("Invalid date format. Please enter the date in YYYYMMDD format.");
            return;
        }

        //시작날짜와 종료날짜 설정
        string bidNtceEndDt = endDate.ToString("yyyyMMdd") + "2359";
        DateTime startDate = endDate.AddDays(-5);
        string bidNtceBgnDt = startDate.ToString("yyyyMMdd") + "0000";

        //endpoint와 request parameter설정
        string endpoint = "https://apis.data.go.kr/1230000/PubDataOpnStdService/getDataSetOpnStdBidPblancInfo";
        string serviceKey = "xru7CFerj%2FYWuU51dfT4NBZMnSNdLYseoTxNIMw02m%2BfZAJKNGBPKnjgVLHpEf7t2vOqupIpkKxOXA6kAOhqow%3D%3D";
        int numOfRows = 999; //한번에 가져올 row수
        string type = "json"; // 응답형식
        int pageNo = 1; //첫페이지부터 시작
        bool hasMorePages = true; //더 가져올 페이지가 있는지 여부

        //FileInfo를 이용해 파일 존재 여부 확인
        FileInfo fileInfo = new FileInfo(excelFilePath);
        //ExcelPackage를 이용해 파일을 열거나 없으면 새로 생성
        using (var package = new ExcelPackage(fileInfo)) //using- 객체에 블록을 지정해 메모리 낭비를 방지
        {
            ExcelWorksheet worksheet;
            //파일이 존재하면 첫번째 시트를 가져오고, 그렇지 않으면 새로운 시트를 생성하여 헤더 추가
            if (fileInfo.Exists)
            {
                worksheet = package.Workbook.Worksheets[0];
            }
            else
            {
                worksheet = package.Workbook.Worksheets.Add("Bid Notices");
                //헤더
                worksheet.Cells[1, 1].Value = "공고명";
                worksheet.Cells[1, 2].Value = "검색키워드";
                worksheet.Cells[1, 3].Value = "공고기관";
                worksheet.Cells[1, 4].Value = "수요기관";
                worksheet.Cells[1, 5].Value = "입력일시";
                worksheet.Cells[1, 6].Value = "검색일시";
            }
            //데이터 추가 위치 결정
            int row = worksheet.Dimension?.End.Row + 1 ?? 2;

            while (hasMorePages)
            {
                //요청할 url
                string url = $"{endpoint}?serviceKey={serviceKey}&pageNo={pageNo}&numOfRows={numOfRows}&type={type}&bidNtceBgnDt={bidNtceBgnDt}&bidNtceEndDt={bidNtceEndDt}";

                //Http 요청
                using (HttpClient client = new HttpClient())
                {
                    try
                    {
                        //Get 요청
                        HttpResponseMessage response = await client.GetAsync(url);
                        response.EnsureSuccessStatusCode(); //응답이 성공적인지 확인(200-299)

                        //응답내용 읽기
                        string responseBody = await response.Content.ReadAsStringAsync(); //응답본문을 문자열로 읽음

                        //json응답을 Root타입 객체로 파싱
                        var rootResponse = JsonSerializer.Deserialize<Root>(responseBody);

                        if (rootResponse?.Response?.Body != null)
                        {
                            if (rootResponse.Response.Body.Items != null)
                            {
                                foreach (var item in rootResponse.Response.Body.Items)
                                {
                                    //rpa포함하는 것만 필터링 후 엑셀에 추가
                                    if (item.BidNtceNm != null && item.BidNtceNm.Contains("RPA"))
                                    {
                                        worksheet.Cells[row, 1].Value = item.BidNtceNm; //공고명
                                        worksheet.Cells[row, 2].Value = "RPA"; //검색키워드
                                        worksheet.Cells[row, 3].Value = item.NtceInsttNm; //공고기관
                                        worksheet.Cells[row, 4].Value = item.DmndInsttNm; //수요기관
                                        worksheet.Cells[row, 5].Value = item.BidNtceDate; //입력일
                                        worksheet.Cells[row, 5].Value += " " + item.BidNtceBgn; //입력시
                                        worksheet.Cells[row, 6].Value = now; //검색일시
                                        row++;
                                    }
                                }
                            }
                            //다음페이지로 이동 여부 결정(남은 항목이 있는지)
                            hasMorePages = rootResponse.Response.Body.PageNo * numOfRows < rootResponse.Response.Body.TotalCount;
                            pageNo++;
                        }
                        else
                        {
                            Console.WriteLine("No items found or response is empty.");
                            hasMorePages = false;
                        }
                    }
                    catch (HttpRequestException e)
                    {
                        Console.WriteLine("Request error: " + e.Message);
                        if (e.InnerException != null)
                        {
                            Console.WriteLine("Inner exception: " + e.InnerException.Message);
                        }
                        hasMorePages = false;
                    }
                    catch (JsonException e)
                    {
                        Console.WriteLine("JSON parsing error: " + e.Message);
                        hasMorePages = false;
                    }
                }
            }
            package.Save();
            Console.WriteLine($"Data has been saved to {excelFilePath}"); //저장성공
        }
    }
}
