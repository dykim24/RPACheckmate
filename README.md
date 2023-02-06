# RPACheckmate

## 엑셀

### 코드 예시

#### 기본 엑셀 열고 정보입력후 닫기

using System;<enter>
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

public partial class CustomScript
{
	public void Execute_Code()
	{
		Excel.Application excel = new Excel.Application(); // 엑셀 앱이 백그라운드에 존재
		Workbook workbook = excel.Workbooks.Add(); // 엑셀 열었음
		//excel.Workbooks.Open()
		Worksheet worksheet = workbook.Sheets [1] as Worksheet; // 엑셀 안 시트 1
		
		string[] arr = { "a","b","c","d","e"};
		for(int i=1;i<=5;i++)
		{	
			Range cell = worksheet.Cells [i, 1] as Range;
			cell.Value2 = arr[i-1];
		}
		
		for(int i=1;i<=5;i++)
		{	
			for(int j=1;j<=5;j++)
			{	
				Range cell = worksheet.Cells [i, j] as Range;
				cell.Value2 = arr[i-1];
			}
			
		}
		//Range cell = worksheet.Cells [1, 1] as Range;
		//Range cell = worksheet.Range ["A1:D1"] as Range;
		
		//cell.Value2 = "dsinofidoaf";
		
		workbook.SaveAs(@"C:\Users\a\Desktop\test\test.xlsx");
		
		workbook.Close();
		excel.Quit();
	}
}
