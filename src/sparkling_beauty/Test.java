package sparkling_beauty;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

public class Test {

	public static void main(String[] args) {
		// <title>, og:title, keywords, description, 제목, img, 내용, 태그

		//File input = new File("/Users/uzini/Desktop/id/Sparkling Beauty.html");
		File input = new File("/Users/uzini/Desktop/test.html");
		
		try {
/*
			//스파클링 뷰티 html 문서 불러오기
			Document doc = Jsoup.parse(input, "UTF-8", "https://www.laneige.com/id/id/sparkling-beauty.html");
			
			// 각 item의 href 파싱
			Elements el = doc.select(".sparkling-list-item").not(".sparkling-banner");
			String[] hrefs = new String[el.size()];
			
			// test
			//System.out.println(el.get(207));
			
			hrefs[207] = el.get(207).select("a").attr("href");
			Document itemDoc = Jsoup.connect(hrefs[207]).get();
			
			String[][] itemList = new String[el.size()][8];
			System.out.println("title: " + itemDoc.select(".content_Title").html());
			System.out.println("content: " + itemDoc.select(".sparkling-view-context").select("p").html());
			*/
			
			Document doc = Jsoup.parse(input, "UTF-8");
			Elements el = doc.select(".hi");
			//System.out.println(el.html());
			
			String a = replaceStr(el.html());	
			
			System.out.println(a);
			
			
			
			
			/*
			// href 문자열 저장
			for(int i=0; i<el.size(); i++) {
				hrefs[i] = el.get(i).select("a").attr("href");
			}
			
			// href 접속
			int colSize = 8;
			Document itemDoc;
			String[][] itemList = new String[el.size()][colSize];
			String src;
			Elements li;
			String tag = "";

			for(int i=0; i<el.size(); i++) {
				itemDoc = Jsoup.connect(hrefs[i]).get();
				// [0] <title>
				itemList[i][0] = itemDoc.title();
				System.out.println("["+i+"] <title> : " + itemList[i][0]);
				// [1] og:title
				itemList[i][1] = itemDoc.select("meta[property^=og:title]").attr("content");
				System.out.println("["+i+"] og:title : " + itemList[i][1]);
				// [2] description
				itemList[i][2] = itemDoc.select("meta[property^=og:description]").attr("content");
				System.out.println("["+i+"] og:description : " + itemList[i][2]);
				// [3] 제목
				itemList[i][3] = itemDoc.select(".content_Title").html();
				System.out.println("["+i+"] 제목 : " + itemList[i][3]);
				// [4] img src
				src = itemDoc.select(".custom-sparkling-view-imgtype").select("img").attr("src");
				itemList[i][4] = src.substring(src.lastIndexOf("/")+1);
				System.out.println("["+i+"] src : " + itemList[i][4]);
				// [5] 내용 p
				itemList[i][5] = itemDoc.select(".sparkling-view-context").select("p").html();
				System.out.println("["+i+"] 내용 : " + itemList[i][5]);
				// [6] 태그
				li = itemDoc.select(".sparkling-hash").select("li");
				tag = "";
				for(int j=0; j<li.size(); j++) {
					tag += li.get(j).select("span").html() + " | ";
				}
				itemList[i][6] = tag;
				System.out.println("["+i+"] tag " + itemList[i][6]);
 			}
			
			// 엑셀 삽입
			
	        HSSFWorkbook workbook = new HSSFWorkbook(); // 새 엑셀 생성
	        HSSFSheet sheet = workbook.createSheet("sheet1"); // 새 시트(Sheet) 생성
	        
	        HSSFRow row;
	        HSSFCell cell;
	        
	        for(int i=0; i<el.size(); i++) {
	        	row = sheet.createRow(i);
	        	for(int j=0; j<colSize; j++) {
	        		cell = row.createCell(j);
	        		cell.setCellValue(itemList[i][j]);
	        	}
	        }
	        try {
	            FileOutputStream fileoutputstream = new FileOutputStream("/Users/uzini/Desktop/Sparkling Beauty/id/test_id_data.xlsx");
	            workbook.write(fileoutputstream);
	            fileoutputstream.close();
	            System.out.println("엑셀파일생성성공");
	        } catch (IOException e) {
	            e.printStackTrace();
	            System.out.println("엑셀파일생성실패");
	        }
			*/
			
						
			
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
	
	public static String replaceStr(String a) {
		a= a.replaceAll("\"", "&quot;");
		a= a.replaceAll("\'", "&#39;");
		a= a.replaceAll("\\(", "&#40;");
		a= a.replaceAll("\\)", "&#41;");
		a= a.replaceAll("\\<", "&lt;");
		a= a.replaceAll("\\>", "&gt;");
		
		return a;
	}

}
