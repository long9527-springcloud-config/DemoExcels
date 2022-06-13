package com.example.demoexcel.test;
 
import java.io.FileOutputStream;
import java.io.IOException;
 
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.hssf.usermodel.HSSFTextbox;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextBox;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing;

public class TestHSSFTextBox {
 
	public static void main(String[] args) throws IOException {
		// 创建一个工作博
		Workbook workbook = new HSSFWorkbook();
		// 创建一个sheet
		Sheet sheet = workbook.createSheet();
		// 画图的顶级管理器对象HSSFPatriarch, 一个sheet只能获取一个
		HSSFPatriarch hssfPatriarch = (HSSFPatriarch) sheet.createDrawingPatriarch();
		
		/****************************************文本框水平垂直对齐**************************************************/

		// 形状在sheet中的锚点位置
		HSSFClientAnchor anchor1 = new HSSFClientAnchor(0, 0, 0, 0, (short)1, 2, (short)4, 8);
		// 创建一个文本框


		HSSFTextbox textbox1 = hssfPatriarch.createTextbox(anchor1);

		HSSFRichTextString richTextString = new HSSFRichTextString("港版支付宝上线后，所有香港居民都可以通过绑定香港当地银行信用卡或余额充值使用支付宝，直接用港币付款，使用更加方便。");
		textbox1.setString(richTextString);
		
		// 设置文本框水平垂直对齐
		textbox1.setHorizontalAlignment(HSSFTextbox.HORIZONTAL_ALIGNMENT_CENTERED);
		textbox1.setVerticalAlignment(HSSFTextbox.VERTICAL_ALIGNMENT_CENTER);

        XSSFClientAnchor xssfClientAnchor = new XSSFClientAnchor();

		/****************************************文本框文字留白**************************************************/
		HSSFClientAnchor anchor2 = new HSSFClientAnchor(0, 0, 0, 0, (short)5, 2, (short)8, 8);
		HSSFTextbox textbox2 = hssfPatriarch.createTextbox(anchor2);
		textbox2.setString(richTextString);
		
		// 设置文本框留白, HSSFShape.LINEWIDTH_ONE_PT是形状中单位，1pt=1/72英寸
		textbox2.setMarginLeft(10 * HSSFShape.LINEWIDTH_ONE_PT);
		textbox2.setMarginTop(10 * HSSFShape.LINEWIDTH_ONE_PT);
		textbox2.setMarginRight(10 * HSSFShape.LINEWIDTH_ONE_PT);
		textbox2.setMarginBottom(10 * HSSFShape.LINEWIDTH_ONE_PT);
 
		/****************************************文本框边框样式及填充色**************************************************/
		HSSFClientAnchor anchor3 = new HSSFClientAnchor(0, 0, 0, 0, (short)10, 2, (short)14, 8);
		HSSFTextbox textbox3 = hssfPatriarch.createTextbox(anchor3);
		textbox3.setString(richTextString);
		
		// 设置填充颜色 - 绿色
		textbox3.setFillColor(106, 168, 79);
		// 设置文本框边框颜色 - 蓝色
		textbox3.setLineStyleColor(0, 0, 255);
		// 设置文本框边框宽度 - 3pt
		textbox3.setLineWidth(3 * HSSFShape.LINEWIDTH_ONE_PT);
		// 设置文本框边框样式 - 长破折号和点间隔
		textbox3.setLineStyle(HSSFShape.LINESTYLE_LONGDASHDOTGEL);
		
		/****************************************文本框setWrapText**************************************************/
		HSSFClientAnchor anchor4 = new HSSFClientAnchor(0, 0, 0, 0, (short)15, 2, (short)19, 8);
		HSSFTextbox textbox4 = hssfPatriarch.createTextbox(anchor4);
		HSSFRichTextString richTextString1 = new HSSFRichTextString("daxdd afcsdtadt svxgy sfyfx yfxts ffst fxs hgfx gfs fgsddsr gsfxxs gsfxts sjs");
		textbox4.setString(richTextString1);
		
		// 暂时没看出什么效果
//		textbox4.setWrapText(HSSFSimpleShape.WRAP_SQUARE);
//		textbox4.setWrapText(HSSFSimpleShape.WRAP_BY_POINTS);
		textbox4.setWrapText(HSSFSimpleShape.WRAP_NONE);
		
		/****************************************文本框旋转**************************************************/
		HSSFClientAnchor anchor5 = new HSSFClientAnchor(0, 0, 0, 0, (short)1, 15, (short)4, 20);
		HSSFTextbox textbox5 = hssfPatriarch.createTextbox(anchor5);
		textbox5.setString(richTextString1);
		
		/**
		 * 设置文本框旋转多少度，围绕形状的中心旋转，该属性的默认值为0x00000000
		 * 正值：顺时针旋转
		 * 负值：逆时针旋转
		 * */
		textbox5.setRotationDegree((short)44);
		
		HSSFClientAnchor anchor6 = new HSSFClientAnchor(0, 0, 0, 0, (short)5, 15, (short)8, 20);
		HSSFTextbox textbox6 = hssfPatriarch.createTextbox(anchor6);
		textbox6.setString(richTextString1);
		
		// 被他的选择搞懵了，求解
		textbox6.setRotationDegree((short)45);
		
		/****************************************文本框水平反转和垂直翻转**************************************************/
		HSSFClientAnchor anchor7 = new HSSFClientAnchor(0, 0, 0, 0, (short)10, 15, (short)14, 20);
		HSSFTextbox textbox7 = hssfPatriarch.createTextbox(anchor7);
		textbox7.setString(richTextString1);
		
		/**
		 * 设置文本框是否水平翻转或垂直翻转
		 * true:表示将文本框水平或垂直翻转了
		 * false：不进行翻转
		 * */
		textbox7.setFlipHorizontal(true);
		
		HSSFClientAnchor anchor8 = new HSSFClientAnchor(0, 0, 0, 0, (short)15, 15, (short)19, 20);
		HSSFTextbox textbox8 = hssfPatriarch.createTextbox(anchor8);
		textbox8.setString(richTextString1);
		
		// 设置垂直翻转
		textbox8.setFlipVertical(true);
		
		/********************************************************************************************************/
		
		FileOutputStream file = new FileOutputStream("C://Users//long//Desktop//test.xls");
		workbook.write(file);
		file.close();
	}
}