import java.io.*;
import jxl.*;
import java.text.DecimalFormat;

def readInput(workbook,inputSheetName)
{    
    Sheet sheet=workbook.getSheet(inputSheetName);
            
    rows = sheet.getRows();
    columns = sheet.getColumns();
    
/*Get content into array*/
    input = new Object [columns][rows]
    for (a=0;a<columns;a++)
        {
        for (b=0;b<rows;b++)
            {
                input[a][b]= sheet.getCell(a,b).getContents();
            }
        }
}

def readBaseline(workbook,sheetName)
{    
    Sheet sheet=workbook.getSheet(sheetName);
            
    Rows = sheet.getRows(); 
    Columns = sheet.getColumns();
    
/*Get content into array*/
    baseline = new Object [Columns][Rows]
    for (a=0;a<Columns;a++)
        {
        for (b=0;b<Rows;b++)
            {
                baseline[a][b]= sheet.getCell(a,b).getContents();
            }
        }
}

def updateOutput(writableWorkbook,sheetName,start_row,rows,columnNo,result,resultTag){
    
        Sheet[] sheet= writableWorkbook.getSheets();
        ws = writableWorkbook.getSheet(sheetName);
        if(ws==null){
        ws = writableWorkbook.createSheet(sheetName,sheet.length)
        }
        for (a=0;a<columnNo;a++)
        {
            ws.addCell(new jxl.write.Label(a,0,result[a][0]));
            for (b=start_row;b<rows;b++)
            {
                wc = new jxl.write.WritableCellFormat();
                if (resultTag[a][b]=='PASS')            
                {
                    wc.setBackground(jxl.format.Colour.WHITE);
                    ws.addCell(new jxl.write.Label(a,b,result[a][b],wc));
                }
                else if(resultTag[a][b]=='FAIL'){
                    wc.setBackground(jxl.format.Colour.RED);
                    ws.addCell(new jxl.write.Label(a,b,result[a][b],wc));
                }
            }
        }
}

def updateResult(writableWorkbook,sheetName,start_row,rows,columnNo,result){
    
        Sheet[] sheet= writableWorkbook.getSheets();
        ws = writableWorkbook.getSheet(sheetName);
        if(ws==null){
        ws = writableWorkbook.createSheet(sheetName,sheet.length)
        }
        for (a=0;a<columnNo;a++)
        {
            ws.addCell(new jxl.write.Label(a,0,result[a][0]));
            for (b=start_row;b<rows;b++)
            {
               wcf = new jxl.write.WritableCellFormat();
               wcf.setBackground(jxl.format.Colour.RED);
               
                  if(result[a][b] == 'FAIL'){
                       ws.addCell(new jxl.write.Label(a,b,result[a][b],wcf));
                  }    else             
                    ws.addCell(new jxl.write.Label(a,b,result[a][b]));
            }
        }

        wc = new jxl.write.WritableCellFormat();
          wcc = new jxl.write.WritableCellFormat();
          wc.setBackground(jxl.format.Colour.GRAY_25);
          wcc.setBackground(jxl.format.Colour.LIGHT_TURQUOISE);
          ws.addCell(new jxl.write.Label(0,rows-3,result[0][rows-3],wcc));   
          ws.addCell(new jxl.write.Label(0,rows-2,result[0][rows-2],wc));
        ws.addCell(new jxl.write.Label(0,rows-1,result[0][rows-1],wc));
        
}

def updateComparison(writableWorkbook,sheetName,start_row,rows,columnNo,output,outputTag,result,baseline){
    
        Sheet[] sheet= writableWorkbook.getSheets();
        ws = writableWorkbook.getSheet(sheetName);
        if(ws==null){
        ws = writableWorkbook.createSheet(sheetName,sheet.length)
        }
        for (a=0;a<columnNo;a++)
        {
            ws.addCell(new jxl.write.Label(a,0,output[a][0]));
            
            x= 1;
            for (b=start_row;b<rows;b++)
            {
                if(result[1][b] == 'FAIL'){
                  wc = new jxl.write.WritableCellFormat();
                  if (outputTag[a][b]=='PASS')            
                  {
                        wc.setBackground(jxl.format.Colour.WHITE);
                     ws.addCell(new jxl.write.Label(a,x,output[a][b],wc));                     
                  }
                  else {
                     wc.setBackground(jxl.format.Colour.RED);
                     ws.addCell(new jxl.write.Label(a,x,output[a][b],wc));
                  }
                wbc = new jxl.write.WritableCellFormat();
                 wbc.setBackground(jxl.format.Colour.ICE_BLUE);
                ws.addCell(new jxl.write.Label(a,x+1,baseline[a][b],wbc));
                x+=2; 
              }
            }
        }
}

def removeSheetByName(writableWorkbook, name){
    
        Sheet[] sheet= writableWorkbook.getSheets();    
        ws = writableWorkbook.getSheet(name);
        if(ws != null){
            for (i = 0; i < sheet.length; i++) {
                sheetName = writableWorkbook.getSheet(i).getName();            
                if (sheetName.equalsIgnoreCase(name)) {
                    writableWorkbook.removeSheet(i);
                }
            }
        }
}

def setProperties(Name,Value,Place)
{    
    name = Name;
    target = testRunner.testCase.getTestStepByName(Place);
    target.setPropertyValue(name,Value);
    }

def cleanProperty(PropertyListName)
{
         PropertyList = testRunner.testCase.getTestStepByName(PropertyListName);
         size=PropertyList.getPropertyCount();
         if (size!=0)
         {
                   for (i=0;i<size;i++)
                   {
                            PropertyList.removeProperty(PropertyList.getPropertyAt(0).name);
                   }
         }
}

def xlsName = context.expand('${#TestCase#Workbook}');

def project = testRunner.testCase.getTestSuite().getProject();
def testSuite = testRunner.testCase.getTestSuite();
def testcase = testRunner.testCase

inputSheetName = "Input";

def cal = Calendar.instance;
def sysdate = cal.getTime();
sleepTime=context.expand('${#Project#sleepTime}').toInteger()

try{    
       WorkbookSettings setting=new WorkbookSettings();
       setting.setEncoding("iso-8859-1"); 
       Workbook workbook=Workbook.getWorkbook(new File(xlsName),setting);
       cleanProperty(inputSheetName);
       readInput(workbook,inputSheetName);
      
    for (i=0;i<columns;i++)
    {
      setProperties(input[i][0],'',inputSheetName)
    }
       workbook.close();
       
}catch(Exception e){
      e.printStackTrace();
}
        
/*-----------Set runSelect as 'true' in Project property to run selected test cases from startTag till endTag-------*/
def runSelected = context.expand('${#Project#runSelected}')
start_Test=1;
end_Test=rows-1;

if ('true'.equalsIgnoreCase(runSelected))
{
    startTag=context.expand('${#Project#startTag}').toInteger();
     endTag=context.expand('${#Project#endTag}').toInteger();
     
    if((0<startTag)&&(startTag<rows)&&(endTag>=startTag)&&(rows>endTag)){
    start_Test=startTag;
    end_Test=endTag;
    }
}

    for (m=start_Test;m<=end_Test;m++)
    {
        for (i=0;i<columns;i++)
        {
            setProperties(input[i][0],input[i][m],inputSheetName)
        }
        testRunner.runTestStepByName("Start");
        Thread.sleep(sleepTime);       
		testRunner.runTestStepByName("Request-xml");
		testRunner.runTestStepByName("Request-json");
}

