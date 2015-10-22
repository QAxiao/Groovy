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

def logFile = new File(context.expand('${#TestCase#LogFile - Check Response}'));
def xlsName = context.expand('${#TestCase#Workbook}');

def project = testRunner.testCase.getTestSuite().getProject();
def testSuite = testRunner.testCase.getTestSuite();
def testcase = testRunner.testCase

inputSheetName = "Input";
baselineSheet = "Baseline";
outputSheet = "Output";
resultSheet = "Result";
fieldResult = testcase.getTestStepByName('fieldResult');
ComparisonSheet = "Comparison";

Baseline = testcase.getTestStepByName(baselineSheet);
baselineSize = Baseline.getPropertyCount();
Output = testcase.getTestStepByName(outputSheet);

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

      cleanProperty(baselineSheet);
      readBaseline(workbook,baselineSheet);
    for (i=0;i<Columns;i++)
    {
      setProperties(baseline[i][0],'',baselineSheet)
    }
    //clean the property of fiedlresult could avoid comparison error in result excel
      cleanProperty("fieldResult")
       workbook.close();
       
}catch(Exception e){
      e.printStackTrace();
}
        
/*-----------Set runSelect as 'true' in Project property to run selected test cases from startTag till endTag-------*/
def runSelected = context.expand('${#Project#runSelected}')
start_Test=1;
end_Test=rows-1;

def passNumbers = 0;
def decFormat = new DecimalFormat("##.00%"); 

if ('true'.equalsIgnoreCase(runSelected))
{
    startTag=context.expand('${#Project#startTag}').toInteger();
     endTag=context.expand('${#Project#endTag}').toInteger();
     
    if((0<startTag)&&(startTag<rows)&&(endTag>=startTag)&&(rows>endTag)){
    start_Test=startTag;
    end_Test=endTag;
    }
}
    
    result = new Object [2][rows+3]
    result[0][0]='Case Description';
    result[1][0]='Result'
    result[0][rows+1]='Start Time:';
    result[1][rows+1]=sysdate.toString();
    
/*--New object and put the output name and value into this list--*/
    output = new Object [baselineSize][rows]
    outputTag = new Object [baselineSize][rows]
    for (i=0;i<baselineSize;i++)
    {
        output[i][0]= Baseline.getPropertyAt(i).name;
        outputTag[i][0]= 'PASS';
        }
        
    for (m=start_Test;m<=end_Test;m++)
    {
         logFile.append('\n'+ testcase.name + ": "+m+" "+ ". "+sysdate+'\n');
        for (i=0;i<columns;i++)
        {
            setProperties(input[i][0],input[i][m],inputSheetName)
        }
        for (j=0;j<Columns;j++)
        {    
             setProperties(baseline[j][0],baseline[j][m],baselineSheet)
        }
            
        testRunner.runTestStepByName("Request-xml");
        Thread.sleep(sleepTime);
        testRunner.runTestStepByName("Check Response");
                
        result[0][m]=context.expand('${'+inputSheetName+'#Case Description}')
        result[1][m]=context.expand('${'+resultSheet+'#result}')

        if(result[1][m] == 'PASS'){
                 passNumbers++;
        }

        for (i=0;i<baselineSize;i++)
        {
            output[i][m]= Output.getPropertyAt(i).value;
            outputTag[i][m]= fieldResult.getPropertyAt(i).value;
        }
        
}
          result[0][rows+2]='End Time:';
          result[1][rows+2]=sysdate.toString();
          result[0][rows]='Pass Percentage:';
          
          passPercentage = decFormat.format(passNumbers/(end_Test-start_Test+1));
          result[1][rows] = passPercentage

/*--------------Update Output, Result, Comparison sheet---------*/
    try{
        WorkbookSettings setting=new WorkbookSettings();
	setting.setEncoding("iso-8859-1"); 
	Workbook workbook=Workbook.getWorkbook(new File(xlsName),setting);
        writableWorkbook =  Workbook.createWorkbook(new File(xlsName), workbook);

        updateOutput(writableWorkbook,outputSheet,start_Test,end_Test+1,baselineSize,output,outputTag);
        updateResult(writableWorkbook,resultSheet,start_Test,rows+3,2,result);

        removeSheetByName(writableWorkbook,ComparisonSheet);
                  
        if(passPercentage != '100.00%'){
               updateComparison(writableWorkbook,ComparisonSheet,start_Test,end_Test+1,baselineSize,output,outputTag,result,baseline);
          }

          writableWorkbook.write();
        writableWorkbook.close();   
        workbook.close();
        
        }catch(Exception e){
            e.printStackTrace();
        }
           
          setProperties('passPercentage', passPercentage ,'Result');
          
        testRunner.gotoStepByName('End');
        
