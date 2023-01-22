# UnitOffice.dcu
### 
public
    { Public declarations }

    stGrid:string[30];
    stGrid_Light:string[30];
    stGrid_Simple_1:string[30];
    stGrid_Simple_2:string[30];
    stGrid_Simple_3:string[30];
    stGridTable_4:string[30];
    wdOrientPortrait:byte;
    wdOrientLandscape:byte;
    wdBorderTop:ShortInt;
    wdBorderLeft:ShortInt;
    wdBorderBottom:ShortInt;
    wdBorderRight:ShortInt;
    wdAlignParagraphLeft:byte;
    wdAlignParagraphCenter:byte;
    wdAlignParagraphRight:byte;

    ---------------------

    Function CreateWord:boolean;
    Function CreateExcel:boolean;
    Function VisibleWord(visible:boolean):boolean;
    Function AddDoc:boolean;
    Function AddTable(iRow:integer; iCount:integer; st:string):boolean;
    Function main_(v:byte):boolean;
    Function WordPageOrientacion(v:byte):boolean;
    Function AddTableRowBelow(n:byte):boolean;
    Procedure loadCaption(v1:byte);
    function TableCount:integer;
    Function TableColumsWidth(iTables:byte; iColum:byte; iColumWidth:byte):boolean;
    Procedure SetWordTablesCellValue(Table,Row,Column:integer;S:string);
    Function SetFormattedColumns (colRec:tab;iRow:byte;iCol:byte):string;
    Procedure ClearTabRec;
    Procedure WordTablesCellValue(const vWord: Variant; const Table, Row, Column: Integer;
       const Value: string;  const FontName: string; const FontSize:byte; const FontBold, FontItalic: boolean;
       const FontUnderLine: byte; const iColor:TColor; const iAlignment:byte);
    Function VisibleExcel(visible:boolean):boolean;
    Procedure PrintCellExcel(const iX, iY:integer; iStr:String; iVal:TLines; const iColor:TColor);
    Procedure SetWidthCellExcel(const iStr:string; const iWidth:integer);
    Procedure SetColorTextExcel(const iX, iY:integer; const iColor:TColor);
    Procedure MergeCellsExcel(const iSt:string);
    Procedure BorderCellsExcel(const iSt:string);
    function  CheckExcelInstalled(AValue:string):boolean; 

### Пример:
  formOffice.CheckExcelInstalled('Excel.Application') - проверяет установлен ли Microsoft Office Excel

### formOffice.main_(0) - работа будет с Excel 
formOffice.VisibleExcel(True);
+ ![Загрузка Excel](/images/excel.jpg)

### formOffice.main_(1) - работа будет с Word
formOffice.VisibleWord(True);
+ ![Загрузка Word](/images/word.jpg)

### Примеры:
formOffice.WordTablesCellValue(formOffice.mWord,2,i,1,'Пример','Arial',12,False,False,0,clBlack,1);

FormOffice.SetWidthCellExcel('A:A',4);