unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  IdHTTP, StdCtrls, OleCtrls, SHDocVw, ComCtrls, StrUtils, ComObj, ActiveX, Vcl.Grids,
  Xml.xmldom, Xml.XMLIntf, Xml.Win.msxmldom, Xml.XMLDoc;

type
  TForm1 = class(TForm)
    edURl: TEdit;
    wbWebBrowser: TWebBrowser;
    btn1: TButton;
    btnRun: TButton;
    btnExec2: TButton;
    btnColStr: TButton;
    Save: TButton;
    Label1: TLabel;
    procedure btn1Click(Sender: TObject);
    procedure btnRunClick(Sender: TObject);
    procedure btnExec2Click(Sender: TObject);
    procedure btnColStrClick(Sender: TObject);
    procedure edURlKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormCreate(Sender: TObject);
    procedure wbWebBrowserDocumentComplete(ASender: TObject; const pDisp: IDispatch; const URL: OleVariant);
    procedure SaveClick(Sender: TObject);


  private
    { Private declarations }
  public
    { Public declarations }

  end;



var
  Form1: TForm1;
  Link,LinkCounter : String;
  LC2, EndBlockCounter : Integer;
  TotalInfo : TStrings;
  Links : TStrings;
  Start : Boolean;
  TempInfo : TStrings;
  CMS: Integer; // CountMaxString


implementation


{$R *.dfm}

procedure MsgBox(Capt, Msg: string);
begin
  MessageBox(0, PChar(Msg), PChar(Capt), mb_Right);
end;


/// список автоматических процедур (пользователь не может их вызвать)
procedure RecLinks(LinksNav: TStrings);   //создал процедуру отдельно от кнопки, для загрузки в мемо линки.
var cn,i:integer;
begin
 // Form1.mmResult.Lines.Clear;
  LinksNav.Clear;
  cn:=Form1.wbWebBrowser.OleObject.Document.Links.length;
  for i:=0 to cn-1 do LinksNav.Add(Form1.wbWebBrowser.OleObject.Document.Links.Item(i));
end;

// проц. для обрезки текста, в s - идет значение, после которого нужно сохранить значение.

procedure Rezka(s: string; Records: TStrings); 
var
 GC, i, j, tabcount:integer;
 rec: TStrings;
 str,CR,link: string;
begin
   GC:=Records.Count-1;
   tabcount:=0;
   link:='';
    for i:=0 to GC do  if s = Records.Strings[i] then  tabcount:=i;   // начинаю искать "table 1". и записываю его номер (строки) в список строк

    rec:= TStringList.Create();          // инициализирую новый список строк

    for j:=tabcount to GC do rec.Add(Records.Strings[j]);    // копирую нужные строки в новый список

    Records.Clear;       // ощищаю старый

    /// ищу какой журнал, чтоб убрать итоговую строку в конце, в калите оно есть, в других ее нету.
    link:=Copy(Form1.wbWebBrowser.OleObject.Document.Links.Item((Form1.wbWebBrowser.OleObject.Document.Links.length-15)), 36,6);

    if link = 'kalita'then
    for j:=0 to (rec.count-9) do Records.Add(rec.Strings[j])  // записываю старый заново, убираю 10 строк снизу (сальдо).
    else
    for j:=0 to (rec.count-1) do Records.Add(rec.Strings[j]);

     // ищем макс. кол записей (Запись 1 ... Запись Н).
    GC:=Records.Count-1;
    str:='Запись ';
    CR:='1';
    CMS:=1;
    for i:=0 to GC do
    begin
      if Concat(str, CR) = Records.Strings[i] then
      begin
      CMS:=CMS+1; // max. кол записей в тексте.
      end;
      Delete(str,Length(str),(Length(CR)-1));
      Delete(CR,1,Length(CR));
      CR:=IntToStr(CMS);
    end;

    rec.free;    // удаляю новый список.
end;


procedure ColStr(s: string; countRecords: integer; Records,FutureRecords: TStrings);
var
 GC, i, j, tabcount, aftercount:integer;
 rec: TStrings;
 CR : string;
begin
  GC:=Records.Count-1;
  tabcount:=0;
  aftercount:=0;
  CR:= IntToStr(countRecords);

  for i:=0 to GC do
  begin
  if Concat(s,CR) = Records.Strings[i] then
    begin
    tabcount:=i;
    countRecords:=countRecords+1;
    end;
  end;
    if countRecords > 10 then
    Delete(s,Length(s),(Length(CR)-2))
    else
    Delete(s,Length(s),(Length(CR)-1));

    Delete(CR,1,Length(CR));
    CR:=IntToStr(countRecords);
    for i:=0 to GC do
   begin
   if Concat(s,CR) = Records.Strings[i] then
     begin
     aftercount:=i;   // узнаю размер текстового блока. (иначе - конец текста)
     end;
   end;
      Delete(s,Length(s),(Length(CR)-1));
      Delete(CR,1,Length(CR));

   rec:= TStringList.Create();          // инициализирую новый список строк

     for j:=tabcount to aftercount do rec.Add(Records.Strings[j]);
     for j:=0 to (rec.count-2) do FutureRecords.Add(rec.Strings[j]);   // записываю старый заново

   rec.free;    // удаляю новый список из оперативки.
end;



procedure ColStr3(Records,FutureRecords: TStrings);
var
 GC, i, j1, x :Integer;
 rec: TStrings;

begin
  rec:= TStringList.Create();          // инициализирую новый список строк
  GC:=Records.Count;
  x:=GC;
  i:=1;
  
  while i < x do
  begin
     if 'Данные 6' = Records.Strings[i] then
     begin
       i:=i+1;
                                             //MsgBox('dannie', '6'); // сверка
          repeat
          rec.Add(Records.Strings[i]);
          i:=i+1;
          until 'Данные 7' = Records.Strings[i];
       i:=i+1;
                                             //MsgBox('dannie', '7');
       rec.Add(Records.Strings[i]);
       rec.Add('EndBlock');
       If not (i=(x-1))
       then
       i:=i+2;

     end;

     if AnsiLeftStr(Records.Strings[i],6) = 'Данные' then
         begin
                                              //    MsgBox('dannie', IntToStr(i));
         i:=i+1;
         rec.Add(Records.Strings[i]);
         end;
    i:=i+1;
  end;


 if rec.Count > 1 then
 for j1:=0 to (rec.count-1) do FutureRecords.Add(rec.Strings[j1]);
 
  rec.free;    // удаляю новый список из оперативки.

end;

// процедура для парсинга строк.

procedure ParsString(fs: String;  a: string; a1: string; countmax: Integer; FutureRecords: TStrings); 
var
 y,count,count2: Integer;
 s: string;
begin
 for y:=0 to countmax do
 begin
 count := PosEx(a,fs,(pos(a1,fs)));
 count2 := PosEx(a1,fs,count);
   if count2 = 0 then
      begin
      FutureRecords.Add(Trim(Copy(fs,(count+1),length(fs))));
      end
      else
      begin

      s:=Trim(Copy(fs,(count+1),(count2-(count+1))));
        if s > '' then
        begin
         FutureRecords.Add(Trim(Copy(fs,(count+1),(count2-(count+1)))));
        end;
      end;

   Delete(fs,1,(count2-1));
 end;

end;


procedure SLToArr(a: array of string; Records: TStrings);
var
    i : integer;
    rec: TStrings;
begin
         rec:= TStringList.Create();
         for i:=0 to (Records.Count-1) do rec.Add(Records.Strings[i]);
         for i:=0 to (rec.Count-1) do  a[i]:=rec.Strings[i];
         rec.free
end;

procedure SLToSL(a: TStrings; Records: TStrings);
var
    i : integer;
begin
    Records.Clear;
    for i:=0 to (a.Count-1) do Records.Add(a.Strings[i]);

end;

// для TotalInfo чтоб его не чистить.

procedure SLToSL2(a: TStrings; Records: TStrings);
var
    i : integer;
begin

    for i:=0 to (a.Count-1) do Records.Add(a.Strings[i]);

end;

procedure InMemo(a: TMemo; b: TStrings);
var
     j : integer;
begin
     a.Lines.Clear; //Clearing MEMO
     for j:=0 to b.Count-1 do  a.Lines.Add(b.Strings[j]);     // Check copy data
end;

procedure InStringList(a: TMemo; b: TStrings);  // Tmemo, Strings/
var
    j : integer;
begin
    for j:=0 to  (a.Lines.Count-1) do b.Add(a.Lines.Strings[j]);
end;


 // beetweenStrCount   функция возвращает количество подстрок в строке, на выбор  а или а1.
 
function BetweenStrCount (fs: string; a: string; a1: string) : Integer;
var
count,i: Integer;
begin
  count:=0;
  for i:=1 to length(fs) do
  begin
  if  fs[i] = a1 then   // если подэлемент строки fs равен a то
  count:=count+1;     //  счетчик + 1 
  end;
  Result:= count;
end;


// Работа над строками, получаю нужные данные и удаляю левые,
// кодом записи полученным ранее в ColStr.  (Запись 1 данные .. тдп).

procedure BlockRez (Records:TStrings; FutureRecords:TStrings);
var
  rec: TStrings;
  a,a1,fs,sif4:string;      // fs - for search
  i,y, imax, ymax, count, count2:Integer;
  if1,if2,if3,if4: Integer;
  if2s,if3s,if4s:string;

begin
  rec:= TStringList.Create();
  // перечень блоков между которыми в строке находятся нужные данные.
  a:='>';
  a1:='<';
  if2s:='<br>';
  if3s:='<div>';
  if4s:='&gt;';
  sif4:='';
  // конец перечня блока
  imax:= Records.Count-1;

  // отдельно удаляю значок, который плохо сработал в кодировке отображения страниц в ИЕ. &gt;  '>'
  // попутно удалил 4 пробела (лишних) - или знак табуляции хз).
  // Зато теперь рабочее место и профессия разделены лишь " - ".
  // Проще искать, 2 знака до профы, а рабочее после того что останется убрать 3 знака из конца " - ".
  // Может пригодится.

  for i:=0 to imax do
  begin
  if4:=AnsiPos(if4s,Records.Strings[i]);
    if if4>0 then
    begin
    sif4:=Records.Strings[i];
    Delete(sif4,if4, 8);
    Records.Strings[i]:=sif4;
    end;
  end;


  for i:=0 to imax do
  begin
        if1:= AnsiPos(a1,Records.Strings[i]);
        if2:= AnsiPos(if2s,Records.Strings[i]);
        if3:= AnsiPos(if3s,Records.Strings[i]);

        if (if1>0) and (if2=0) and (if3=0) then     //for defalut (1 records per strting)
        begin
             fs:=Records.Strings[i];
                count := PosEx(a,fs,(pos(a1,fs)));
                count2 := PosEx(a1,fs,count);
                fs:=Copy(fs,(count+1),(count2-(count+1)));
             FutureRecords.Add(Trim(fs));
        end;
        if (if1=0) and (if2=0) and (if3=0) then
        begin
            FutureRecords.Add(Records.Strings[i]);
        end;

        if (if3>0) or (if2>0) then       //for commets
        begin
            fs:=Records.Strings[i];
            ymax:=(BetweenStrCount(fs,a,a1)-1);
            ParsString(fs,a,a1,ymax,rec);
            for y:= 0 to rec.count-1 do
            begin
              FutureRecords.Add(rec.Strings[y]);
            end;
        rec.clear;
        end;
  end;
  rec.Free;
end;

procedure ArrayDateClick(Records:TStrings; DataArray: Array of String);
var
 FutureRecords: TStrings;
begin
  FutureRecords:= TStringList.Create();
  ColStr3(Records,FutureRecords);
  SLToArr(DataArray,FutureRecords);
  FutureRecords.Free;
end;


// Событийные процедуры, которыми управляет пользователь

procedure TForm1.btn1Click(Sender: TObject);
var cn,cnrows,cncols, i,j, k:integer; vTags:OleVariant; s:string;
       Node,RowNode, ColNode:TTreeNode; vCols:OleVariant;
begin
      TempInfo.Clear;
      vTags:=wbWebBrowser.OleObject.Document.getElementsByTagName('Table');
      cn:=vTags.length;

      //нужно перебрать все строки и столбцы таблицы
      for i:=0 to cn-1 do
      begin
         s:='';
         if vTags.Item(i).id<>'' then s:=s+' id='+vTags.Item(i).id;
         TempInfo.Add('Table '+IntToStr(i)); // добавляю "таблицю"
         cnrows:=vTags.Item(i).rows.length;
         for j:=0 to cnrows-1 do
         begin
             vCols:=vTags.Item(i).rows.item(j).getElementsByTagName('TD');
             TempInfo.Add('Запись '+IntToStr(j)); // добавляю "запись"
             cncols:=vCols.length;
             for k:=0 to cncols-1 do
             begin
                  TempInfo.Add('Данные '+IntToStr(k));
                  s:=vCols.Item(k).innerHTML;
                  TempInfo.Add(s);
             end;
          end;
       end;

end;


procedure TForm1.btnRunClick(Sender: TObject);
var i : Integer;
    LinksNav : TStrings;
    index: Integer;
    s,s1,s2: string;
    ch: Set of Char;
begin
   LinksNav := TStringList.Create();

   RecLinks(LinksNav);
   ch:=['0'..'9'];
   index:=0;
   s:= LinksNav.Strings[wbWebBrowser.OleObject.Document.Links.length-14];

   s1:= Copy(s,Length(s)-3,4);  //копирую в строку линк каунтер, значение 2 последних элементов строки 16, число макс. страниц в калите
   for i:=1 to length(s1) do
      if s1[i] in  ch then
        index:=index+1;

      if index > 0 then
       while Pos('/', s1) > 0 do
       begin
       Delete(s1,1,1);
       end
     else
      s1:='1';

   s2:= Copy(s,1,(Length(s)-index)); // желательно позже добавить проверку чтоб без цыфр:) на случай если будет переход не с первой страницы.
   LC2 := StrToInt(s1);   // макс кол страниц.

   for i:=2 to LC2 do
    begin
     // делаю массив ссылок по которым буду переходить :).
    Links.Add(s2+IntToStr(i));
    end;
    // запускаю автопереходы ;).
    Start:=True;
    wbWebBrowser.Navigate(s2);

  LinksNav.Free;
end;

//Функция работы окна сохранения, создаем тхт-пободный файл,
//с разделителем и сохраняем его в формате CSV - для дальнешего экспорта в электронные таблицы.

procedure TForm1.SaveClick(Sender: TObject);
var
 saveDialog1: Tsavedialog;
 CSV: Tstrings;
 i : Integer;
 str : string;
begin
 saveDialog1 := TSaveDialog.Create(self);;
 CSV := TstringList.Create();

 CSV.Add(';Операции калиты: ');
 CSV.Add('');

 str:= '';
 for i := 1 to TotalInfo.Count do
   begin
       if  TotalInfo[i-1]='EndBlock'then
          begin
          CSV.Add(str);
          str:='';
          end
           else
           begin
           str:=str+TotalInfo[i-1]+';';   // ";" - разделитель строк. 
           end;
 end;

 SaveDialog1.Filter:='CSV files (*.csv)|*.csv|Text file|*.txt|All files| *.*';
 SaveDialog1.FileName := 'Report';
 if SaveDialog1.Execute then begin
     CSV.SaveToFile
        (SaveDialog1.FileName + '.csv');

 end;

 saveDialog1.Free;
 CSV.Free;
 TotalInfo.Free;
 save.Enabled:=False;
end;

procedure TForm1.btnExec2Click(Sender: TObject);
Var
    s:string;
    Records:TStrings;
    //Records2: array of string;
begin
   s:= 'Запись 1';
   Records:=TStringList.Create();   

   SlToSl(TempInfo,Records); 
   Rezka(s,Records);          //Переписываем лист - со с искомой строки, если такая там есть.
   SLToSl(Records,TempInfo);  //выводим результат на просмотр :) сначала Мемо, потом Тстринг.

   Records.Free;
end;

procedure TForm1.btnColStrClick(Sender: TObject);   // для второго этапа резки (чтоб программу не перезапускать много раз).
var
  FutureRecords, Records, FutureRecords1, FutureRecords2 : TStrings;
  s,CR:string;
  i: Integer;

begin
  s:= 'Запись ';
  CR:='';
  FutureRecords:= TStringList.Create();
  Records:= TStringList.Create();
  FutureRecords1:= TStringList.Create();
  FutureRecords2:= TStringList.Create();
  Records.Clear;


  SlToSl(TempInfo,Records);
  for i:=1 to (CMS+1) do             // цыкл перебирает все записи в массиве.
  begin
    ColStr(s,i,Records,FutureRecords);  // процедура выборки записи,
    //(строка (шаблон), номер записи, из какого списска, в какой).
  end;
    BlockRez(FutureRecords,FutureRecords1);   // добавил блок перевода в Данные :).
    // этап 2 ... пройден! (блоки [кода] подключались поэтапно с проверкой работоспобности, в 3 этапа.)

   ColStr3(FutureRecords1,FutureRecords2);
    // этап 3  ... finish~!!!

  SlToSL(FutureRecords2,TempInfo);
  SlToSL2(TempInfo,TotalInfo);

  FutureRecords2.Free;
  FutureRecords1.Free;
  FutureRecords.Free;
  Records.Free;
end;


procedure TForm1.wbWebBrowserDocumentComplete(ASender: TObject;
  const pDisp: IDispatch; const URL: OleVariant);
begin
if Start then
  Begin
  try
            btn1.Enabled:=True;
              btnExec2.Enabled:=True;
               btnColStr.Enabled:=True;

   if Links.Count > 0 then
    begin
      wbWebBrowser.Navigate(Links.Strings[0]);
             btn1.Click;
              btnExec2.Click;
               btnColStr.Click;
      form1.edURL.Text:=wbWebBrowser.LocationURL;
     Links.Delete(0);
    end
    else
     begin
        wbWebBrowser.Stop;
             btn1.Click;
              btnExec2.Click;
               btnColStr.Click;
                Start:=False;
                      save.Enabled:=True;
                       btnRun.Enabled:=False;
              btn1.Enabled:=False;
              btnExec2.Enabled:=False;
               btnColStr.Enabled:=False;
               TempInfo.Free;
        form1.edURL.Text:=wbWebBrowser.LocationURL;
      end;
except
    wbWebBrowser.Stop;
    form1.edURL.Text:=wbWebBrowser.LocationURL
end;

 End;
form1.edURL.Text:=wbWebBrowser.LocationURL
end;

// для работы клаваиши "Enter"  в навигации браузера.

procedure TForm1.edURlKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if Key = VK_RETURN then
wbWebBrowser.Navigate(edURL.Text);
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
Start:=False;
TotalInfo:=TStringList.Create();
Links:=TStringList.Create();
wbWebBrowser.Navigate(edURL.Text);
TempInfo:=TStringList.Create();
              btn1.Enabled:=False;
              btnExec2.Enabled:=False;
              btnColStr.Enabled:=False;
              btn1.Visible:=False;
              btnExec2.Visible:=False;
              btnColStr.Visible:=False;


end;

end.
