unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, XPMan, OleCtrls, SHDocVw_EWB, EwbCore,
  EmbeddedWB, ComCtrls, Buttons, FileCtrl, pngimage, JPEG;

type
  TMainForm = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    ScrollBox1: TScrollBox;
    WB: TEmbeddedWB;
    XPManifest1: TXPManifest;
    Panel1: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    SizeCombo: TComboBox;
    BtnStop: TSpeedButton;
    BtnGo: TSpeedButton;
    WBStatus: TStatusBar;
    Splitter1: TSplitter;
    PreviewScroll: TScrollBox;
    PreviewImage: TImage;
    BtnPreviewGen: TButton;
    HeightBox: TCheckBox;
    TabSheet3: TTabSheet;
    Label3: TLabel;
    AddressBox: TComboBox;
    Panel2: TPanel;
    PreviewList: TFileListBox;
    Label4: TLabel;
    BrowserVersion: TLabel;
    MonitorImage: TImage;
    PreviewMode: TComboBox;
    ScreenImage: TImage;
    EngineVer: TComboBox;
    procedure SizeComboChange(Sender: TObject);
    procedure WBStatusTextChange(ASender: TObject; const Text: WideString);
    procedure WBProgressChange(ASender: TObject; Progress,
      ProgressMax: Integer);
    procedure BtnStopClick(Sender: TObject);
    procedure WBNavigateComplete2(ASender: TObject; const pDisp: IDispatch;
      var URL: OleVariant);
    procedure BtnGoClick(Sender: TObject);
    procedure BtnPreviewGenClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure WBDocumentComplete(ASender: TObject; const pDisp: IDispatch;
      var URL: OleVariant);
    procedure HeightBoxClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure PreviewListChange(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure asdChange(Sender: TObject);
    procedure PreviewModeChange(Sender: TObject);
  private
    { Private declarations }
    CurrentPreview: TJPEGImage;
    CurrentPreviewFileName: String;
  public
    { Public declarations }
    procedure RecalcHeight;
    procedure SetComboSize;
    function IntFromCSS(Val:String):Integer;
    procedure SetPreviewMonitorSize;
  end;

const
  EngineVars: array[0..8] of integer = (
    11001,
    11000,
    10001,
    10000,
    9999,
    9000,
    8888,
    8000,
    7000
  );

var
  MainForm: TMainForm;
  AppPath: String;

implementation

uses Math, Registry;

{$R *.dfm}

procedure TMainForm.SizeComboChange(Sender: TObject);
var
  sW, sH, sNew: string;
  iW, iH, xPos: integer;
begin
  { Если пользователь выбирает последний пункт "Другое...", предлагаем
    ввести разрешение в определённом формате.
    Либо берём разрешение из списка.
    После этого разбиваем строку по символу х.
    Пытаемся преобразовать полученные части строки в числа.
    Если указанного разрешения в списке нет - добавляем его в конец списка.
  }

  if SizeCombo.ItemIndex = (SizeCombo.Items.Count-1) then
    begin
      sNew := InputBox('Новое разрешение', 'Введите разрешение в формате ШИРИНАхВЫСОТА (1024x768)', '');
      if sNew='' then exit;
      xPos := Pos('x', sNew);
      if xPos = -1 then
        begin
          MessageDlg('Разрешение указано в неверном формате!', mtError, [mbOk], 0);
          Exit;
        end;
      sW := Copy(sNew, 1, xPos-1);
      sH := Copy(sNew, xPos+1, Length(SizeCombo.Text));
    end
  else
    begin
      xPos := Pos('x', SizeCombo.Text);
      sW := Copy(SizeCombo.Text, 1, xPos-1);
      sH := Copy(SizeCombo.Text, xPos+1, Length(SizeCombo.Text));
    end;
  try
    iW := StrToInt(sW);
  except
    MessageDlg('Неправильно указана ширина!', mtError, [mbOk], 0);
    Exit;
  end;

  try
    iH := StrToInt(sH);
  except
    MessageDlg('Неправильно указана высота!', mtError, [mbOk], 0);
    Exit;
  end;

  sNew := IntToStr(iW)+'x'+IntToStr(iH);
  xPos := SizeCombo.Items.IndexOf(sNew);

  if xPos=-1 then
    begin
      SizeCombo.Items.Insert(SizeCombo.Items.Count-1, sNew);
      SizeCombo.ItemIndex := SizeCombo.Items.Count-1;
    end;

  WB.Width := iW;

  if not HeightBox.Checked then
    WB.Height := iH
  else
    RecalcHeight;

end;

procedure TMainForm.WBStatusTextChange(ASender: TObject;
  const Text: WideString);
begin
  WBStatus.SimpleText := Text;
end;

procedure TMainForm.WBProgressChange(ASender: TObject; Progress,
  ProgressMax: Integer);
begin
  if ProgressMax<>0 then
    WBStatus.SimpleText := 'Загрузка (' + IntToStr(Ceil(Progress/ProgressMax)) + ')...';
end;

procedure TMainForm.BtnStopClick(Sender: TObject);
begin
  WB.Stop;
end;

procedure TMainForm.WBNavigateComplete2(ASender: TObject;
  const pDisp: IDispatch; var URL: OleVariant);
begin
  WBStatus.SimpleText := '';
end;

procedure TMainForm.BtnGoClick(Sender: TObject);
begin
  if AddressBox.Items.IndexOf(AddressBox.Text)=-1 then
    AddressBox.Items.Insert(0, AddressBox.Text);
  WB.Navigate(AddressBox.Text);
end;

procedure TMainForm.BtnPreviewGenClick(Sender: TObject);
var
  i, current: integer;
begin
  { Подготавливаем превью для всех разрешений списка }

  current := SizeCombo.ItemIndex;
  for i:=0 to SizeCombo.Items.Count - 2 do
    begin
      SizeCombo.ItemIndex := i;
      SetComboSize;
      if HeightBox.Checked then
        RecalcHeight;
      WB.GetJpegFromBrowser(AppPath + 'previews\' + SizeCombo.Text + '.jpg', WB.Height, WB.Width, WB.Height, WB.Width);
    end;

  SizeCombo.ItemIndex := current;
  SetComboSize;
  if HeightBox.Checked then
    RecalcHeight;
end;

procedure TMainForm.FormCreate(Sender: TObject);
var
  PNG: TPNGObject;
begin
  { Загружаем настройки из файлов }

  AppPath := ExtractFilePath(Application.ExeName);
  if FileExists(AppPath+'ssizes.txt') then
    begin
      SizeCombo.Items.LoadFromFile(AppPath + 'ssizes.txt');
      SizeCombo.Items.Add('Другой...');
      SizeCombo.ItemIndex := 0;
    end;
  if FileExists(AppPath+'history.txt') then
    begin
      AddressBox.Items.LoadFromFile(AppPath + 'history.txt');
    end;
  if not FileExists(AppPath+'previews') then
    CreateDir(AppPath + 'previews');

  { Загружаем "монитор" из ресурса }

  PNG := TPNGObject.Create;
  try
    PNG.LoadFromResourceName(HInstance, 'MONITORIMAGE');
    MonitorImage.Picture.Assign(PNG);
  finally
    PNG.Free;
  end;

  { Начальная инициализация }
  CurrentPreview := TJPEGImage.Create;
  SetPreviewMonitorSize;
  MainForm.DoubleBuffered := true;
  WB.Navigate('about:blank');
end;

procedure TMainForm.RecalcHeight;
var
  iH: Integer;
  Doc: Variant;
begin
  { Выбираем максимальную высоту, включая внешние отступы объектов
    body и documentElement }

  if not HeightBox.Checked then exit;
  if not WB.DocumentLoaded then exit;
  Doc := WB.Doc5;
  try
    iH := Doc.body.scrollHeight + IntFromCSS(Doc.body.style.marginTop) + IntFromCSS(Doc.body.style.marginBottom);
    iH := Max(iH, Doc.documentElement.scrollHeight + IntFromCSS(Doc.documentElement.style.marginTop) + IntFromCSS(Doc.documentElement.style.marginBottom));
  except
    Exit;
  end;

  { Добавляем немного высоты, чтобы не появлялась вертикальная прокрутка,
    если появилась горизонтальная. }
  WB.Height := iH+20;
end;

procedure TMainForm.WBDocumentComplete(ASender: TObject;
  const pDisp: IDispatch; var URL: OleVariant);
begin
  //После загрузки документа высчитываем высоту окна.
  RecalcHeight;
end;

procedure TMainForm.SetComboSize;
var
  sW, sH: string;
  iW, iH, xPos: integer;
begin
  //Переустановка размеров окна согласно выбранного разрешения

  xPos := Pos('x', SizeCombo.Text);
  sW := Copy(SizeCombo.Text, 1, xPos-1);
  sH := Copy(SizeCombo.Text, xPos+1, Length(SizeCombo.Text));

  try
    iW := StrToInt(sW);
  except
    MessageDlg('Неправильно указана ширина!', mtError, [mbOk], 0);
    Exit;
  end;

  try
    iH := StrToInt(sH);
  except
    MessageDlg('Неправильно указана высота!', mtError, [mbOk], 0);
    Exit;
  end;

  WB.Width := iW;
  WB.Height := iH;
end;

procedure TMainForm.HeightBoxClick(Sender: TObject);
begin
  { Флажок меняет высоту с фиксированной на резиновую и наоборот }
  if HeightBox.Checked then
    RecalcHeight
  else
    SetComboSize;
end;

function TMainForm.IntFromCSS(Val: String): Integer;
var
  res: string;
  i: integer;
begin
  { ПРИМИТИВНЫЙ преобразователь чисел из формата CSS в обычные целые числа.
    Не учитывает вариации сложнее px
   }
  res := '';
  for i:=1 to Length(Val) do
    begin
      if Val[i] in ['0'..'9'] then
        res := res + Val[i];
    end;
  if res = '' then
    Result := 0
  else
    try
      Result := StrToInt(res);
    except
      Result := 0;
    end;
end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  { Сохраняем настройки и освобождаем ресурсы }
  SizeCombo.Items.Delete(SizeCombo.Items.Count - 1);
  SizeCombo.Items.SaveToFile(AppPath + 'ssizes.txt');
  AddressBox.Items.SaveToFile(AppPath + 'history.txt');
  CurrentPreview.Free;
end;

procedure TMainForm.PreviewListChange(Sender: TObject);
begin
  PreviewModeChange(nil);
end;

procedure TMainForm.PageControl1Change(Sender: TObject);
var Doc:Variant;
begin
  if PageControl1.ActivePageIndex = 1 then
    begin
      //Обновляем список файлов в каталоге
      PreviewList.ApplyFilePath(AppPath + 'previews');
      PreviewList.Update;
    end
  else if PageControl1.ActivePageIndex = 2 then
    begin
      //Показываем на странице "О программе" useragent браузера
      try
        Doc := WB.Doc5;
        BrowserVersion.Caption := Doc.parentWindow.navigator.userAgent;
      except
        BrowserVersion.Caption := '';
      end;
    end;
end;

procedure TMainForm.FormResize(Sender: TObject);
begin
  //Пересчитываем вывод превью при ресайзе формы
  PreviewModeChange(nil);
end;

procedure TMainForm.asdChange(Sender: TObject);
var
  R:TRegistry;
begin
  { Устанавливаем версию IE с которой работает программа }

  R := TRegistry.Create;
  try
    R.RootKey := HKEY_CURRENT_USER;
    if R.OpenKey('\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION', True) then
      begin
        if EngineVer.ItemIndex>0 then
          R.WriteInteger(ExtractFileName(Application.ExeName), EngineVars[EngineVer.ItemIndex])
        else
          R.DeleteValue(ExtractFileName(Application.ExeName));
        R.CloseKey;
      end;
  finally
    R.Free;
    MessageDlg('Для применения настроек необходимо перезапустить приложение.', mtWarning, [mbOk], 0);
    WB.Free;
  end;
end;

procedure TMainForm.PreviewModeChange(Sender: TObject);
var
  bmp: TBitmap;
  bH: Real;
  UpdateFlag: Boolean;
begin
  UpdateFlag := false;
  try
    { Если выбор файла изменился, перечитываем его в память }
    if (PreviewList.Filename<>'') and (CurrentPreviewFileName<>PreviewList.FileName) then
      begin
        CurrentPreviewFileName := PreviewList.Filename;
        CurrentPreview.LoadFromFile(CurrentPreviewFileName);
        PreviewImage.Picture.Assign(CurrentPreview);
        UpdateFlag := true;
      end;
  except
  end;
  case PreviewMode.ItemIndex of
    1:
      begin
        MonitorImage.Visible := false;
        PreviewImage.AutoSize := false;
        ScreenImage.Visible := false;
        PreviewImage.Top := 0;
        PreviewImage.Left := 0;

        { Делаем запас в 25 пикселей чтобы не появлялась горизонтальная прокрутка }

        if CurrentPreview.Width>(PreviewScroll.Width-25) then
          begin
            PreviewImage.Width := PreviewScroll.Width - 25;
            PreviewImage.Height := Ceil((PreviewImage.Width-25) / PreviewScroll.Width * CurrentPreview.Height) - 20;
          end
        else
          begin
            PreviewImage.Width := CurrentPreview.Width;
            PreviewImage.Height := CurrentPreview.Height;
          end;
        PreviewImage.Stretch := true;
        PreviewImage.Proportional := true;
        PreviewImage.Visible := true;
      end;
    2:
      begin
        MonitorImage.Visible := false;
        ScreenImage.Visible := false;
        PreviewImage.Visible := true;
        PreviewImage.AutoSize := false;
        PreviewImage.Top := 0;
        PreviewImage.Left := 0;

        { Делаем запас чтобы не появлялись полосы прокрутки }
        PreviewImage.Width := PreviewScroll.Width-20;
        PreviewImage.Height := PreviewScroll.Height-20;
        PreviewImage.Stretch := true;
        PreviewImage.Proportional := true;
        PreviewImage.Visible := true;
      end;
    3:
      begin
        if MonitorImage.Visible=false then
          begin
            MonitorImage.Visible := true;
            PreviewImage.Visible := false;
          end;
        ScreenImage.Visible := false;
        //Пересчитываем размеры "Монитора" и "Экрана"
        SetPreviewMonitorSize;

        { Вписываем на экран превью }
        bmp := TBitmap.Create;
        bmp.Assign(CurrentPreview);
        
        //Вычисляем высоту вписываемой области (ширина известна)
        if ScreenImage.Width < CurrentPreview.Width then
          begin
            bH := CurrentPreview.Width * ( ScreenImage.Height / ScreenImage.Width );
          end
        else
          begin
            bH := CurrentPreview.Width * (ScreenImage.Height / ScreenImage.Width);
          end;

        SetStretchBltMode(ScreenImage.Canvas.Handle, HALFTONE);
        StretchBlt(ScreenImage.Canvas.Handle,0,0,ScreenImage.Width, ScreenImage.Height,
                   bmp.Canvas.Handle, 0, 0, CurrentPreview.Width, Ceil(bH),
                   SRCCOPY);
        bmp.Free;
        ScreenImage.Visible := true;
      end
    else
      begin
        ScreenImage.Visible := false;
        MonitorImage.Visible := false;
        PreviewImage.AutoSize := false;
        PreviewImage.Stretch := false;
        PreviewImage.Proportional := false;
        PreviewImage.Top := 0;
        PreviewImage.Left := 0;
        PreviewImage.Width := 0;
        PreviewImage.Height := 0;
        PreviewImage.AutoSize := true;
        PreviewImage.Visible := true;
      end;
  end;
end;

procedure TMainForm.SetPreviewMonitorSize;
var k:real;
begin
  if PreviewScroll.Width>PreviewScroll.Height then
    begin
      MonitorImage.Height := PreviewScroll.Height;
      MonitorImage.Width := Ceil(PreviewScroll.Height / MonitorImage.Picture.Height * MonitorImage.Width);
      k := MonitorImage.Height / MonitorImage.Picture.Height
    end
  else
    begin
      MonitorImage.Width := PreviewScroll.Width;
      MonitorImage.Height := Ceil(PreviewScroll.Width / MonitorImage.Picture.Width * MonitorImage.Height);
      k := MonitorImage.Width / MonitorImage.Picture.Width;
    end;
  ScreenImage.Width := Ceil(1330*k);
  ScreenImage.Height := Ceil(747*k);
  ScreenImage.Top := Ceil(47*k);
  ScreenImage.Left:= Ceil(8*k);
  ScreenImage.Picture.Bitmap.Width := ScreenImage.Width;
  ScreenImage.Picture.Bitmap.Height := ScreenImage.Height;
end;

end.
