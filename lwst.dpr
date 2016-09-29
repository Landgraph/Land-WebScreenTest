program lwst;

{$R 'resources.res' 'resources.rc'}

uses
  Forms,
  Main in 'Main.pas' {MainForm};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Land WebScreenTest';
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
