program NF10EX100;

uses
  Vcl.Forms,
  MainForm in 'MainForm.pas' {FormMain},
  UDataModule in 'UDataModule.pas' {DataModuleCIMT: TDataModule},
  SettingForm in 'SettingForm.pas' {FormSetting},
  DetailForm in 'DetailForm.pas' {FormDetail};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFormMain, FormMain);
  Application.CreateForm(TDataModuleCIMT, DataModuleCIMT);
  Application.CreateForm(TFormSetting, FormSetting);
  Application.CreateForm(TFormDetail, FormDetail);
  Application.Run;
end.
