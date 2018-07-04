unit SynchroSx;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, HSysMenu, HTB97, StdCtrls, Mask, Hctrls, ExtCtrls, PgiEnv, inifiles,
  HEnt1, HPanel, UIUtil, LicUtil, PGIExec;

type
  TFSynchro = class(TFVierge)
    TRANSFERTVERS: TRadioGroup;
    TDATEECR: THLabel;
    DATEECR: THCritMaskEdit;
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
    function LanceTraitement(LigneDeCommande: string): boolean;
  public
    { Déclarations publiques }
  end;


Procedure LanceSynchroSxPCL(bForceModal: Boolean=False);

implementation

uses EntDP;

{$R *.DFM}


Procedure LanceSynchroSxPCL(bForceModal: Boolean=False);
var
PP        : THPanel;
FSynchro  : TFSynchro;
begin
     FSynchro := TFSynchro.Create(Application);

     PP:=FindInsidePanel ;
     if (bForceModal) or (PP=Nil) then
     begin
            try
              FSynchro.ShowModal;
            finally
              FSynchro.Free;
            end;
    end
    else
    begin
           InitInside(FSynchro,PP) ;
           FSynchro.Show ;
    end;
end;


function TFSynchro.LanceTraitement(LigneDeCommande: string): boolean;
var TSI: TStartupInfo;
  TPI: TProcessInformation;
begin
  Result := True;
  FillChar(TSI, SizeOf(TStartupInfo), 0);
  TSI.cb := SizeOf(TStartupInfo);
  if CreateProcess(nil, PCHAR(LigneDeCommande), nil, nil, False, CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS, nil, nil, TSI, TPI) then
  begin
    while WaitForSingleObject(Tpi.hProcess, 1000) <> WAIT_OBJECT_0 do
    begin
      if WaitForSingleObject(Tpi.hProcess, 1000) = WAIT_OBJECT_0 then
        break;
      Application.ProcessMessages;
    end;
    CloseHandle(TPI.hProcess);
    CloseHandle(TPI.hThread);
  end else
  begin
    Result := False;
  end;
end;


procedure TFSynchro.BValiderClick(Sender: TObject);
var
Fileini                 : string;
FicIni                  : TIniFile;
LigneDeCommande         : string;
Action                  : string;
begin
  inherited;
  if PGISrv=Nil then exit;

      // envoyer  Création par la COMSX, le fichier dossier.TRA
      // et remonter par S1, le fichier TRA
      if TRANSFERTVERS.Itemindex = 0 then
      begin
            // Paramètres complémentaires pour l'exe
            VH_DP.ExeParam := '/TRF=EXPORT;'
              + V_PGI_Env.NoDossier + ';'+V_PGI.UserLogin+';;;X;S1;SYN;' + V_PGI_Env.PathDos+'\'
              + V_PGI_Env.NoDossier+'.TRA;;01/01/1900;01/01/2099;;;Fichierrapport.txt'
              + 'DATEARRET='+ DATEECR.text;
            PGISrv.Execute('COMSX.EXE');
            While VH_DP.ExeEnCours do Application.ProcessMessages;
            VH_DP.ExeParam := '';

            Action := 'IMPORT';
            FileIni := V_PGI_Env.PathDos+'\'+ 'COMSX.INI';
            FicIni        := TIniFile.Create(Fileini);
            FicIni.WriteString (Action, 'expert', 'PGI');
            FicIni.WriteString (Action, 'type', 'SYN');
            FicIni.WriteString (Action, 'fichier', V_PGI_Env.PathDos+'\'+V_PGI_Env.NoDossier+'.TRA');
            FicIni.WriteString (Action, 'dateArrete', DateEcr.Text);
            FicIni.WriteString (Action, 'log',  V_PGI_Env.PathDos+'\'+'Rapport.txt');
            FicIni.free;
            LigneDeCommande :=
              V_PGI_Env.PathSynchro
             + '\S1.exe IMPORT /u '
             + V_PGI.UserLogin + ' /p '+ DeCryptageSt(V_PGI.PassWord) +  ' /s '
             + V_PGI_Env.NoDossier + ' /ini '+ V_PGI_Env.PathDos+'\'+'COMSX.INI';
            LanceTraitement (LigneDeCommande);
            DeleteFile(FileIni);
      end
      else  // recevoir création par S1 le fichier TRA puis remonté par la COMSX dans la compta
      if TRANSFERTVERS.Itemindex = 1 then
      begin

            LigneDeCommande :=
            V_PGI_Env.PathSynchro
             + '\S1.exe EXPORT /u '
             + V_PGI.UserLogin + ' /p '+ DeCryptageSt(V_PGI.PassWord) +  ' /s '
             + V_PGI_Env.NoDossier + ' /ini '+ V_PGI_Env.PathDos+'\'+'EXPCOMSX.INI';

            FileIni := V_PGI_Env.PathDos+'\'+ 'EXPCOMSX.INI';
            FicIni        := TIniFile.Create(Fileini);
            FicIni.WriteString ('EXPORT', 'expert', 'PGI');
            FicIni.WriteString ('EXPORT', 'type', 'SYN');
            FicIni.WriteString ('EXPORT', 'fichier', V_PGI_Env.PathDos+'\'+V_PGI_Env.NoDossier+'.TRA');
            if DateEcr.Text = '  /  /    ' then
               FicIni.WriteString ('EXPORT', 'dateArrete', FormatDateTime('dd/mm/yyyy',Date))
            else
               FicIni.WriteString ('EXPORT', 'dateArrete', DateEcr.Text);
            FicIni.WriteString ('EXPORT', 'log', V_PGI_Env.PathDos+'\'+'Rapport.txt');
            FicIni.free;
            LanceTraitement (LigneDeCommande);

            // paramètres complémentaires pour l'exe
            VH_DP.ExeParam := '/TRF='+V_PGI_Env.PathDos+'\;'+V_PGI_Env.NoDossier+'.TRA;IMPORT;'
              + V_PGI_Env.NoDossier+';'+V_PGI.UserLogin;
           // Lance le traitement
           PGISrv.Execute('COMSX.EXE');
           While VH_DP.ExeEnCours do Application.ProcessMessages;
           VH_DP.ExeParam := '';
           DeleteFile(FileIni);

      end;

end;

procedure TFSynchro.FormShow(Sender: TObject);
begin
  inherited;
  DATEECR.Text :=  FormatDateTime('dd/mm/yyyy', now);
end;

end.
