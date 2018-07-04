{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFREVEXPORTINDICE ()
Mots clefs ... : TOF;AFREVEXPORTINDICE
*****************************************************************}
unit utofAfRevExportIndice;

interface

uses StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fe_Main,
  {$ELSE}
  MainEagl,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF, UTob, HTB97, buttons, paramsoc,
  FileCtrl, Shellapi, windows, dicobtp;

type
  TOF_AFREVEXPORTINDICE = class(TOF)
    BTNValider: TToolbarButton97;
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
    procedure BTNValiderClick(sender: Tobject);
    procedure Exporte;
  end;

procedure AFLanceFiche_AFREVEXPORTINDICE;
                                                   
implementation

procedure TOF_AFREVEXPORTINDICE.Exporte;
var MaTob: Tob;
  Q: Tquery;
  St: string;
  DateLimite: TdateTime;
  i : Integer;
begin

  MaTob := TOB.Create('LesValeursIndices', nil, -1);
  DateLimite := strtodate(getcontroltext('DATEEXP'));
  try
    st := 'select *  from afvalindice left join afindice on AFV_INDCODE=AIN_INDCODE ';
    st := st + ' where AFV_INDDATEVAL >"' + UsDateTime(DateLimite) + '"';
    st := st + 'order by AIN_INDCODE';

    Q := nil;
    try
      Q := OpenSQL(st, TRUE);
      MaTob.LoadDetailDB('', '', '', Q, false);

      // dates en string pour bien les récupérer
      for i := 0 to maTob.detail.count - 1 do
      begin
        MaTob.detail[i].PutValue('AFV_INDVALEUR', varianttosql(MaTob.detail[i].GetValue('AFV_INDVALEUR')));
        MaTob.detail[i].PutValue('AFV_COEFRACCORD', varianttosql(MaTob.detail[i].GetValue('AFV_COEFRACCORD')));
        MaTob.detail[i].PutValue('AFV_COEFPASSAGE', varianttosql(MaTob.detail[i].GetValue('AFV_COEFPASSAGE')));
        MaTob.detail[i].AddChampSupValeur('AIN_INDDATECREA_STRING', datetostr(matob.detail[i].getvalue('AIN_INDDATECREA')));
        MaTob.detail[i].AddChampSupValeur('AIN_INDDATEFIN_STRING', datetostr(matob.detail[i].getvalue('AIN_INDDATEFIN')));
        MaTob.detail[i].AddChampSupValeur('AIN_DATEMAJ_STRING', datetostr(matob.detail[i].getvalue('AIN_DATEMAJ')));
        MaTob.detail[i].AddChampSupValeur('AFV_INDDATEVAL_STRING', datetostr(matob.detail[i].getvalue('AFV_INDDATEVAL')));
        MaTob.detail[i].AddChampSupValeur('AFV_DATECREATION_STRING', datetostr(matob.detail[i].getvalue('AFV_DATECREATION')));
        MaTob.detail[i].AddChampSupValeur('AFV_DATEMODIF_STRING', datetostr(matob.detail[i].getvalue('AFV_DATEMODIF')));
        MaTob.detail[i].AddChampSupValeur('AFV_INDDATEPUB_STRING', datetostr(matob.detail[i].getvalue('AFV_INDDATEPUB')));
        MaTob.detail[i].AddChampSupValeur('AFV_INDDATEFIN_STRING', datetostr(matob.detail[i].getvalue('AFV_INDDATEFIN')));
        MaTob.detail[i].AddChampSupValeur('AFV_DATEMAJ_STRING', datetostr(matob.detail[i].getvalue('AFV_DATEMAJ')));
      end;

    finally
      Ferme(Q);
    end;
    MaTob.SaveToXMLFile(GetControlText('FICHIERIMP'), true, true);

    if PGIAskCancelAF ('Le fichier ' + GetControlText('FICHIERIMP') + ' a été Généré. Voulez - vous le consulter ?', '') = MrYes then
      ShellExecute (0, PCHAR('open'),PChar(GetControlText('FICHIERIMP')), Nil,Nil,SW_RESTORE);
  finally
    MaTob.free;
  end;
end;

procedure TOF_AFREVEXPORTINDICE.OnNew;
begin
  inherited;
end;

procedure TOF_AFREVEXPORTINDICE.OnDelete;
begin
  inherited;
end;

procedure TOF_AFREVEXPORTINDICE.OnUpdate;
begin
  inherited;
end;

procedure TOF_AFREVEXPORTINDICE.OnLoad;
begin
  inherited;
end;

procedure TOF_AFREVEXPORTINDICE.BTNValiderClick(sender: Tobject);
begin
  Exporte;
end;

procedure TOF_AFREVEXPORTINDICE.OnArgument(S: string);
var
  vStFile : String;

begin
  inherited;
  BTNValider := TToolbarButton97(Getcontrol('BVALIDER'));
  BTNValider.OnClick := BTNValiderClick;

  if GetParamSoc('SO_AFINDIMPEXP') <> '' then
  begin
    vStFile := GetParamSoc('SO_AFINDIMPEXP');
    if vStFile[length(vStFile)] <> '\' then
      vStFile := vStFile + '\';
    vStFile := vStFile + 'IND' + FormatDateTime('yyyymmdd', Date) + '.xml';

  end
  else
  begin
    vStFile := ExtractFileDrive(Application.ExeName) + '\IND' + FormatDateTime('yyyymmdd', Date) + '.xml';
  end;
  setControltext('FICHIERIMP', vStFile);
  SetFocusControl('FICHIERIMP');
  SetControlText('DATEEXP', datetostr(DebutDeMois(now)));
end;

procedure TOF_AFREVEXPORTINDICE.OnClose;
begin
  inherited;
end;

procedure TOF_AFREVEXPORTINDICE.OnDisplay();
begin
  inherited;
end;

procedure TOF_AFREVEXPORTINDICE.OnCancel();
begin
  inherited;
end;

procedure AFLanceFiche_AFREVEXPORTINDICE;
begin                                    
  AglLanceFiche('AFF', 'AFREVEXPORTINDICE', '', '', '');
end;

initialization
  registerclasses([TOF_AFREVEXPORTINDICE]);
end.
