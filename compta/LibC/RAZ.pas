{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/04/2003
Modifié le ... : 22/04/2003
Description .. : - CA - 22/04/2003 - FQ 10526 - Libellé exercice.
Suite ........ : - CA - 22/04/2003 - FQ 12259 - Contrôle durée d'exercice
Suite ........ : - TGA- 06/09/2006 - FQ 18765 - raz immomvtd
Suite ........ : - BTY- 06/11/07 - raz immoregfr des passages forfait à réel en Agricole
Mots clefs ... :
*****************************************************************}
unit RAZ;

interface

uses WinTypes, WinProcs, Classes, Graphics, Forms, Controls, Buttons,
  StdCtrls, ExtCtrls, Hctrls, DB,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  hmsgbox, hstatus, HSysMenu, ImportConf,
  Ent1, Hent1, SysUtils, Mask, ParamSoc, LicUtil;

procedure RAZSociete;

type
  TFRAZ = class(TForm)
    PBouton: TPanel;
    BValider: THBitBtn;
    HelpBtn: THBitBtn;
    BFerme: THBitBtn;
    Memo1: TMemo;
    Msg: THMsgBox;
    HMTrad: THSystemMenu;
    Panel1: TPanel;
    HLabel1: THLabel;
    FOui: TEdit;
    Deb: TMaskEdit;
    TDeb: TLabel;
    TFin: TLabel;
    Fin: TMaskEdit;
    BExo: THBitBtn;
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BExoClick(Sender: TObject);
    procedure DebExit(Sender: TObject);
    procedure FinExit(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
  private
  public
    { Public declarations }
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcGen,
  {$ENDIF MODENT1}
  UtilPgi,
  OuvreExo,
  uLibExercice;   // CControleDureeExercice

procedure RAZSociete;
var
  FRAZ: TFRAZ;
begin
  if not _BlocageMonoPoste(True) then
    Exit;
  FRAZ := TFRAZ.Create(Application);
  try
    if OKToutSeul then
    begin
      // VL,GCO 03/12/2003 FQ 11738 Superviseur et MDP du jour uniquement pour PGE
      if (ctxPCL in V_PGI.PGIContexte) or
        (not (ctxPCL in V_PGI.PGIContexte)) and (V_PGI.Superviseur = True) and
          (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
      begin
        if HShowMessage(FRAZ.Msg.Mess[5], '', '') = mrYes then
          if HShowMessage(FRAZ.Msg.Mess[6], '', '') = mrNo then
            FRAZ.ShowModal;
      end
      else
        HShowMessage(FRAZ.Msg.Mess[3], '', '')
    end
    else
      HShowMessage(FRAZ.Msg.Mess[0], '', '');

  finally
    FRAZ.Free;
    _DeblocageMonoPoste(True);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 31/03/2005
Modifié le ... : 29/03/2007
Description .. : LG - 31/03/2005 - supprression de la table de cutoff
Suite ........ : LG - 29/03/2007 - suprression des GCD
Mots clefs ... : 
*****************************************************************}
procedure TFRAZ.BValiderClick(Sender: TObject);
var
  Sql: string;
  DDeb, Dfin: TDateTime;
  EvtSup    : TEvtSuptot;
begin
  if FOui.Text <> TraduireMemoire(Msg.Mess[7]) then
  begin
    HShowMessage(Msg.Mess[1], '', '');
    Exit;
  end;
  PBouton.Enabled := False;
  Sql := '';
  InitMove(22, Msg.Mess[4]);
  {JP 04/08/05 : FQ 15592 : En mettant DDeb à 0, les champs ?_DATEDERNMVT étaient mis à 30/12/1899
                 et il n'était plus possible d'éditer les journaux et les comptes généraux et auxiliaires}
  DDeb := iDate1900;
  DFin := IDate2099;

  ExecuteSQL('DELETE FROM ANALYTIQ');
  MoveCur(False);

  ExecuteSQL('DELETE FROM CONTABON');
  MoveCur(False);

  ExecuteSQL('DELETE FROM CUMULS');
  MoveCur(False);

  ExecuteSQL('DELETE FROM BUDECR');
  MoveCur(False);

  ExecuteSQL('DELETE FROM ECRITURE');
  MoveCur(False);

  // Fiche 17025
  EvtSup := TEvtSuptot.Create;
  EvtSup.DegageDocumentGed ('');
  EvtSup.Free;

  ExecuteSQL('DELETE FROM ECRCOMPL');
  MoveCur(False);

  ExecuteSQL('DELETE FROM EDTLEGAL');
  MoveCur(False);

  ExecuteSQL('DELETE FROM EEXBQ');
  MoveCur(False);

  ExecuteSQL('DELETE FROM EEXBQLIG');
  MoveCur(False);

  ExecuteSQL('DELETE FROM HISTOBAL');
  MoveCur(False);

  {JP 09/08/06 : FQ 18668 : Gestion des BAP}
  ExecuteSQL('DELETE FROM CPTACHEBAP');
  MoveCur(False);

  {JP 09/08/06 : FQ 18668 : Gestion des BAP}
  ExecuteSQL('DELETE FROM CPBONSAPAYER');
  MoveCur(False);

  // BPY le 20/07/2004 => Fiche 12276 : effacement dans croiseaxe
  ExecuteSQl('DELETE FROM CROISEAXE');
  MoveCur(False);

  // GCO - 03/12/2002 FQ 13088
  ExecuteSQL('DELETE FROM CBALSIT');
  MoveCur(False);

  ExecuteSQL('DELETE FROM CBALSITECR');
  MoveCur(False);

  // suppression tableau des variations
  ExecuteSQL('DELETE FROM CPTABLEAUVAR');
  MoveCur(False);

  // suppression note de travail
  ExecuteSQL('DELETE FROM CPNOTETRAVAIL');
  MoveCur(False);

  // Suppression de la révision du dossier
  ExecuteSQL('DELETE FROM CREVCYCLE');
  ExecuteSQL('DELETE FROM CREVHISTOCYCLE');
  ExecuteSQL('DELETE FROM CREVGENERAUX');
  ExecuteSQL('DELETE FROM CREVBLOCNOTE');
  ExecuteSQL('DELETE FROM CREVINFODOSSIER WHERE CIR_NODOSSIER = "' + V_Pgi.NoDossier + '"');
  ExecuteSQL('UPDATE GENERAUX SET G_CYCLEREVISION = "", G_VISAREVISION = "-"');
  MoveCur(False);
  
  ExecuteSQL('UPDATE EXERCICE SET EX_ETATCPTA="NON", EX_ETATBUDGET="NON", EX_VALIDEE="------------------------"');
  MoveCur(False);

  // suppression des créances douteuses
  ExecuteSQL('DELETE FROM CPGCDOPERATION');
  MoveCur(False);

 // suppression des créances douteuses
  ExecuteSQL('DELETE FROM CPGCDCUMULS');
  MoveCur(False);

  // FQ 19021 : SBO 30/10/2006 : pas de RAZ de la table si elle est partagée
  if (not EstMultiSoc) or (EstMultiSoc and not EstTablePartagee('GENERAUX')) then
    begin
    Sql := 'UPDATE GENERAUX SET G_DATEDERNMVT="' + UsDateTime(DDeb) + '", G_DEBITDERNMVT=0, G_CREDITDERNMVT=0, ' +
      'G_NUMDERNMVT=0, G_LIGNEDERNMVT=0, G_TOTALDEBIT=0, G_TOTALCREDIT=0, ' +
      'G_DERNLETTRAGE="", G_DEBNONPOINTE=0, G_CREDNONPOINTE=0, ' +
      'G_TOTDEBP=0, G_TOTCREP=0, G_TOTDEBE=0, G_TOTCREE=0, G_TOTDEBS=0, G_TOTCRES=0, ' +
      'G_TOTDEBN2=0, G_TOTCREN2=0, ' +
      'G_TOTDEBANO=0, G_TOTCREANO=0, G_TOTDEBPTP=0, G_TOTCREPTP=0, G_TOTDEBPTD=0, G_TOTCREPTD=0';
    ExecuteSQL(Sql);
    MoveCur(False);
    end ;

  ExecuteSQL('DELETE FROM IMPECR');
  MoveCur(False);

  // FQ 19021 : SBO 30/10/2006 : pas de RAZ de la table si elle est partagée
  if (not EstMultiSoc) or (EstMultiSoc and not EstTablePartagee('JOURNAL')) then
    begin
    Sql := 'UPDATE JOURNAL SET J_DATEDERNMVT="' + UsDateTime(DDeb) +
      '", J_DEBITDERNMVT=0, J_CREDITDERNMVT=0, ' +
      'J_NUMDERNMVT=0, J_TOTALDEBIT=0, J_TOTALCREDIT=0, ' +
      'J_TOTDEBP=0, J_TOTCREP=0, J_TOTDEBE=0, J_TOTCREE=0, J_TOTDEBS=0, J_TOTCRES=0, ' +
      'J_VALIDEEN="------------------------", J_VALIDEEN1="------------------------"';
    ExecuteSQL(Sql);
    MoveCur(False);
    end ;

  // FQ 19021 : SBO 30/10/2006 : pas de RAZ de la table si elle est partagée
  if (not EstMultiSoc) or (EstMultiSoc and not EstTablePartagee('SECTION')) then
    begin
    Sql := 'UPDATE SECTION SET S_DATEDERNMVT="' + UsDateTime(DDeb) +
      '", S_DEBITDERNMVT=0, S_CREDITDERNMVT=0, ' +
      'S_NUMDERNMVT=0, S_LIGNEDERNMVT=0, S_TOTALDEBIT=0, S_TOTALCREDIT=0, ' +
      'S_TOTDEBP=0, S_TOTCREP=0, S_TOTDEBE=0, S_TOTCREE=0, S_TOTDEBS=0, S_TOTCRES=0, ' +
      'S_TOTDEBANO=0, S_TOTCREANO=0';
    ExecuteSQL(Sql);
    MoveCur(False);
    end ;

{$IFDEF SPEC302}
  Sql := 'UPDATE SOCIETE SET SO_DATECLOTUREPER="' + UsDateTime(DDeb) +
    '", SO_DATECLOTUREPRO="' + UsDateTime(DDeb) + '", ' +
    'SO_DATEDERNENTREE="' + UsDateTime(DDeb) + '", ' +
    'SO_DATECLOTUREPER="' + UsDateTime(DDeb) + '", SO_DATECLOTUREPRO="' +
      UsDateTime(DDeb) + '" ';
  ExecuteSQL(Sql);
  MoveCur(False);
{$ELSE}
  SetParamSoc('SO_DATECLOTUREPER', DDeb);
  SetParamSoc('SO_DATECLOTUREPRO', DDeb);
  SetParamSoc('SO_DATEDERNENTREE', DDeb);
  SetParamSoc('SO_EXOV8', ''); // CA - 13/12/2002
{$ENDIF}
  MoveCur(False);

  Sql := 'UPDATE SOUCHE SET SH_NUMDEPART=1, SH_DATEDEBUT="' + UsDateTime(DDeb) +
    '", SH_DATEFIN="' + UsDateTime(DFin) + '" WHERE SH_TYPE="CPT" ';
  ExecuteSQL(Sql);
  MoveCur(False);

  // FQ 19021 : SBO 30/10/2006 : pas de RAZ de la table si elle est partagée
  if (not EstMultiSoc) or (EstMultiSoc and not EstTablePartagee('TIERS')) then
    begin
    Sql := 'UPDATE TIERS SET T_DATEDERNMVT="' + UsDateTime(DDeb) +
      '", T_DEBITDERNMVT=0, T_CREDITDERNMVT=0, ' +
      'T_NUMDERNMVT=0, T_LIGNEDERNMVT=0, T_TOTALDEBIT=0, T_TOTALCREDIT=0, ' +
      'T_DERNLETTRAGE="", T_DATEDERNRELEVE="' + UsDateTime(DDeb) +
        '", T_SCORERELANCE=0, ' +
      'T_TOTDEBP=0, T_TOTCREP=0, T_TOTDEBE=0, T_TOTCREE=0, T_TOTDEBS=0, T_TOTCRES=0, ' +
      'T_TOTDEBANO=0, T_TOTCREANO=0';
    ExecuteSQL(Sql);
    MoveCur(False);
    end ;

  // FQ 18765 Raz des immos sérialisés ou non
  // + raz de IMMOMVTD
  // et maj de SO_IMMODPIEC
  // + RAZ IMMOREGFR du passage régime forfait à réel BTY 06/11/07

  //if ((VH^.OkModImmo) or (V_PGI.VersionDemo)) then
  //begin

    ExecuteSQL('DELETE FROM IMMO');
    MoveCur(False);

    ExecuteSQL('DELETE FROM IMMOLOG');
    MoveCur(False);

    ExecuteSQL('DELETE FROM IMMOAMOR');
    MoveCur(False);

    ExecuteSQL('DELETE FROM IMMOECHE');
    MoveCur(False);

    ExecuteSQL('DELETE FROM IMMOCPTE');
    MoveCur(False);

    ExecuteSQL('DELETE FROM IMMOUO');
    MoveCur(False);

    ExecuteSQL('DELETE FROM IMMOMVTD');
    MoveCur(False);

    ExecuteSQL('DELETE FROM IMMOREGFR');
    MoveCur(False);

    SetParamSoc('SO_DATECLOTUREIMMO', 2);
    SetParamSoc('SO_EXOCLOIMMO', '');
    SetParamSoc('SO_IMMODPIEC', '-');

  //end;

  FiniMove;
  PBouton.Enabled := True;
  HShowMessage(Msg.Mess[2], '', '') ;
  FOui.Visible := False;
  HLabel1.Visible := False;
  Memo1.Lines.Clear;
  TDeb.Visible := True;
  TFin.Visible := True;
  Deb.Visible := True;
  Fin.Visible := True;
  Memo1.Lines.Add(Msg.Mess[8]);
  Memo1.Lines.Add('');
  Memo1.Lines.Add(Msg.Mess[9]);
  Memo1.Lines.Add('');
  Memo1.Lines.Add(Msg.Mess[10]);
  BExo.Enabled := True;
  BValider.Enabled := False;
end;

procedure TFRAZ.FormShow(Sender: TObject);
begin
  Deb.Text := StDate1900;
  Fin.Text := StDate1900;
  TDeb.Visible := False;
  TFin.Visible := False;
  Deb.Visible := False;
  Fin.Visible := False;
  BExo.Enabled := False;
  FOui.SetFocus;
  Memo1.Lines[0]:=TraduireMemoire(Memo1.Lines[0]) ;
  Memo1.Lines[2]:=TraduireMemoire(Memo1.Lines[2]) ;
  Memo1.Lines[4]:=TraduireMemoire(Memo1.Lines[4]) ;
end;

procedure TFRAZ.BExoClick(Sender: TObject);
var
  D1, D2: TDateTime;
  DateExo: TExoDate;
  DDeb: TDateTime;
begin
  D1 := StrToDate(Deb.Text);
  D2 := StrToDate(Fin.Text);
  DDeb := 0;
  DateExo.Deb := D1;
  DateExo.Fin := D2;

  if D2 < D1 then
  begin
    PgiInfo('La date de fin d''exercice doit être supérieure ou égale à la date de début.', 'Ouverture d''exercice');
    Fin.SetFocus;
    Exit;
  end;

  // CA - 22/04/2003 - FQ 12259
  if not CControleDureeExercice(D1, D2) then
  begin
    PgiInfo('La durée d''exercice ne doit pas excéder 24 mois.', 'Ouverture d''exercice');
    Fin.SetFocus;
    Exit;
  end;
  
  ExecuteSql('Delete From Exercice');
  ExecuteSql('INSERT INTO EXERCICE (EX_EXERCICE,EX_LIBELLE,EX_ABREGE,EX_DATEDEBUT,EX_DATEFIN,'
    +
    'EX_ETATCPTA,EX_ETATBUDGET,EX_ETATADV,EX_ETATAPPRO,EX_ETATPROD,EX_SOCIETE,EX_VALIDEE,' +
    'EX_DATECUM,EX_DATECUMRUB,EX_BUDJAL,EX_NATEXO ) ' +
    //           'VALUES ("001","'+Msg.Mess[12]+'","'+Copy(Msg.Mess[12],1,17)+'","'+UsDateTime(D1)+'",'+
    // CA - 22/04/2003 - 10526 - Libellé exercice.
    'VALUES ("001","' + LibelleExerciceDefaut(D1, D2, False) + '","' +
      LibelleExerciceDefaut(D1, D2, True) + '","' + UsDateTime(D1) + '",' +
    '"' + UsDateTime(D2) + '","OUV","OUV","NON","NON","NON","' +
      V_PGI.CodeSociete + '",' +
    '"------------------------","' + UsDateTime(DDeb) + '","' + UsDateTime(DDeb)
      + '","","")');
  HShowMessage(Msg.Mess[13], '', '');
  VH^.CPEXORef.Code := ''; // CA - 27/04/2004 - Sinon, CPExoRef n'est pas réaffecté dans ChargeMagExo
  {$IFDEF EAGLCLIENT}    { FQ 16112 - CA - 12/07/2005 }
  AvertirCacheServer('EXERCICE');
  AvertirCacheServer('PARAMSOC');
  {$ENDIF}
  AvertirMultiTable('ttExercice');
  ChargeMagExo(True);
  Close;
end;

procedure TFRAZ.DebExit(Sender: TObject);
begin
(*
  if Deb.Text <> '' then
    Deb.Text := DateToStr(DebutDeMois(StrToDate(Deb.Text)));
*)
end;

procedure TFRAZ.FinExit(Sender: TObject);
begin
(*
  if Fin.Text <> '' then
    Fin.Text := DateToStr(FinDeMois(StrToDate(Fin.Text)));
*)
end;

procedure TFRAZ.HelpBtnClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

end.

