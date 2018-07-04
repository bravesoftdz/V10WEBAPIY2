{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/08/2003
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : PGMULDAD2SHONOR ()
Suite ........ :
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
{
PT1   : 28/03/2007 VG V_70 Portage CWAS
PT2   : 02/05/2007 FC V_72 FQ 13870 Rajout de la duplication d'honoraires
PT3   : 16/11/2007 NA V_80 Section = 01 pour les honoraires
}
unit UTOFPG_MULDADS2_HONOR;

interface

uses
  HStatus,
  HEnt1,
  ed_tools,
  UTOF,
  HCtrls,
  hqry,
{$IFDEF EAGLCLIENT}
  MaineAgl,
  emul,
{$ELSE}
  HDB,
  mul,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  FE_Main,
{$ENDIF}
  UTOB,
  HTB97,
  sysutils,
  HMsgBox,
  Classes,
  AGLInit,
  PGDADSCommun,
  UtomDADS2HONORAIRES;       //LanceFiche_DADS2Honoraire

type
  TOF_PGMULDAD2SHONOR = class(TOF)
  public
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;

  private
    WW: THEdit; // Clause XX_WHERE
    Q_Mul: THQuery;
    BCherche: TToolbarButton97;
    param: string;
    ComboAnnee2: THValComboBox;
    THonor2 : TOB;
{$IFNDEF EAGLCLIENT}
    Liste: THDBGrid;
{$ELSE}
    Liste: THGrid;
{$ENDIF}

    procedure ActiveWhere(Sender: TObject);
    procedure GrilleDblClick(Sender: TObject);
    procedure NewHonor(Sender: TObject);
    //DEB PT2
    procedure GrilleClick(Sender: TObject);
    procedure Duplication(Sender: TObject);
    procedure Duplique_un(var StOrdre : String);
    procedure Annee2Change(Sender: TObject);
    procedure ChargeTOBHonoraire(Honoraire : String);
    procedure Duplic(StOrdre,Honoraire : String);
    //FIN PT2
  end;

  procedure LanceMul_DADS2Honoraires;

implementation

uses DB
{$IFDEF COMPTA}
    ,Ent1 // FQ 22285 GetEncours
{$ENDIF}
    ;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : procédure pour appel depuis unités externes
Mots clefs ... : 
*****************************************************************}
procedure LanceMul_DADS2Honoraires;
begin
  AGLLanceFiche('PAY', 'MUL_DADS2_HONOR', '', '', 'S');
end;

procedure TOF_PGMULDAD2SHONOR.OnNew;
begin
  inherited;
end;


procedure TOF_PGMULDAD2SHONOR.OnDelete;
begin
  inherited;
end;


procedure TOF_PGMULDAD2SHONOR.OnUpdate;
begin
  inherited;
  if (param = 'S') then
    TFMul(Ecran).BOuvrir.Enabled := True;
  if ((param = 'D') and ((TFMul(Ecran).FListe.nbSelected <> 0) or (TFMul(Ecran).FListe.AllSelected=True))) then
     TFMul(Ecran).BOuvrir.Enabled := True;
end;

procedure TOF_PGMULDAD2SHONOR.OnLoad;
begin
  inherited;
  ActiveWhere(nil);
end;


procedure TOF_PGMULDAD2SHONOR.OnArgument(S: string);
var
  AnneePrec,AnneeCours: string;
  JourJ: TDateTime;
  AnneeA, Jour, MoisM: Word;
  TT: TFMul;
  Arg : String;
  LibAnnee1, LibAnnee2 : THLabel;
begin
  inherited;
  //DEB PT2
  Arg := S;
  Param := Trim(ReadTokenPipe(Arg, ';'));

  TT := TFMul(Ecran);

  {$IFNDEF EAGLCLIENT}
  Liste := THDBGrid (GetControl ('FListe'));
  {$ELSE}
  Liste := THGrid (GetControl ('FListe'));
  {$ENDIF}
  LibAnnee1 := THLabel (GetControl ('LANNEE'));
  LibAnnee2 := THLabel (GetControl ('LANNEE1'));
  ComboAnnee2 := THValComboBox (GetControl ('ANNEE1'));

  if (param = 'S') then
  begin
    TFMul(Ecran).Caption := 'Saisie - Consultation des honoraires';
    TFMul(Ecran).BOuvrir.OnClick := GrilleDblClick;
    TFMul(Ecran).BInsert.OnClick := NewHonor;
    if Liste <> nil then
    begin
      {$IFNDEF EAGLCLIENT}
        Liste.MultiSelection := False;
      {$ENDIF}
      Liste.OnDblClick := GrilleDblClick;
    end;
  end
  else if (param = 'D') then
  begin
    TFMul (Ecran).Caption := 'Duplication des honoraires';
    TFMul (Ecran).BOuvrir.OnClick := Duplication;
    TFMul (Ecran).BOuvrir.Enabled := False;
    TFMul (Ecran).BInsert.Visible := False;
    {$IFNDEF EAGLCLIENT}
      TFMul (Ecran).FListe.MultiSelection:= True;
    {$ENDIF}
    TFMul (Ecran).bSelectAll.Visible:= True;
    if (LibAnnee1 <> nil) then
      LibAnnee1.Caption:= 'De l''année';
    if (LibAnnee2 <> nil) then
      LibAnnee2.Visible:= True;
    if (ComboAnnee2 <> nil) then
    begin
      ComboAnnee2.Visible:= True;
      ComboAnnee2.OnChange:= Annee2Change;
    end;
    if (Liste <> nil) then
    begin
      Liste.OnFlipSelection:= GrilleClick;
      TFMul (Ecran).bSelectAll.OnClick:= GrilleClick;
    end;
  end;
  //FIN PT2

  if TT <> nil then
    UpdateCaption(TT);

  Q_Mul := THQuery(Ecran.FindComponent('Q'));

  //DEB PT2                   
{ FQ 22285 BVE 28.01.08 }
{$IFDEF COMPTA}
  JourJ := GetEnCours.Deb ;
{$ELSE}
  JourJ := Date;
{$ENDIF}
{ END FQ 22285 }
  DecodeDate(JourJ, AnneeA, MoisM, Jour);
  if (param = 'D') then
  begin
    if (MoisM > 9) then
    begin
      AnneePrec := IntToStr(AnneeA - 1);
      AnneeCours := IntToStr(AnneeA);
    end
    else
    begin
      AnneePrec := IntToStr(AnneeA - 2);
      AnneeCours := IntToStr(AnneeA - 1);
    end;
    SetControlText('ANNEE', copy(AnneePrec, 1, 1) + copy(AnneePrec, 3, 2));
    ComboAnnee2.Value := copy(AnneeCours, 1, 1) + copy(AnneeCours, 3, 2); 
    PGAnnee := GetControlText('ANNEE');
    PGExercice := AnneePrec;
  end
  else
  begin
    if (MoisM > 9) then
    begin
      AnneePrec:= IntToStr(AnneeA);
      AnneeCours:= IntToStr(AnneeA + 1);
    end
    else
    begin
      AnneePrec:= IntToStr(AnneeA - 1);
      AnneeCours:= IntToStr(AnneeA);
    end;
    SetControlText('ANNEE', copy(AnneePrec, 1, 1) + copy(AnneePrec, 3, 2));
    PGAnnee := GetControlText('ANNEE');
    PGExercice := AnneePrec;
  end;
  //FIN PT2

  if Q_Mul <> nil then
    TFMul(Ecran).SetDBListe('PGHONORDADS2');

  WW := THEdit(GetControl('XX_WHERE'));
  BCherche := TToolbarButton97(GetControl('BCherche'));
  ActiveWhere(nil);
end;

//DEB PT2
procedure TOF_PGMULDAD2SHONOR.GrilleClick(Sender: TObject);
begin
  {$IFNDEF EAGLCLIENT}
  Liste:= THDBGrid(GetControl('FListe'));
  {$ELSE}
  Liste:= THGrid(GetControl('FListe'));
  {$ENDIF}
  if Sender = TFMul(Ecran).bSelectAll then
    Liste.AllSelected := not Liste.AllSelected ;

  if (param = 'D') then
  begin
    if Liste <> NIL then
    BEGIN
      if (Liste.NbSelected<>0) or (Liste.AllSelected) then
        TFMul(Ecran).BOuvrir.Enabled:= True
      else
        TFMul(Ecran).BOuvrir.Enabled:= False;
    END;
  end;
end;
//FIN PT2

procedure TOF_PGMULDAD2SHONOR.OnClose;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/08/2003
Modifié le ... :   /  /
Description .. : Double-click sur la grille
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOF_PGMULDAD2SHONOR.GrilleDblClick(Sender: TObject);
var
Annee, Ordre: string;
begin
Annee:= '';
if ((Q_Mul <> nil) and (Q_Mul.RecordCount = 0)) then
   exit;

{$IFDEF EAGLCLIENT}
TFMul (Ecran).Q.TQ.Seek (TFMul (Ecran).FListe.Row-1);         //PT1
{$ENDIF}

Ordre:= Q_Mul.FindField ('PDH_HONORAIRE').AsString;

if (PGAnnee <> '') then
   Annee:= RechDom ('PGANNEE', PGAnnee, FALSE);

{$IFNDEF EAGLCLIENT}
TheMulQ:= THQuery (Ecran.FindComponent('Q'));
{$ELSE}
TheMulQ:= TOB (Ecran.FindComponent('Q'));
{$ENDIF}

if (PGAnnee <> '') then
   LanceFiche_DADS2Honoraire('', Ordre+';'+Annee, 'ACTION=MODIFICATION;'+Ordre+';'+Annee)
//   AGLLanceFiche ('PAY', 'DADS2_HONOR', '', Ordre+';'+Annee,
//                  'ACTION=MODIFICATION;'+Ordre+';'+Annee)
else
   PGIBox ('L''année n''est pas valide', 'TD Bilatéral');

if (BCherche <> nil) then
   BCherche.click;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/08/2003
Modifié le ... :   /  /
Description .. : XX_WHERE
Mots clefs ... :
*****************************************************************}

procedure TOF_PGMULDAD2SHONOR.ActiveWhere(Sender: TObject);
var
  Where: string;
begin
  PGAnnee := GetControlText('ANNEE');
  if (WW <> nil) then
  begin
    if Q_Mul <> nil then
      TFMul(Ecran).SetDBListe('PGHONORDADS2');

    Where := ' PDH_VALIDITE="' + RechDom('PGANNEE', PGAnnee, FALSE) + '"';
    SetControlText('XX_WHERE', where);

    TFMul(Ecran).BOuvrir.Enabled := False;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/08/2003
Modifié le ... :   /  /
Description .. : Création d'un honoraire
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}

procedure TOF_PGMULDAD2SHONOR.NewHonor(Sender: TObject);
var
  OrdreDest: integer;
  QHonor: TQuery;
  Annee, StOrdre: string;
begin
  if PGAnnee <> '' then
    Annee := RechDom('PGANNEE', PGAnnee, FALSE)
  else
    Annee := '';
  OrdreDest := 1;
  QHonor := OpenSQL('SELECT MAX(PDH_HONORAIRE) AS NUMMAX FROM DADS2HONORAIRES', True);
  if not QHonor.eof then
  try
    OrdreDest := (QHonor.FindField('NUMMAX').AsInteger) + 1;
  except
    on E: EConvertError do
      OrdreDest := 1;
  end;
  Ferme(QHonor);
  StOrdre := FormatFloat('00000000000000000', OrdreDest);
  LanceFiche_DADS2Honoraire('', StOrdre + ';' + Annee, 'ACTION=CREATION;' + StOrdre + ';' + Annee);
  //  AGLLanceFiche('PAY', 'DADS2_HONOR', '', StOrdre + ';' + Annee, 'ACTION=CREATION;' + StOrdre + ';' + Annee);
  if BCherche <> nil then
    BCherche.click;
end;

//DEB PT2
procedure TOF_PGMULDAD2SHONOR.Duplication(Sender: TObject);
var
  i,IMax : integer;
  QRechNumHonor :TQuery;
  StOrdre: string;
begin
  {$IFNDEF EAGLCLIENT}
  Liste:= THDBGrid(GetControl('FListe'));
  {$ELSE}
  Liste:= THGrid(GetControl('FListe'));
  {$ENDIF}
  if Liste <> NIL then
  BEGIN
    if (Liste.NbSelected=0) and (not Liste.AllSelected) then
    begin
      MessageAlerte('Aucun élément sélectionné');
      exit;
    end;

    IMax:=1;
    QRechNumHonor := OpenSQL('SELECT PDH_HONORAIRE FROM DADS2HONORAIRES ORDER BY PDH_HONORAIRE DESC', TRUE);
    if not QRechNumHonor.EOF then
    begin
      StOrdre := QRechNumHonor.FindField('PDH_HONORAIRE').AsString;
      IMax := StrToInt(StOrdre)+1;
    end;
    Ferme(QRechNumHonor);

    StOrdre := FormatFloat('00000000000000000', IMax);

    if (Liste.AllSelected=TRUE) then
    begin
      InitMoveProgressForm (NIL, 'Duplication en cours',
                              'Veuillez patienter SVP ...',
                              TFmul(Ecran).Q.RecordCount,FALSE,TRUE);
      InitMove (TFmul(Ecran).Q.RecordCount,'');

      {$IFDEF EAGLCLIENT}
        if (TFMul(Ecran).bSelectAll.Down) then
          TFMul(Ecran).Fetchlestous;
      {$ENDIF}

      TFmul(Ecran).Q.First;
      while Not TFmul(Ecran).Q.EOF do
      BEGIN
        Duplique_un(StOrdre);
        TFmul(Ecran).Q.Next;
      END;

      Liste.AllSelected:= False;
      TFMul(Ecran).bSelectAll.Down:= Liste.AllSelected;
    end
    else
    begin
      InitMoveProgressForm (NIL, 'Duplication en cours',
                              'Veuillez patienter SVP ...', Liste.NbSelected,
                              FALSE, TRUE);
      InitMove(Liste.NbSelected,'');

      for i:=0 to Liste.NbSelected-1 do
      BEGIN
        Liste.GotoLeBOOKMARK(i);
        {$IFDEF EAGLCLIENT}
          TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row-1) ;
        {$ENDIF}
        Duplique_un(StOrdre);
      END;

      Liste.ClearSelected;
    end;

    FiniMove;
    FiniMoveProgressForm;
    TFMul(Ecran).BOuvrir.Enabled:= False;
    PGIBox ('Traitement terminé', 'Duplication des honoraires');

  END;
  if BCherche<>nil then
    BCherche.click;
end;

procedure TOF_PGMULDAD2SHONOR.Duplique_un(var StOrdre : String);
var
  Honoraire : String;
  IMax : Integer;
begin
  Honoraire:= TFmul(Ecran).Q.FindField('PDH_HONORAIRE').AsString;
  try
     begintrans;
     ChargeTOBHonoraire(Honoraire);
     Duplic(StOrdre,Honoraire);
     FreeAndNil(THonor2);
     CommitTrans;
  Except
     Rollback;
  end;

  MoveCur (False);
  MoveCurProgressForm(StOrdre);

  IMax := StrToInt(StOrdre) + 1;
  StOrdre := FormatFloat('00000000000000000', IMax);
end;

procedure TOF_PGMULDAD2SHONOR.ChargeTOBHonoraire(Honoraire : String);
var
  StHonor : string;
  QRechHonor : TQuery;
begin
  //Chargement de la TOB HONORAIRES
  StHonor := 'SELECT * FROM DADS2HONORAIRES WHERE PDH_HONORAIRE = "' + Honoraire + '"';
  QRechHonor := OpenSql(StHonor,TRUE);
  THonor2 := TOB.Create('Les honoraires', NIL, -1);
  THonor2.LoadDetailDB('DADS2HONORAIRES','','',QRechHonor,False);
  Ferme(QRechHonor);
end;

procedure TOF_PGMULDAD2SHONOR.Annee2Change(Sender: TObject);
begin
  TFMul(Ecran).BOuvrir.Enabled := False;
end;

procedure TOF_PGMULDAD2SHONOR.Duplic(StOrdre,Honoraire : String);
var
  TDADSHonoraire,THonorD : Tob;
begin
  THonorD := THonor2.FindFirst(['PDH_HONORAIRE'], [Honoraire], TRUE);
  TDADSHonoraire := TOB.Create ('DADS2HONORAIRES', THonor2, -1);
  TDADSHonoraire.PutValue('PDH_HONORAIRE', StOrdre);
  TDADSHonoraire.PutValue('PDH_VALIDITE', ComboAnnee2.text);
  TDADSHonoraire.PutValue('PDH_SIRET', THonorD.GetValue('PDH_SIRET'));
  TDADSHonoraire.PutValue('PDH_ETABLISSEMENT', THonorD.GetValue('PDH_ETABLISSEMENT'));
 // TDADSHonoraire.PutValue('PDH_SECTIONETAB', THonorD.GetValue('PDH_SECTIONETAB'));  // pt3
  TDADSHonoraire.PutValue('PDH_SECTIONETAB', '01');     // pt3
  TDADSHonoraire.PutValue('PDH_TYPEDADS', THonorD.GetValue('PDH_TYPEDADS'));
  TDADSHonoraire.PutValue('PDH_SIRETBEN', THonorD.GetValue('PDH_SIRETBEN'));
  TDADSHonoraire.PutValue('PDH_NOMBEN', THonorD.GetValue('PDH_NOMBEN'));
  TDADSHonoraire.PutValue('PDH_PRENOMBEN', THonorD.GetValue('PDH_PRENOMBEN'));
  TDADSHonoraire.PutValue('PDH_RAISONSOCBEN', THonorD.GetValue('PDH_RAISONSOCBEN'));
  TDADSHonoraire.PutValue('PDH_PROFESSIONBEN', THonorD.GetValue('PDH_PROFESSIONBEN'));
  TDADSHonoraire.PutValue('PDH_ADRCOMPL', THonorD.GetValue('PDH_ADRCOMPL'));
  TDADSHonoraire.PutValue('PDH_ADRNUM', THonorD.GetValue('PDH_ADRNUM'));
  TDADSHonoraire.PutValue('PDH_ADRBISTER', THonorD.GetValue('PDH_ADRBISTER'));
  TDADSHonoraire.PutValue('PDH_ADRNOM', THonorD.GetValue('PDH_ADRNOM'));
  TDADSHonoraire.PutValue('PDH_ADRCOMMINSEE', THonorD.GetValue('PDH_ADRCOMMINSEE'));
  TDADSHonoraire.PutValue('PDH_ADRCOMMUNE', THonorD.GetValue('PDH_ADRCOMMUNE'));
  TDADSHonoraire.PutValue('PDH_CODEPOSTAL', THonorD.GetValue('PDH_CODEPOSTAL'));
  TDADSHonoraire.PutValue('PDH_BUREAUDISTRIB', THonorD.GetValue('PDH_BUREAUDISTRIB'));
  TDADSHonoraire.PutValue('PDH_REMHONOR', 0);
  TDADSHonoraire.PutValue('PDH_REMCOMMISS', 0);
  TDADSHonoraire.PutValue('PDH_REMCOURTAGE', 0);
  TDADSHonoraire.PutValue('PDH_REMRISTOURNE', 0);
  TDADSHonoraire.PutValue('PDH_REMJETON', 0);
  TDADSHonoraire.PutValue('PDH_REMAUTEUR', 0);
  TDADSHonoraire.PutValue('PDH_REMINVENT', 0);
  TDADSHonoraire.PutValue('PDH_REMAUTRE', 0);
  TDADSHonoraire.PutValue('PDH_REMINDEMNITE', 0);
  TDADSHonoraire.PutValue('PDH_REMAVANTAGE', 0);
  TDADSHonoraire.PutValue('PDH_RETENUESOURC', 0);
  TDADSHonoraire.PutValue('PDH_AVANTAGENATN', '');
  TDADSHonoraire.PutValue('PDH_NTIC', '');
  TDADSHonoraire.PutValue('PDH_CHARGEINDEMN', '');
  TDADSHonoraire.PutValue('PDH_TAUXSOURCE', '');
  TDADSHonoraire.PutValue('PDH_TVAAUTEUR', 0);
  TDADSHonoraire.InsertOrUpdateDB;
end;
//FIN PT2


initialization
  registerclasses([TOF_PGMULDAD2SHONOR]);
end.

