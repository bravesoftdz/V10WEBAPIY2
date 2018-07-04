{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 23/08/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGRETROSIMULCOT ()
Mots clefs ... : TOF;PGRETROSIMULCOT
*****************************************************************
PT1   : 27/08/2007 FC  V_72 Ajout onglet Compléments
}
Unit UtofPGRETROSIMULCOT ;

Interface

Uses
{$IFDEF VER150}
     Variants,
{$ENDIF}
     StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,
{$else}
     MainEagl,
     eMul,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     Ed_Tools,
     utobdebug,
     UTOF,
     EntPaie,
     P5Def ;


Type
  TOF_PGRETROSIMULCOT = Class (TOF)
  private
    procedure ChangeSortie ( Sender : Tobject );
    procedure LanceSimultation ( Sender : TObject );
    procedure CalculTx;

  public
    btOuvrir            : TToolbarButton97;
    TxSal               : extended;
    TxPat               : extended;
    LeWhere             : String;
{
    BaseRem : Real;
    TauxRem : Real;
    CoefRem : Real;
    MtRem : Real  ;
}
//    NatureRub : string;
    Rubrique            : String;
    Dossier             : String;
    Predefini           : String;
    Datevalidite        : String;
    Datedebut           : String;
    Datefin             : String;
    LibCot              : String;
    TobHisto            : Tob;
    ckSortie            : THCheckbox;

    procedure OnArgument (Arguments : String ) ; override ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

  procedure PgRetroSimulcot_LanceFiche ( Arguments : String );

Implementation

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 01/12/2006
Modifié le ... :   /  /
Description .. : lancement de la fiche de simulation de calcul des
Suite ........ : cotisations
Mots clefs ... :
*****************************************************************}
procedure PgRetroSimulcot_LanceFiche ( Arguments : String );
begin
  AGLLanceFiche('PAY','PG_RETROSIMULCOT','','', Arguments);
end;

procedure TOF_PGRETROSIMULCOT.OnArgument (Arguments : String ) ;
var
  Q                     : TQuery;
  F                     : TFMul;
  StrCaption            : String;
  Num : integer;
begin
  Inherited ;
  F                     := TFMul(Ecran);

//  NatureRub := READTOKENST(Arguments);
  Rubrique              := READTOKENST(Arguments);
  Dossier               := READTOKENST(Arguments);
  Predefini             := READTOKENST(Arguments);
  Datevalidite          := READTOKENST(Arguments);
  Datedebut             := READTOKENST(Arguments);
  Datefin               := READTOKENST(Arguments);

  //DEB PT1
  for Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    if Num > 4 then Break;
    VisibiliteChampSalarie (IntToStr (Num),
                            GetControl ('PSA_TRAVAILN'+IntToStr (Num)),
                            GetControl ('TPSA_TRAVAILN'+IntToStr (Num)));
  end;
  VisibiliteStat (GetControl ('PSA_CODESTAT'), GetControl ('TPSA_CODESTAT'));
  //FIN PT1

  LeWhere               := 'PHB_NATURERUB="COT" AND PHB_RUBRIQUE="' + Rubrique + '" AND ' +
                          'PHB_DATEDEBUT>="' + USDATETIME(strtodate(Datedebut)) +
                          '" AND PHB_DATEFIN<="' + USDATETIME(strtodate(Datefin)) + '"';

  SetControlText ('XX_WHERE', LeWhere);

  btOuvrir              := TToolbarButton97(GetControl('BOuvrir'));
  btOuvrir.OnClick      := LanceSimultation;

  ckSortie              := THCheckbox(GetControl('CKSORTIE'));
  ckSortie.OnClick      := ChangeSortie;

  Q                     := OpenSQL('SELECT PGR_DIFFTXSAL,PGR_DIFFTXPAT,PGR_DIFFBASEREM FROM PGRETROACTIV WHERE PGR_NATURERUB="COT"' +
                          ' AND PGR_RUBRIQUE="' + Rubrique + '" AND PGR_NODOSSIER="' + Dossier +
                          '" AND PGR_PREDEFINI="' + Predefini + '" AND PGR_DATEVALIDITE="' +
                          USDATETIME(strtodate(Datevalidite)) + '" AND ##PGR_PREDEFINI##', True);

  TxSal                 := Q.Fields[0].AsFloat;
  TxPat                 := Q.Fields[1].AsFloat;
  Ferme(Q);

  TobHisto              := Tob.Create('LES BULLETINS', nil, -1);

  Q                     := OpenSQL('SELECT PCT_LIBELLE FROM COTISATION WHERE PCT_RUBRIQUE="' + Rubrique +
                          '" AND ##PCT_PREDEFINI##', True);
  LibCot                := Q.Fields[0].AsString;
  StrCaption            := 'Simulation sur ' + Rubrique + ' : ' + LibCot + ' du ' + Datedebut + ' au ' + Datefin;
  F.Caption             := StrCaption;
  Ferme(Q);
  UpdateCaption(F);
end ;

procedure TOF_PGRETROSIMULCOT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGRETROSIMULCOT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGRETROSIMULCOT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGRETROSIMULCOT.OnLoad ;
var
  CompWhere             : String;
  LaDate                : String;

begin
  CompWhere := '';
  if GetCheckBoxState('CKSORTIE') = cbgrayed then
  begin
    LaDate              := USDATETIME(strtoDate(GetControlText('DATESORTIE')));
    CompWhere           := ' AND (PSA_DATESORTIE>="' + LaDate + '" OR PSA_DATESORTIE="")';
  end;
  if GetCheckBoxState('CKSORTIE') = cbchecked then
  begin
    LaDate              := USDATETIME(strtoDate(GetControlText('DATESORTIE')));
    CompWhere           := ' AND PSA_DATESORTIE>="' + LaDate + '"';
  end;
  //DEB PT1
  if (GetControlText('POPULATION') <> '') and (GetControlText('POPULATION') <> '<<Tous>>') then
    CompWhere := CompWhere + ' AND PSA_SALARIE IN (SELECT PNA_SALARIE FROM SALARIEPOPUL WHERE PNA_TYPEPOP="PAI"' +
      ' AND PNA_POPULATION="' + GetControlText('POPULATION') + '")';
  //FIN PT1
  SetControlText('XX_WHERE', LeWhere + CompWhere);

  Inherited ;
end ;

procedure TOF_PGRETROSIMULCOT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGRETROSIMULCOT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGRETROSIMULCOT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PGRETROSIMULCOT.LanceSimultation(Sender: TObject);
var
  F                     : TFmul;
  i                     : integer;

begin
  if GetField('PHB_SALARIE') = null then
    exit;

  F                     := TFmul(ecran);
  //Traitement sur la sélection
  if (F.FListe.nbSelected = 0) and (not F.FListe.AllSelected) then
  begin
    PGIInfo('Aucun élément sélectionné.', F.Caption);
    exit;
  end;


  if F.FListe.AllSelected then
  begin
  {$IFDEF EAGLCLIENT}
      if not F.FetchLesTous then
        PGIInfo('Impossible de récupérer tous les enregistrements')
      else
  {$ENDIF}
     F.Q.First;
     InitMoveProgressForm(nil,'Simulation en cours ...', '', (F.Q.recordcount)*2, True, True);
     while not F.Q.Eof do
     begin
        CalculTx;
        MoveCurProgressForm('Recherche données');
        F.Q.Next;
     end;
  end
  else
  begin
     InitMoveProgressForm(nil,'Simulation en cours ...', '', (F.FListe.nbSelected-1)*2, True, True);
     for i := 0 to F.FListe.nbSelected-1 do
     begin
        F.Fliste.GotoLeBookmark(i);
        {$IFDEF EAGLCLIENT}
        F.Q.TQ.Seek(F.FListe.row -1);
        {$ENDIF}
        CalculTx;

        MoveCurProgressForm('Recherche données');

     end;
  end;

//  tobdebug(TobHisto);



  //enregistrement ou mise à jour dans l'historique retro
  for i := 0 to TobHisto.Detail.Count-1 do
  begin
    TobHisto.Detail[i].InsertOrUpdateDB(False);
    MoveCurProgressForm('Traitement')
  end;

  FiniMoveProgressForm;
  F.Close;
end;

procedure TOF_PGRETROSIMULCOT.CalculTx;
var
  TobH                  : Tob;
  Calcul                : extended;
  TxSalDiff             : extended;
  TxPatDiff             : extended;
  BaseCot               : extended;


begin
  TobH                  := TobHisto.FindFirst(['PGT_SALARIE'], [GetField('PHB_SALARIE')], False);
  if TobH = nil then
  begin
    TobH                := Tob.Create('PGHISTRETRO', TobHisto, -1);
    TobH.PutValue('PGT_SALARIE', GetField('PHB_SALARIE'));
    TobH.PutValue('PGT_NATURERUB', 'COT');
    TobH.PutValue('PGT_RUBRIQUE', Rubrique);
    TobH.PutValue('PGT_DATEVALIDITE', StrToDate(Datevalidite));
    TobH.PutValue('PGT_DATEDEBUT', StrToDate(Datedebut));
    TobH.PutValue('PGT_DATEFIN', StrToDate(Datefin));
    TxSalDiff           := TxSal - GetField('PHB_TAUXSALARIAL') ;
    TobH.PutValue('PGT_DIFFTXSAL', TxSalDiff);
    TxPatDiff           := TxPat - GetField('PHB_TAUXPATRONAL') ;
    TobH.PutValue('PGT_DIFFTXPAT', TxPatDiff);
    BaseCot             := GetField('PHB_BASECOT');
    TobH.PutValue('PGT_DIFFBASECOT', BaseCot);
    Calcul              := TxSalDiff * BaseCot / 100;
    TobH.PutValue('PGT_DIFFMTSAL', Calcul);
    Calcul              := TxPatDiff * BaseCot / 100;
    TobH.PutValue('PGT_DIFFMTPAT', Calcul);
    TobH.PutValue('PGT_LIBELLE1', LibCot);
    TobH.PutValue('PGT_LIBELLE2', 'Du ' + Datedebut + ' au ' + Datefin);
    TobH.PutValue('PGT_DIFFBASEREM', 0);
    TobH.PutValue('PGT_DIFTAUXREM', 0);
    TobH.PutValue('PGT_DIFCOEFFREM', 0);
    TobH.PutValue('PGT_DIFFMTREM', 0);
  end
  else
  begin
    TxSalDiff           := TobH.GetValue('PGT_DIFFTXSAL');
    TxPatDiff           := TobH.GetValue('PGT_DIFFTXPAT');
    BaseCot             := GetField('PHB_BASECOT') + TobH.GetValue('PGT_DIFFBASECOT');
    TobH.PutValue('PGT_DIFFBASECOT', BaseCot);
    Calcul              := TxSalDiff * BaseCot / 100;
    TobH.PutValue('PGT_DIFFMTSAL', Calcul);
    Calcul              := TxPatDiff * BaseCot / 100;
    TobH.PutValue('PGT_DIFFMTPAT', Calcul);

  end;
end;

procedure TOF_PGRETROSIMULCOT.ChangeSortie(Sender: Tobject);
begin
  if (GetCheckBoxState('CKSORTIE') = cbchecked) or (GetCheckBoxState('CKSORTIE') = cbgrayed) then
  begin
    SetControlEnabled('DATESORTIE', True);
  end
  else
  begin
    SetControlEnabled('DATESORTIE', False);
  end;
end;

Initialization
  registerclasses ( [ TOF_PGRETROSIMULCOT ] ) ;
end.
