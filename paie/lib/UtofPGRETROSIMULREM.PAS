{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 28/08/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGHISTOSIMULREM ()
Mots clefs ... : TOF;PGHISTOSIMULREM
*****************************************************************
PT1   : 27/08/2007 FC  V_72 Ajout onglet Compléments
}
Unit UtofPGRETROSIMULREM ;

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
     forms,
     uTob,
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
  TOF_PGRETROSIMULREM = Class (TOF)
  private
    procedure ChangeSortie ( Sender : Tobject );
    procedure LanceSimulation ( Sender : TObject );
    procedure CalculRem;

  public
    LeWhere             : String;
    Rubrique            : String;
    Dossier             : String;
    Predefini           : String;
    Datevalidite        : String;
    Datedebut           : String;
    Datefin             : String;
    LibRem              : String;
    CodeCalcul          : String;
    BaseRem             : Extended;
    TauxRem             : Extended;
    CoefRem             : Extended;
    MtRem               : Extended;

    TobHisto            : TOB;
    bOuvrir             : TToolbarButton97;
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

  procedure PgRetroSimulrem_LanceFiche ( Arguments : String );

Implementation

uses UtofPGHISTRETRO;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 01/12/2006
Modifié le ... :   /  /    
Description .. : lancement de la fiche de simulation de calcul des 
Suite ........ : rémunération
Mots clefs ... :
*****************************************************************}
procedure PgRetroSimulrem_LanceFiche ( Arguments : String );
begin
  AGLLanceFiche('PAY','PG_RETROSIMULREM','','', Arguments);
end;

procedure TOF_PGRETROSIMULREM.OnArgument (Arguments : String ) ;
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

  LeWhere               := 'PHB_NATURERUB="AAA" AND PHB_RUBRIQUE="' + Rubrique + '" AND ' +
                          'PHB_DATEDEBUT>="' + USDATETIME(strtodate(Datedebut)) +
                          '" AND PHB_DATEFIN<="' + USDATETIME(strtodate(Datefin)) + '"';

  SetControlText ('XX_WHERE', LeWhere);

  Q                     := OpenSQL('SELECT PGR_DIFFBASEREM,PGR_DIFTAUXREM,PGR_DIFCOEFFREM,PGR_DIFFMTREM FROM PGRETROACTIV WHERE PGR_NATURERUB="AAA"'+
                          ' AND PGR_RUBRIQUE="' + Rubrique + '" AND PGR_NODOSSIER="' + Dossier +
                          '" AND PGR_PREDEFINI="' + Predefini + '" AND PGR_DATEVALIDITE="' +
                          USDATETIME(strtodate(Datevalidite)) + '" AND ##PGR_PREDEFINI##', True);

  BaseRem               := Q.Fields[0].AsFloat;
  TauxRem               := Q.Fields[1].AsFloat;
  CoefRem               := Q.Fields[2].AsFloat;
  MtRem                 := Q.Fields[3].AsFloat;
  Ferme(Q);

  bOuvrir               := TToolbarButton97(GetControl('BOuvrir'));
  bOuvrir.OnClick       := LanceSimulation;

  ckSortie              := THCheckbox(GetControl('CKSORTIE'));
  ckSortie.OnClick      := ChangeSortie;

  TobHisto := Tob.Create('LES BULLETINS', nil, -1);

  Q                     := OpenSQL('SELECT PRM_LIBELLE, PRM_CODECALCUL FROM REMUNERATION WHERE PRM_RUBRIQUE="' + Rubrique +
                          '" AND ##PRM_PREDEFINI##', True);
  LibRem                := Q.Fields[0].AsString;
  CodeCalcul            := Q.Fields[1].AsString;
  StrCaption            := 'Simulation sur ' + LibRem + ' du ' + Datedebut + ' au ' + Datefin;
  F.Caption             := StrCaption;

  Ferme(Q);
  UpdateCaption(F);

end ;

procedure TOF_PGRETROSIMULREM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGRETROSIMULREM.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGRETROSIMULREM.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGRETROSIMULREM.OnLoad ;
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
    CompWhere           := ' AND PSA_DATESORTIE>="' + Ladate + '"';
  end;
  //DEB PT1
  if (GetControlText('POPULATION') <> '') and (GetControlText('POPULATION') <> '<<Tous>>') then
    CompWhere := CompWhere + ' AND PSA_SALARIE IN (SELECT PNA_SALARIE FROM SALARIEPOPUL WHERE PNA_TYPEPOP="PAI"' +
      ' AND PNA_POPULATION="' + GetControlText('POPULATION') + '")';
  //FIN PT1
  SetControlText('XX_WHERE', LeWhere + CompWhere);

  Inherited ;
end ;

procedure TOF_PGRETROSIMULREM.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGRETROSIMULREM.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGRETROSIMULREM.OnCancel () ;
begin
  Inherited ;
end ;



procedure TOF_PGRETROSIMULREM.LanceSimulation(Sender: TObject);
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
     InitMoveProgressForm(nil,'Simulation en cours ...', '', (F.Q.recordcount)*2,True,True);
     while not F.Q.Eof do
     begin
        CalculRem;
        MoveCurProgressForm('Recherche données');
        F.Q.Next;
     end;
  end
  else
  begin
     InitMoveProgressForm(nil,'Simulation en cours', '', (F.FListe.nbSelected-1)*2,True,True);
     for i := 0 to F.FListe.nbSelected-1 do
     begin
        F.Fliste.GotoLeBookmark(i);
        {$IFDEF EAGLCLIENT}
        F.Q.TQ.Seek(F.FListe.row -1);
        {$ENDIF}
        CalculRem;

        MoveCurProgressForm('Recherche données');

     end;
  end;


  //enregistrement ou mise à jour dans l'historique retro
  for i := 0 to TobHisto.Detail.Count-1 do
  begin
    TobHisto.Detail[i].InsertOrUpdateDB(False);
    MoveCurProgressForm('Traitement')
  end;

  FiniMoveProgressForm;
  F.Close;



end;

procedure TOF_PGRETROSIMULREM.CalculRem ;
var
  TobH                  : Tob;
  Calcul                : Extended;
  MtDiff                : Extended;
  TxDiff                : Extended;
  CoefDiff              : Extended;
  BaseDiff              : Extended;
begin
  BaseDiff              := 0;
  MtDiff                := 0;
  TxDiff                := 0;
  CoefDiff              := 0;

  if (BaseRem <> GetField('PHB_BASEREM')) then
  begin
    if BaseRem <> 0 then
      BaseDiff := BaseRem - GetField('PHB_BASEREM')
    else
      BaseDiff := GetField('PHB_BASEREM');
  end
  else
    BaseDiff := BaseRem;
  if (TauxRem <> GetField('PHB_TAUXREM')) then
    TxDiff := TauxRem - GetField('PHB_TAUXREM')
  else
    TxDiff := TauxRem;
  if (CoefRem <> GetField('PHB_COEFFREM')) then
    CoefDiff := CoefRem - GetField('PHB_COEFFREM')
  else
    CoefDiff := CoefRem;

  TobH := TobHisto.FindFirst(['PGT_SALARIE'], [GetField('PHB_SALARIE')], False);
  if TobH = nil then
  begin
    TobH := Tob.Create('PGHISTRETRO', TobHisto, -1);
    TobH.PutValue('PGT_SALARIE', GetField('PHB_SALARIE'));
    TobH.PutValue('PGT_NATURERUB', 'AAA');
    TobH.PutValue('PGT_RUBRIQUE', Rubrique);
    TobH.PutValue('PGT_DATEVALIDITE', StrToDate(Datevalidite));
    TobH.PutValue('PGT_DATEDEBUT', StrToDate(Datedebut));
    TobH.PutValue('PGT_DATEFIN', StrToDate(Datefin));
    TobH.PutValue('PGT_DIFFTXSAL', 0);
    TobH.PutValue('PGT_DIFFTXPAT', 0);
    TobH.PutValue('PGT_DIFFBASECOT', 0);
    TobH.PutValue('PGT_DIFFMTSAL', 0);
    TobH.PutValue('PGT_DIFFMTPAT', 0);
    TobH.PutValue('PGT_LIBELLE1', LibRem);
    TobH.PutValue('PGT_LIBELLE2', 'Du '+Datedebut+' au '+Datefin);
    TobH.PutValue('PGT_DIFFBASEREM', BaseDiff);
    TobH.PutValue('PGT_DIFTAUXREM', TxDiff);
    TobH.PutValue('PGT_DIFCOEFFREM', CoefDiff);

    if (CodeCalcul = '01') then  //montant
      MtDiff := MtRem - GetField('PHB_MTREM');
    if (CodeCalcul = '02') then  //base*taux*coef
      MtDiff := BaseDiff * TxDiff * CoefDiff;
    if (CodeCalcul = '03') then  //base*taux/100*coef
      MtDiff := (BaseDiff * TxDiff / 100) * CoefDiff;
    if (CodeCalcul = '04') then  //base*taux
      MtDiff := BaseDiff * TxDiff;
    if (CodeCalcul = '05') then  //base*taux/100
      MtDiff := BaseDiff * TxDiff / 100;
    if (CodeCalcul = '06') then  //base*taux/coef
      MtDiff := BaseDiff * TxDiff / CoefDiff;
    if (CodeCalcul = '07') then  //base*taux/100/coef
      MtDiff := (BaseDiff * TxDiff / 100) / CoefDiff;
    if (CodeCalcul = '08') then  //base*coef
      MtDiff := GetField('PHB_BASEREM') * CoefDiff;

    TobH.PutValue('PGT_DIFFMTREM', MtDiff);
  end
  else
  begin
    //on fait les cumuls
    if (CodeCalcul = '01') then  //montant
      MtDiff := TobH.GetValue('PGT_DIFFMTREM') + (MtRem - GetField('PHB_MTREM'));
    if (CodeCalcul = '02') then  //base*taux*coef
      MtDiff := TobH.GetValue('PGT_DIFFMTREM') + ( BaseDiff * TxDiff * CoefDiff);
    if (CodeCalcul = '03') then  //base*taux/100*coef
      MtDiff := TobH.GetValue('PGT_DIFFMTREM') + ( (BaseDiff * TxDiff / 100) * CoefDiff);
    if (CodeCalcul = '04') then  //base*taux
      MtDiff := TobH.GetValue('PGT_DIFFMTREM') + (BaseDiff * TxDiff);
    if (CodeCalcul = '05') then  //base*taux/100
      MtDiff := TobH.GetValue('PGT_DIFFMTREM') + (BaseDiff * TxDiff / 100);
    if (CodeCalcul = '06') then  //base*taux/coef
      MtDiff := TobH.GetValue('PGT_DIFFMTREM') + ( BaseDiff * TxDiff / CoefDiff);
    if (CodeCalcul = '07') then  //base*taux/100/coef
      MtDiff := TobH.GetValue('PGT_DIFFMTREM') + ( (BaseDiff * TxDiff / 100) / CoefDiff);
    if (CodeCalcul = '08') then  //base*coef
      MtDiff := TobH.GetValue('PGT_DIFFMTREM') + (BaseDiff * CoefDiff);
    TobH.PutValue('PGT_DIFFMTREM', MtDiff);

    Calcul := TobH.GetValue('PGT_DIFFBASEREM') + BaseDiff;
    TobH.PutValue('PGT_DIFFBASEREM', Calcul);
  end;
end;

procedure TOF_PGRETROSIMULREM.ChangeSortie(Sender: Tobject);
begin
  if (GetCheckBoxState('CKSORTIE') = cbchecked) or (GetCheckBoxState('CKSORTIE') = cbgrayed) then
  begin
    SetControlEnabled('DATESORTIE', True);
    SetFocusControl('DATESORTIE');
  end
  else
  begin
    SetControlEnabled('DATESORTIE', False);
  end;

end;

Initialization
  registerclasses ( [ TOF_PGRETROSIMULREM ] ) ;
end.
