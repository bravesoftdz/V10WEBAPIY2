{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 28/02/2017
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTRESTITUPREFACT ()
Mots clefs ... : TOF;BTRESTITUPREFACT
*****************************************************************}
Unit BTRESTITUPREFACT_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     db,
     {$IFNDEF DBXPRESS}
     dbtables,
     {$ELSE}
     uDbxDataSet,
     {$ENDIF}
     fe_main,
     mul,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HPanel,
     HTB97,
     HEnt1,
     HMsgBox,
     Stat,
     utobview,
     uTOB,
     Paramsoc,
     LookUp,
     Vierge,
     HSysMenu,
     HRichEdt,
     HRichOLE,
     Graphics,
     Grids,
     Types,
     UTOF,
     Windows,
     uTofAfBaseCodeAffaire,
     UtilsEtat,
     Messages;

Type
  TOF_BTRESTITUPREFACT = class(TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

private
    //
    OkPlanCharge  : Boolean;
    //
    TResponsable      : THEdit;
    TDomaine          : THEdit;
    TEtablissement    : THEdit;
    LibResponsable    : THLabel;
    LibDomaine        : THLabel;
    LibEtablissement  : THLabel;
    //
    Affaire           : THEdit;
    Affaire0          : THEdit;
    Affaire1          : THEdit;
    Affaire2          : THEdit;
    Affaire3          : THEdit;
    Avenant           : THEdit;
    //
    BSelect1          : TToolbarButton97;
    BEfface           : TToolbarButton97;
    //
    SaisieDate        : THEdit;
    SaisieDate_       : THEdit;
    //
    TOBBTB            : TOB;
    //
    CodeTiers         : String;
    CodeAffaire       : String;
    CodeAffaire0      : String;
    CodeAffaire1      : String;
    CodeAffaire2      : String;
    CodeAffaire3      : String;
    CodeAvenant       : String;
    //
    AnneeMois         : String;
    AnneeMois_        : String;
    //
    DateSaisie        : TDateTime;
    DateSaisie_       : Tdatetime;
    //
    FF                : String;
    //
    //Modif FV : Ajout pour gestion du modèle d'état...
    OptionEdition : TOptionEdition;
    //
    TheType       : String;
    BParamEtat    : TToolBarButton97;
    //
    ChkApercu     : TCheckBox;
    ChkReduire    : TCheckBox;
    //
    FETAT         : THValComboBox;
    TEtat         : ThLabel;
    //
    Pages         : TPageControl;
    //
    procedure AffaireOnExit(Sender: TObject);

    procedure BRechResponsable(Sender: TObject);
    procedure bSelect1OnClick(Sender: TObject);
    //
    procedure ChargeMontantFacture(TOBLBTB : TOB; CodeAffaire : String);
    procedure ChargeMontantFactureApres(TOBLBTB: TOB; CodeAffaire: String);
    procedure ChargeMontantFactureAvant(TOBLBTB: TOB; CodeAffaire: String);
    procedure ChargeMontantMarche(TOBLBTB : TOB; CodeAffaire : String);
    procedure ChargeMontantMarcheApres(TOBLBTB: TOB; CodeAffaire: String);
    procedure ChargeMontantMarcheAvant(TOBLBTB: TOB; CodeAffaire: String);
    Procedure ChargePrevisionnelFacture;
    procedure ChargePrevisionnelFactureAvant(TOBLBTB : TOB; CodeAffaire : String);
    procedure ChargePrevisionnelFactureApres(TOBLBTB : TOB; CodeAffaire : String);
    procedure ChargePrevisionnelPlanCharge(TOBLBTB : TOB; CodeAffaire : String);
    procedure ChargePrevisionnelPlanChargeApres(TOBLBTB: TOB; CodeAffaire: String);
    procedure ChargePrevisionnelPlanChargeAvant(TOBLBTB: TOB; CodeAffaire: String);
    //
    procedure CreateTOB;
    procedure Controlechamp(Champ, Valeur: String);

    procedure DestroyTOB;

    procedure ExitResponsable(Sender: TObject);

    procedure GetObjects;

    procedure SetScreenEvents;
    procedure bEffaceOnClick(Sender: TObject);
    procedure BParamEtatClick(Sender: TObject);
    procedure ChargeEtatTableauBord;
    procedure OnChangeFEtat(Sender: TObject);


  end ;

Implementation

Uses  BTPUtil,
      AffaireUtil,
      DateUtils,
      EntGC,
      SAISUTIL,
      UtilRessource,
      TntGrids;

procedure TOF_BTRESTITUPREFACT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTRESTITUPREFACT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTRESTITUPREFACT.OnUpdate ;
Var CodeChantier : string;
    TempoDate    : string;
    Ind          : Integer;
    TOBLBTB      : TOB;
    TOBLAVANT    : TOB;
    TOBLAPRES    : TOB;
    yy, mm       : Word;
    ResteFacturer: Double;
    RestePrevoir : Double;
    SaveD1       : string;
    SaveD2       : String;
    Saveaff      : String;
    NumOrdre     : Integer;
    TraduitMois  : string;
begin
  Inherited ;

  If TOBBTB = nil then exit;

  CodeAffaire0 := Affaire0.text;
  CodeAffaire1 := Affaire1.text;
  CodeAffaire2 := Affaire2.text;
  CodeAffaire3 := Affaire3.text;
  CodeAvenant  := Avenant.text;
  NumOrdre     := 0;

  SaveD1 := AnneeMois;
  SaveD2 := AnneeMois_;

  CodeChantier := CodeAffaireRegroupe(CodeAffaire0, CodeAffaire1,CodeAffaire2,CodeAffaire3,CodeAvenant,TaModif,false,false,false);

  TOBBTB.cleardetail;

  ChargePrevisionnelFacture;

  For Ind := 0 to TobBTB.detail.count -1 do
  begin
    TOBLBTB := TOBBTB.Detail[ind];
    //
    CodeAffaire := TOBLBTB.GetValue('BPF_AFFAIRE');
    //Rupture sur l'affaire....
    if CodeAffaire <> SaveAff then
    Begin
      //On charge le montant prévu après la date de saisie...
      AnneeMois   := SaveD2;
      yy          := StrToInt(Copy(AnneeMois,0,4));
      mm          := StrToInt(Copy(AnneeMois,5,2));
      DateSaisie  := EncodeDate(yy, mm, 1);
      if yy <> 2099 then
      begin
        TOBLAPRES   := TOB.Create('LIG ANALYSE PREV', TOBBTB, -1);
        TOBLAPRES.Dupliquer(TOBBTB.Detail[Ind], True, True);
        //
        ChargePrevisionnelFactureApres(TOBLAPRES, CodeAffaire);
        ChargePrevisionnelPlanChargeApres(TOBLAPRES, CodeAffaire);
        ChargeMontantFactureApres(TOBLAPRES, CodeAffaire);
        ChargeMontantMarcheApres(TOBLAPRES, CodeAffaire);
        //
        //On découpe la date prévisionnel
        TOBLAPRES.PutValue('ANNEE', yy);
        TOBLAPRES.PutValue('MOIS',  FormatDateTime('mmmm', DateSaisie));
        TraduitMois := FormatDateTime('mmmm', DateSaisie) + ' ' + IntToStr(yy);
        TOBLAPRES.PutValue('DATEPREV', 'Après ' + TraduitMois);
        TOBLAPRES.PutValue('ORDRE', 9999);
      end;
      //
      //On charge le montant prévu avant la date de saisie...
      NumOrdre    := 0;
      AnneeMois   := SaveD1;
      //
      yy          := StrToInt(Copy(AnneeMois,0,4));
      mm          := StrToInt(Copy(AnneeMois,5,2));
      DateSaisie  := EncodeDate(yy, mm, 1);
      if yy <> 1900 then
      begin
        TOBLAVANT   := TOB.Create('LIG ANALYSE PREV', TOBBTB, -1);
        TOBLAVANT.Dupliquer(TOBLBTB, True, True);
        //
        ChargePrevisionnelFactureAvant(TOBLAVANT, CodeAffaire);
        ChargePrevisionnelPlanChargeAvant(TOBLAVANT, CodeAffaire);
        ChargeMontantFactureAvant(TOBLAVANT, CodeAffaire);
        ChargeMontantMarcheAvant(TOBLAVANT, CodeAffaire);
        //
        TraduitMois := FormatDateTime('mmmm', DateSaisie) + ' ' + IntToStr(yy);
        //On découpe la date prévisionnel
        TOBLAVANT.PutValue('ANNEE', yy);
        TOBLAVANT.PutValue('MOIS',  FormatDateTime('mmmm', DateSaisie));
        TOBLAVANT.PutValue('DATEPREV', 'Avant ' + TraduitMois);
        TOBLAVANT.PutValue('ORDRE', NumOrdre);
        //
      end;
      //
      //On pourrait ici faire les lectures indépendantes des tiers, adresse, affaire,...
      //
      //S'il faut chercher des information sur la pièce c'est ici que ça se passe !!!!
      //
      //S'il faut chercher des montants cumulés c'est aussi ici que ça se passe...
      //
      SaveAff      := TOBLBTB.GetVALUE('BPF_AFFAIRE');
    end;                                              
    //
    Inc(NumOrdre);
    //
    TempoDate   := TOBLBTB.GetValue('DATEPREV');
    yy          := StrToInt(Copy(TempoDate,0,4));
    mm          := StrToInt(Copy(TempoDate,5,2));
    DateSaisie  := EncodeDate(yy, mm, 1);
    Datesaisie_ := EndOfAMonth(yy, mm);
    //On découpe la date prévisionnel
    TOBLBTB.PutValue('ANNEE', yy);
    TOBLBTB.PutValue('MOIS',  FormatDateTime('mmmm', DateSaisie));
    //
    ChargePrevisionnelPlanCharge(TOBLBTB, CodeAffaire);
    ChargeMontantFacture(TOBLBTB, CodeAffaire);
    ChargeMontantMarche(TOBLBTB, CodeAffaire);
    //
    ResteFacturer := Valeur(TOBLBTB.GetValue('MTPREVU')) -  Valeur(TOBLBTB.GetValue('MTFACT')) ;
    TOBLBTB.PutValue('ResteAFacturer', ResteFacturer);
    //
    RestePrevoir  := Valeur(TOBLBTB.GetValue('MTMARCHE')) - Valeur(TOBLBTB.GetValue('MTPLANCHARGE'));
    TOBLBTB.PutValue('ResteAPrevoir', RestePrevoir);
    //
    TOBLBTB.PutValue('ORDRE', NumOrdre);
    //
    TraduitMois := FormatDateTime('mmmm', DateSaisie) + ' ' + IntToStr(yy);
    TOBLBTB.PutValue('DATEPREV', TraduitMois);
    //
  end;

  //On vérifie si la ligne 9999999 existe pour la dernière affaire traitée
  If TOBBTB.detail.count > 0 Then
  begin
    TOBLBTB := TOBBTB.Findfirst(['BPF_AFFAIRE','ORDRE'],[CodeAffaire, 9999], False);
    If TOBLBTB = nil then
    begin
      AnneeMois   := SaveD2;
      yy          := StrToInt(Copy(AnneeMois,0,4));
      mm          := StrToInt(Copy(AnneeMois,5,2));
      DateSaisie  := EncodeDate(yy, mm, 1);
      if yy <> 2099 then
      begin
        TOBLAPRES   := TOB.Create('LIG ANALYSE PREV', TOBBTB, -1);
        TOBLAPRES.Dupliquer(TOBBTB.Detail[TobBTB.detail.count-1], True, True);
        //
        ChargePrevisionnelFactureApres(TOBLAPRES, CodeAffaire);
        ChargePrevisionnelPlanChargeApres(TOBLAPRES, CodeAffaire);
        ChargeMontantFactureApres(TOBLAPRES, CodeAffaire);
        ChargeMontantMarcheApres(TOBLAPRES, CodeAffaire);
        //
        TraduitMois := FormatDateTime('mmmm', DateSaisie) + ' ' + IntToStr(yy);
        //On découpe la date prévisionnel
        TOBLAPRES.PutValue('ANNEE', yy);
        TOBLAPRES.PutValue('MOIS',  FormatDateTime('mmmm', DateSaisie));
        TOBLAPRES.PutValue('DATEPREV', 'Après ' + TraduitMois);
        TOBLAPRES.PutValue('ORDRE', 9999);
      end;
    end;
  End;

  AnneeMois    := SaveD1;
  AnneeMois_   := SaveD2;

  TOBBTB.Detail.sort('BPF_AFFAIRE; ORDRE');

  if TOBBTB <> nil then TFStat(Ecran).LaTOB := TOBBTB;

end ;

procedure TOF_BTRESTITUPREFACT.OnLoad ;
Var yy, mm : Word;

begin
  Inherited ;

  AnneeMois   := '';
  AnneeMois_  := '';

  yy          := StrToInt(Copy(SaisieDate.text,4,4));
  mm          := StrToInt(Copy(SaisieDate.text,0,2));
  AnneeMois   := IntToStr(yy) + Format('%.2d%', [mm]);

  yy          := StrToInt(Copy(SaisieDate_.Text,4,4));
  mm          := StrToInt(Copy(SaisieDate_.Text,0,2));
  AnneeMois_  := IntToStr(yy) + Format('%.2d%', [mm]);

end ;

procedure TOF_BTRESTITUPREFACT.OnArgument (S : String ) ;
var Critere   : string;
    Champ     : string;
    Valeur    : string;
    i         : Integer;
    x         : Integer;
    TempoDate : Tdatetime;
    yy,mm,dd  : Word;
begin
  Inherited ;

  OkPlanCharge := (VH_GC.SeriaPlanCharge);

  //gestion de l'écran
  GetObjects;
  //
  ChargeEtatTableauBord;
  //
  CreateTOB;

  //Récupération des paramètres de lancements.....
  Critere := uppercase(Trim(ReadTokenSt(S)));
  while Critere <> '' do
  begin
     x := pos('=', Critere);
     if x <> 0 then
        begin
        Champ  := copy(Critere, 1, x - 1);
        Valeur := copy(Critere, x + 1, length(Critere));
        end
     else
        Champ  := Critere;
     ControleChamp(Champ, Valeur);
     Critere:= uppercase(Trim(ReadTokenSt(S)));
  end;

  FF:='#';
  if V_PGI.OkDecV>0 then
  begin
    FF:='# ##0.';
    for i:=1 to V_PGI.OkDecP-1 do
    begin
      FF:=FF+'0';
    end;
    FF:=FF+'0';
  end;

  //On force l'affaire comme un chantier
  Affaire0.Text := 'A';

  ChargeCleAffaire(Affaire0,Affaire1,Affaire2,Affaire3,Avenant,Nil,TaModif,CodeAffaire,True);
  //
  Ecran.WindowState := wsMaximized;

  SetScreenEvents;

  TempoDate := Now;
  Decodedate(TempoDate, yy, mm, dd);
  SaisieDate.Text   := Format('%.2d%', [mm]) + '-' + IntToStr(yy);
  //
  TempoDate := idate2099;
  Decodedate(TempoDate, yy, mm, dd);
  SaisieDate_.Text  := Format('%.2d%', [mm]) + '-' + IntToStr(yy);

end ;

procedure TOF_BTRESTITUPREFACT.OnClose ;
begin
  Inherited ;

  DestroyTOB;

  FreeAndNil(OptionEdition);

end ;

procedure TOF_BTRESTITUPREFACT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTRESTITUPREFACT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTRESTITUPREFACT.CreateTOB;
begin
  TOBBTB := TOB.Create('ANALYSE PREVISONNEL', nil, -1);
end;

procedure TOF_BTRESTITUPREFACT.DestroyTOB;
begin

  FreeAndNil(TOBBTB);

end;

procedure TOF_BTRESTITUPREFACT.GetObjects;
begin

  if Assigned(GetControl('AFF_RESPONSABLE')) Then
  begin
    TResponsable                := THEdit(GetControl('AFF_RESPONSABLE'));
    TResponsable.OnElipsisClick := BRechResponsable;
    TResponsable.OnEXit         := ExitResponsable;
  end;

  If Assigned(GetControl('AFF_DOMAINE'))       Then TDomaine := THEdit(GetControl('AFF_DOMAINE'));

  If Assigned(GetControl('AFF_ETABLISSEMENT')) Then TEtablissement := THEdit(GetControl('AFF_ETABLISSEMENT'));

  if Assigned(GetControl('LIBRESSOURCE'))      Then LibResponsable   := THLabel(Getcontrol('LIBRESSOURCE'));
  if Assigned(GetControl('LIBDOMAINE'))        Then LibDomaine       := THLabel(Getcontrol('LIBDOMAINE'));
  if Assigned(GetControl('LIBETABLISSEMENT'))  Then LibEtablissement := THLabel(Getcontrol('LIBETABLISSEMENT'));
  //
  Affaire           := THEdit(GetControl('BTBAFFAIRE'));;
  Affaire0          := THEdit(GetControl('BTBAFFAIRE0'));;
  Affaire1          := THEdit(GetControl('BTBAFFAIRE1'));;
  Affaire2          := THEdit(GetControl('BTBAFFAIRE2'));;
  Affaire3          := THEdit(GetControl('BTBAFFAIRE3'));;
  Avenant           := THEdit(GetControl('BTBAVENANT'));;
  //
  SaisieDate        := THEdit(GetControl('SAISIEDATE'));;
  SaisieDate_       := THEdit(GetControl('SAISIEDATE1'));;
  //
  BSelect1          := TToolbarButton97(GetControl('BSELECTAFF1'));
  BEfface           := TToolbarButton97(GetControl('BEFFACEAFF1'));
  //
end;

procedure TOF_BTRESTITUPREFACT.SetScreenEvents;
begin

  Affaire1.OnExit  := AffaireOnExit;
  Affaire2.OnExit  := AffaireOnExit;
  Affaire3.OnExit  := AffaireOnExit;
  Avenant.OnExit   := AffaireOnExit;

  BSelect1.OnClick := bSelect1OnClick;
  BEfface.OnClick  := bEffaceOnClick;


end;

procedure TOF_BTRESTITUPREFACT.Controlechamp(Champ, Valeur: String);
begin

end;

procedure TOF_BTRESTITUPREFACT.bSelect1OnClick(Sender: TObject);
begin

  SelectionAffaire (Nil,Affaire,Affaire0 ,Affaire1,Affaire2,Affaire3,Avenant);

  AffaireOnExit(Affaire);

end;

procedure TOF_BTRESTITUPREFACT.bEffaceOnClick(Sender: TObject);
begin

  // PL le 15/01/02 pour V575
  EffaceAffaire (EditTiers,Affaire,Affaire0 ,Affaire1,Affaire2,Affaire3,Avenant);

  if (Ecran is TFMul) then TFMul(Ecran).ChercheClick;

  if (Affaire0 <> Nil) then
  begin
    if Affaire0.CanFocus then Affaire0.SetFocus;
  end;

  if (Affaire1.CanFocus) then Affaire1.SetFocus;

end;

procedure TOF_BTRESTITUPREFACT.AffaireOnExit(Sender : TObject);
Var YY,mm,dd  : Word;
    StSQL     : string;
    QQ        : TQuery;
    TempoDate : String;
begin

  try
    StSql := 'SELECT AFF_TIERS, AFF_DATEDEBUT, AFF_DATEFIN FROM AFFAIRE WHERE AFF_AFFAIRE = "' + Affaire.text + '"';
    QQ := OpenSQL(StSQL, False, -1,'',False);
    if Not QQ.eof then
    begin
      CodeTiers         := QQ.Findfield('AFF_TIERS').AsString;
      TempoDate         := QQ.Findfield('AFF_DATEDEBUT').AsString;
      //transformation en Mois/Année
      Decodedate(StrToDate(TempoDate), yy, mm, dd);
      SaisieDate.Text   := Format('%.2d%', [mm]) + '-' + IntToStr(yy);
      //
      TempoDate         := QQ.Findfield('AFF_DATEFIN').AsString;
      Decodedate(StrToDate(TempoDate), yy, mm, dd);
      SaisieDate_.Text  := Format('%.2d%', [mm]) + '-' + IntToStr(yy);
    end;
  finally
    Ferme(QQ);
  end;

end;

Procedure TOF_BTRESTITUPREFACT.ExitResponsable(Sender: TObject);
Var Lib1, Lib2 : string;
begin

  if LibResponsable = nil then Exit;

  IF TResponsable.text <> '' then
  begin
    LibelleRessource(TResponsable.text, lib1, lib2);
    LibResponsable.Visible := True;
    LibResponsable.caption := Lib1 + ' ' + Lib2;
  end
  else
    LibResponsable.Visible := False;

end;

procedure TOF_BTRESTITUPREFACT.BRechResponsable(Sender: TObject);
Var Lib1, Lib2 : String;
begin

  GetRessourceRecherche(TResponsable,'ARS_TYPERESSOURCE="SAL"', '', '');

  if LibResponsable = nil then Exit;

  IF TResponsable.text <> '' then
  begin
    LibelleRessource(TResponsable.text, lib1, lib2);
    LibResponsable.Visible := True;
    LibResponsable.caption := Lib1 + ' ' + Lib2;
  end
  else
    LibResponsable.Visible := False;

end;

{***********A.G.L.***********************************************
Auteur  ...... : F.Vautrain
Créé le ...... : 01/03/2017
Modifié le ... :   /  /
Description .. : Gestion du Montant Prévisionnel Facturation
Mots clefs ... :
*****************************************************************}
procedure TOF_BTRESTITUPREFACT.ChargePrevisionnelFactureAvant(TOBLBTB : TOB; CodeAffaire : String);
Var Req : string;
    QQ  : TQuery;
    MtPrevu : Double;
begin

  MtPrevu := 0;

  Try
    Req := 'SELECT SUM(BPF_MTPREVU) As AVANTPREVU FROM BTPREVFAC ';
    Req := Req + '  WHERE BPF_MOIS < "' + AnneeMois + '" ';
    Req := req + '    AND BPF_AFFAIRE="' + CodeAffaire + '" ';

    QQ := OpenSQL(Req, False, -1, '', True);
    if Not QQ.EOF then
    begin
      MtPrevu := QQ.Findfield('AVANTPREVU').AsFloat;
    end;
  finally
    Ferme(QQ);
  end;

  TOBLBTB.PutValue('MTPREVU', MtPrevu);

end;

procedure TOF_BTRESTITUPREFACT.ChargePrevisionnelFacture;
Var Req : string;
    QQ  : TQuery;
begin

  Try
    Req := 'SELECT BPF_AFFAIRE, BPF_MOIS AS DATEPREV, BPF_MTPREVU AS MTPREVU, BPF_REMARQUE, BPF_VERROU, ';
    Req := Req + ' AFF_AFFAIRE, AFF_AFFAIRE0, AFF_AFFAIRE1, AFF_AFFAIRE2, AFF_AFFAIRE3, AFF_AVENANT, ';
    Req := Req + ' AFF_LIBELLE, AFF_DOMAINE, AFF_RESPONSABLE, AFF_DATEDEBUT, AFF_DATEFIN, AFF_APPORTEUR, ';
    Req := Req + ' AFF_VALLIBRE1, AFF_VALLIBRE2, AFF_VALLIBRE3, ';
    Req := Req + ' AFF_DATELIBRE1, AFF_DATELIBRE2, AFF_DATELIBRE3, ';
    Req := Req + ' AFF_LIBREAFF1, AFF_LIBREAFF2, AFF_LIBREAFF3, AFF_LIBREAFF4, AFF_LIBREAFF5, AFF_LIBREAFF6, AFF_LIBREAFF7, AFF_LIBREAFF8, AFF_LIBREAFF9, AFF_LIBREAFFA, ';
    Req := Req + ' T_TIERS, T_LIBELLE, T_VILLE, T_CODEPOSTAL, 0 AS ORDRE, ';
    Req := Req + ' 0.00 AS MTPLANCHARGE, 0.00 AS MTFACT, 0.00 AS MTMARCHE, 0.00 AS RESTEAFACTURER, 0.00 AS RESTEAPREVOIR,';
    Req := Req + ' ""   AS MOIS, "" AS ANNEE ';
    //Req := Req + ' 0 AS MTPLANCHARGEAVANT, 0 AS MTFACTAVANT, 0 AS MTMARCHEAVANT, 0 AS MTPREVUAVANT, ';
    //Req := Req + ' 0 AS MTPLANCHARGEAPRES, 0 AS MTFACTAPRES, 0 AS MTMARCHEAPRES, 0 AS MTPREVUAPRES ';
    Req := Req + '   FROM BTPREVFAC ';
    Req := Req + '   LEFT JOIN AFFAIRE  ON AFF_AFFAIRE=BPF_AFFAIRE';
    Req := Req + '   LEFT JOIN TIERS    ON T_TIERS=AFF_TIERS';
    Req := Req + '  WHERE BPF_MOIS BETWEEN "' + AnneeMois + '" ';
    Req := Req + '    AND "' + AnneeMois_ + '"';

    if Affaire.text <> '' then Req := req + '    AND BPF_AFFAIRE="' + Affaire.Text+ '" ';

    if Affaire.text = '' then
      Req := Req + ' ORDER BY BPF_AFFAIRE, BPF_MOIS'
    else
      Req := Req + ' ORDER BY BPF_MOIS';

    QQ := OpenSQL(Req, False, -1, '', True);
    if Not QQ.EOF then
    begin
      TOBBTB.LoadDetailDB('LIG ANALYSE PREV','','',QQ,False);
    end;
  finally
    Ferme(QQ);
  end;


end;

procedure TOF_BTRESTITUPREFACT.ChargePrevisionnelFactureApres(TOBLBTB : TOB; CodeAffaire : String);
Var Req : string;
    QQ  : TQuery;
    MtPrevu : double;
begin

  MtPrevu := 0;

  Try
    //
    Req := 'SELECT SUM(BPF_MTPREVU) As APRESPREVU FROM BTPREVFAC ';
    Req := Req + '  WHERE BPF_MOIS > "' + AnneeMois + '" ';
    Req := req + '    AND BPF_AFFAIRE="' + CodeAffaire + '" ';

    QQ := OpenSQL(Req, False, -1, '', True);
    if Not QQ.EOF then
    begin
      MtPrevu := QQ.Findfield('APRESPREVU').AsFloat;
    end;
  finally
    Ferme(QQ);
  end;

  TOBLBTB.PutValue('MTPREVU',  MtPrevu);

end;

{***********A.G.L.***********************************************
Auteur  ...... : F.Vautrain
Créé le ...... : 01/03/2017
Modifié le ... :   /  /
Description .. : Gestion du Montant Facturé
Mots clefs ... :
*****************************************************************}
procedure TOF_BTRESTITUPREFACT.ChargeMontantFactureAvant(TOBLBTB : TOB; CodeAffaire : String);
Var StSql           : string;
    QQ              : Tquery;
    MtFacture       : Double;
begin

  MtFacture := 0;

  try
    StSQl := 'SELECT SUM(GP_TOTALHTDEV) AS MTFACTAVANT FROM PIECE ';
    StSQL := StSQL + 'LEFT JOIN AFFAIRE AS AA on GP_AFFAIRE = AA.AFF_AFFAIRE ';
    STSQL := StSQL + 'LEFT JOIN AFFAIRE AS AD on GP_AFFAIREDEVIS = AD.AFF_AFFAIRE ';
    StSQl := StSQl + 'WHERE GP_AFFAIRE = "'   + CodeAffaire + '" ';
    StSQl := StSQl + '  AND GP_NATUREPIECEG = "FBT" ';
    StSQl := StSQl + '  AND GP_DATEPIECE <"' + USDATETIME(DateSaisie)   + '" ';
    StSQl := StSQl + '  AND AD.AFF_ETATAFFAIRE = "ACP" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then
    Begin
      MtFacture := QQ.findField('MTFACTAVANT').AsFloat;
    end;
  finally
    Ferme(QQ);
  end;

  TOBLBTB.PutValue('MTFACT',  MtFacture);

end;

procedure TOF_BTRESTITUPREFACT.ChargeMontantFacture(TOBLBTB : TOB; CodeAffaire : String);
Var StSql           : string;
    QQ              : Tquery;
    MtFacture       : Double;
begin

  MtFacture := 0;

  try
    StSQl := 'SELECT SUM(GP_TOTALHTDEV) AS MTFACT FROM PIECE ';
    StSQL := StSQL + 'LEFT JOIN AFFAIRE AS AA on GP_AFFAIRE = AA.AFF_AFFAIRE ';
    STSQL := StSQL + 'LEFT JOIN AFFAIRE AS AD on GP_AFFAIREDEVIS = AD.AFF_AFFAIRE ';
    StSQl := StSQl + 'WHERE GP_AFFAIRE = "'   + CodeAffaire + '" ';
    StSQl := StSQl + '  AND GP_NATUREPIECEG = "FBT" ';
    StSQl := StSQl + '  AND GP_DATEPIECE >="' + USDATETIME(DateSaisie)   + '" ';
    StSQl := StSQl + '  AND GP_DATEPIECE <="' + USDATETIME(DateSaisie_)   + '" ';
    StSQl := StSQl + '  AND AD.AFF_ETATAFFAIRE = "ACP" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then
    Begin
      MtFacture := QQ.findField('MTFACT').AsFloat;
    end;
  finally
    Ferme(QQ);
  end;

  TOBLBTB.PutValue('MTFACT',  MtFacture);

end;

procedure TOF_BTRESTITUPREFACT.ChargeMontantFactureApres(TOBLBTB : TOB; CodeAffaire : String);
Var StSql           : string;
    QQ              : Tquery;
    MtFacture       : Double;
begin

  MtFacture := 0;

  try
    StSQl := 'SELECT SUM(GP_TOTALHTDEV) AS MTFACTAPRES FROM PIECE ';
    StSQL := StSQL + 'LEFT JOIN AFFAIRE AS AA on GP_AFFAIRE = AA.AFF_AFFAIRE ';
    STSQL := StSQL + 'LEFT JOIN AFFAIRE AS AD on GP_AFFAIREDEVIS = AD.AFF_AFFAIRE ';
    StSQl := StSQl + 'WHERE GP_AFFAIRE = "'   + CodeAffaire + '" ';
    StSQl := StSQl + '  AND GP_NATUREPIECEG = "FBT" ';
    StSQl := StSQl + '  AND GP_DATEPIECE >"' + USDATETIME(DateSaisie)   + '" ';
    StSQl := StSQl + '  AND AD.AFF_ETATAFFAIRE = "ACP" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then
    Begin
      MtFacture := QQ.findField('MTFACTAPRES').AsFloat;
    end;
  finally
    Ferme(QQ);
  end;

  TOBLBTB.PutValue('MTFACT',  MtFacture);

end;
{***********A.G.L.***********************************************
Auteur  ...... : F.Vautrain
Créé le ...... : 01/03/2017
Modifié le ... :   /  /
Description .. : Gestion du Montant Marché
Mots clefs ... :
*****************************************************************}
procedure TOF_BTRESTITUPREFACT.ChargeMontantMarcheAvant(TOBLBTB : TOB; CodeAffaire : String);
Var StSql           : string;
    QQ              : Tquery;
    MtMarche        : double;
begin

  MtMarche := 0;

  try
    StSQl := 'SELECT SUM(GP_TOTALHTDEV) AS MTMARCHEAVANT FROM PIECE ';
    StSQL := StSQL + 'LEFT JOIN AFFAIRE AS AA on GP_AFFAIRE = AA.AFF_AFFAIRE ';
    STSQL := StSQL + 'LEFT JOIN AFFAIRE AS AD on GP_AFFAIREDEVIS = AD.AFF_AFFAIRE ';
    StSQl := StSQl + 'WHERE GP_AFFAIRE = "'   + CodeAffaire + '" ';
    StSQl := StSQl + '  AND GP_NATUREPIECEG = "DBT" ';
    StSQl := StSQl + '  AND GP_DATEPIECE <"' + USDATETIME(DateSaisie)   + '" ';
    StSQl := StSQl + '  AND AD.AFF_ETATAFFAIRE = "ACP" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then
    Begin
      MtMarche := QQ.findField('MTMARCHEAVANT').AsFloat;
    end;
  finally
    Ferme(QQ);
  end;

  TOBLBTB.PutValue('MTMARCHE',  MtMarche);

end;

procedure TOF_BTRESTITUPREFACT.ChargeMontantMarche(TOBLBTB : TOB; CodeAffaire : String);
Var StSql           : string;
    QQ              : Tquery;
    MtMarche        : double;
begin

  MtMarche := 0;

  try
    StSQl := 'SELECT SUM(GP_TOTALHTDEV) AS MTMARCHE FROM PIECE ';
    StSQL := StSQL + 'LEFT JOIN AFFAIRE AS AA on GP_AFFAIRE = AA.AFF_AFFAIRE ';
    STSQL := StSQL + 'LEFT JOIN AFFAIRE AS AD on GP_AFFAIREDEVIS = AD.AFF_AFFAIRE ';
    StSQl := StSQl + 'WHERE GP_AFFAIRE = "'   + CodeAffaire + '" ';
    StSQl := StSQl + '  AND GP_NATUREPIECEG = "DBT" ';
    StSQl := StSQl + '  AND GP_DATEPIECE >="' + USDATETIME(DateSaisie)   + '" ';
    StSQl := StSQl + '  AND GP_DATEPIECE <="' + USDATETIME(DateSaisie_)   + '" ';
    StSQl := StSQl + '  AND AD.AFF_ETATAFFAIRE = "ACP" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then
    Begin
      MtMarche := QQ.findField('MTMARCHE').AsFloat;
    end;
  finally
    Ferme(QQ);
  end;

  TOBLBTB.PutValue('MTMARCHE',  MtMarche);

end;

procedure TOF_BTRESTITUPREFACT.ChargeMontantMarcheApres(TOBLBTB : TOB; CodeAffaire : String);
Var StSql           : string;
    QQ              : Tquery;
    MtMarche        : double;
begin

  MtMarche := 0;

  try
    StSQl := 'SELECT SUM(GP_TOTALHTDEV) AS MTMARCHEAPRES FROM PIECE ';
    StSQL := StSQL + 'LEFT JOIN AFFAIRE AS AA on GP_AFFAIRE = AA.AFF_AFFAIRE ';
    STSQL := StSQL + 'LEFT JOIN AFFAIRE AS AD on GP_AFFAIREDEVIS = AD.AFF_AFFAIRE ';
    StSQl := StSQl + 'WHERE GP_AFFAIRE = "'   + CodeAffaire + '" ';
    StSQl := StSQl + '  AND GP_NATUREPIECEG = "DBT" ';
    StSQl := StSQl + '  AND GP_DATEPIECE >"' + USDATETIME(DateSaisie)   + '" ';
    StSQl := StSQl + '  AND AD.AFF_ETATAFFAIRE = "ACP" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then
    Begin
      MtMarche := QQ.findField('MTMARCHEAPRES').AsFloat;
    end;
  finally
    Ferme(QQ);
  end;

  TOBLBTB.PutValue('MTMARCHE',  MtMarche);

end;

{***********A.G.L.***********************************************
Auteur  ...... : F.Vautrain
Créé le ...... : 01/03/2017
Modifié le ... :   /  /
Description .. : Gestion du Montant Previsionnel Plan de charge
Mots clefs ... :
*****************************************************************}
procedure TOF_BTRESTITUPREFACT.ChargePrevisionnelPlanChargeAvant(TOBLBTB : TOB; CodeAffaire : String);
Var StSQL     : string;
    QQ        : TQuery;
    Ratio     : Double;
    NbheurePC : Double;
    NbheureAff: Double;
    MTDevisACC: Double;
    MtPlCharge: Double;
begin

  if not OkPlanCharge then exit;

  MtPlCharge  := 0;

  NbheurePC   := 0;
  NbheureAff  := 0;
  MTDevisACC  := 0;
  Ratio       := 0;

  //On calcul d'abord le ratio du nombre d'heure
  try
    StSQl := 'SELECT SUM(BAT_NBHRS) AS HEUREPC FROM BAFFECTCHANT ';
    StSQl := StSQl + 'WHERE BAT_AFFAIRE = "' + CodeAffaire + '" ';
    StSQl := StSQl + '  AND BAT_DATE <"'     + USDATETIME(DateSaisie) + '" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then NbheurePC := QQ.findField('HEUREPC').AsFloat;
  finally
    Ferme(QQ);
  end;

  //On calcul ensuite le nombre d'heures totales de l'affaire, le montant total des devis accepté sur l'affaire...
  try
    StSQl := 'SELECT SUM(GP_TOTALHEURE) AS HEUREAFF, SUM(GP_TOTALHTDEV) AS MONTANTHT FROM PIECE ';
    StSQL := StSQL + 'LEFT JOIN AFFAIRE AS AA on GP_AFFAIRE = AA.AFF_AFFAIRE ';
    STSQL := StSQL + 'LEFT JOIN AFFAIRE AS AD on GP_AFFAIREDEVIS = AD.AFF_AFFAIRE ';
    StSQl := StSQl + 'WHERE GP_AFFAIRE = "'   + CodeAffaire + '" ';
    StSQl := StSQl + '  AND GP_NATUREPIECEG = "DBT" ';
    StSQl := StSQl + '  AND GP_DATEPIECE <"'  + USDATETIME(DateSaisie)   + '" ';
    StSQl := StSQl + '  AND AD.AFF_ETATAFFAIRE ="ACP" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then
    Begin
      NbheureAff := QQ.findField('HEUREAFF').AsFloat;
      MTDevisACC := QQ.findField('MONTANTHT').AsFloat;
    end;
  finally
    Ferme(QQ);
  end;

  //Calcul du Ratio du nombre d'heure
  if NbheureAff <> 0 then Ratio := NbheurePC/NbheureAff;

  MtPlCharge := Ratio * MtDevisAcc;

  TOBLBTB.PutValue('MTPLANCHARGE', MtPlCharge);

end;

procedure TOF_BTRESTITUPREFACT.ChargePrevisionnelPlanCharge(TOBLBTB : TOB; CodeAffaire : String);
Var StSQL     : string;
    QQ        : TQuery;
    Ratio     : Double;
    NbheurePC : Double;
    NbheureAff: Double;
    MTDevisACC: Double;
    MtPlCharge: Double;
begin

  if not OkPlanCharge then exit;

  NbheurePC   := 0;
  NbheureAff  := 0;
  MTDevisACC  := 0;
  Ratio       := 0;

  //On calcul d'abord le ratio du nombre d'heure
  try
    StSQl := 'SELECT SUM(BAT_NBHRS) AS HEUREPC FROM BAFFECTCHANT ';
    StSQl := StSQl + 'WHERE BAT_AFFAIRE = "'  + CodeAffaire        + '" ';
    StSQl := StSQl + '  AND BAT_DATE >="'     + USDATETIME(DateSaisie)   + '" ';
    StSQl := StSQl + '  AND BAT_DATE <="'     + USDATETIME(DateSaisie_)   + '" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then NbheurePC := QQ.findField('HEUREPC').AsFloat;
  finally
    Ferme(QQ);
  end;

  //On calcul ensuite le nombre d'heures totales de l'affaire, le montant total des devis accepté sur l'affaire...
  try
    StSQl := 'SELECT SUM(GP_TOTALHEURE) AS HEUREAFF, SUM(GP_TOTALHTDEV) AS MONTANTHT FROM PIECE ';
    StSQL := StSQL + 'LEFT JOIN AFFAIRE AS AA on GP_AFFAIRE = AA.AFF_AFFAIRE ';
    STSQL := StSQL + 'LEFT JOIN AFFAIRE AS AD on GP_AFFAIREDEVIS = AD.AFF_AFFAIRE ';
    StSQl := StSQl + 'WHERE GP_AFFAIRE = "'   + CodeAffaire + '" ';
    StSQl := StSQl + '  AND GP_NATUREPIECEG = "DBT" ';
    StSQl := StSQl + '  AND GP_DATEPIECE >="' + USDATETIME(DateSaisie)   + '" ';
    StSQl := StSQl + '  AND GP_DATEPIECE <="' + USDATETIME(DateSaisie_)   + '" ';
    StSQl := StSQl + '  AND AD.AFF_ETATAFFAIRE ="ACP" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then
    Begin
      NbheureAff := QQ.findField('HEUREAFF').AsFloat;
      MTDevisACC := QQ.findField('MONTANTHT').AsFloat;
    end;
  finally
    Ferme(QQ);
  end;

  //Calcul du Ratio du nombre d'heure
  if NbheureAff <> 0 then Ratio := NbheurePC/NbheureAff;

  MtPlCharge := Ratio * MtDevisAcc;

  TOBLBTB.PutValue('MTPLANCHARGE', MtPlCharge);

end;

procedure TOF_BTRESTITUPREFACT.ChargePrevisionnelPlanChargeApres(TOBLBTB : TOB; CodeAffaire : String);
Var StSQL     : string;
    QQ        : TQuery;
    Ratio     : Double;
    NbheurePC : Double;
    NbheureAff: Double;
    MTDevisACC: Double;
    MtPlCharge: Double;
begin

  if not OkPlanCharge then exit;

  MtPlCharge := 0;

  NbheurePC   := 0;
  NbheureAff  := 0;
  MTDevisACC  := 0;
  Ratio       := 0;

  //On calcul d'abord le ratio du nombre d'heure
  try
    StSQl := 'SELECT SUM(BAT_NBHRS) AS HEUREPC FROM BAFFECTCHANT ';
    StSQl := StSQl + 'WHERE BAT_AFFAIRE = "' + CodeAffaire + '" ';
    StSQl := StSQl + '  AND BAT_DATE >"'     + USDATETIME(DateSaisie) + '" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then NbheurePC := QQ.findField('HEUREPC').AsFloat;
  finally
    Ferme(QQ);
  end;

  //On calcul ensuite le nombre d'heures totales de l'affaire, le montant total des devis accepté sur l'affaire...
  try
    StSQl := 'SELECT SUM(GP_TOTALHEURE) AS HEUREAFF, SUM(GP_TOTALHTDEV) AS MONTANTHT FROM PIECE ';
    StSQL := StSQL + 'LEFT JOIN AFFAIRE AS AA on GP_AFFAIRE = AA.AFF_AFFAIRE ';
    STSQL := StSQL + 'LEFT JOIN AFFAIRE AS AD on GP_AFFAIREDEVIS = AD.AFF_AFFAIRE ';
    StSQl := StSQl + 'WHERE GP_AFFAIRE = "'   + CodeAffaire + '" ';
    StSQl := StSQl + '  AND GP_NATUREPIECEG = "DBT" ';
    StSQl := StSQl + '  AND GP_DATEPIECE >"'  + USDATETIME(DateSaisie)   + '" ';
    StSQl := StSQl + '  AND AD.AFF_ETATAFFAIRE ="ACP" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then
    Begin
      NbheureAff := QQ.findField('HEUREAFF').AsFloat;
      MTDevisACC := QQ.findField('MONTANTHT').AsFloat;
    end;
  finally
    Ferme(QQ);
  end;

  //Calcul du Ratio du nombre d'heure
  if NbheureAff <> 0 then Ratio := NbheurePC/NbheureAff;

  MtPlCharge := Ratio * MtDevisAcc;

  TOBLBTB.PutValue('MTPLANCHARGE', MtPlCharge);

end;

procedure TOF_BTRESTITUPREFACT.ChargeEtatTableauBord;
Var Idef    : Integer;
    Modele  : string;
begin

  TheType   := 'E';

  Modele         := GetParamSocSecur('SO_RECAPPREVFAC', 'BP3');

  BParamEtat     := TToolBarButton97(GetControl('BParamEtat'));
  ChkApercu      := TCheckBox(GetControl('FApercu'));
  ChkReduire     := TCheckBox(GetControl('FReduire'));
  //
  FETAT          := THValComboBox(GetControl('FEtat'));
  TEtat          := ThLabel(GetControl('TEtat'));
  //
  Pages          := TPageControl(GetControl('Pages'));
  //
  OptionEdition := TOptionEdition.Create(TheType,'BPF', Modele, Ecran.Caption, '', ChkApercu.Checked, ChkReduire.Checked, True, False, False, Pages, fEtat);
  //
  FEtat.OnChange := OnChangeFEtat;
  //FETAT.Plus     := 'MO_TYPE="E" AND MO_NATURE="BPF" AND MO_CODE<>"PFA"';
  FEtat.Value    := Modele;

  BParamEtat.OnClick  := BParamEtatClick;
  BParamEtat.Visible  := BParamEtat.Visible and JaiLeDroitConcept(ccParamEtat,False);

  OptionEdition.first := True;
  OptionEdition.ChargeListeEtat(fEtat, Idef);

end;


Procedure TOF_BTRESTITUPREFACT.OnChangeFEtat(Sender : TObject);
begin

  TFStat(Ecran).CodeEtat  := FEtat.Value;
  OptionEdition.Modele    := FEtat.Value;

end;

procedure TOF_BTRESTITUPREFACT.BParamEtatClick(Sender: TObject);
begin

  OptionEdition.Appel_Generateur

end;



Initialization
  registerclasses ( [ TOF_BTRESTITUPREFACT ] ) ;
end.

