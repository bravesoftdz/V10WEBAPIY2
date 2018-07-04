{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPCHANCELL ()
Mots clefs ... : TOF;CPCHANCELL
*****************************************************************}
Unit CPCHANCELL_TOF ;

//================================================================================
// Interface
//================================================================================
Interface

Uses
    StdCtrls, Controls, Classes, HEnt1, forms, sysutils, ComCtrls, buttons, HCtrls, Ent1,
    HMsgBox,
{$IFDEF EAGLCLIENT}
    MaineAGL,
{$ELSE}
    FE_Main,
    db,
{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
    Purge,
{$ENDIF}
    UTOF,
    UTOB,
    Windows,         // VK_DOWN
    AGLInit,         // ActionToString
    ParamSoc,        // GetparamSoc
    UTOFGRILLEFILTRE
    ;

{$IFDEF EAGLCLIENT}
    type tquery = TOB;
{$ENDIF}

//==================================================
// Externe
//==================================================
Procedure FicheChancel(Quel : String ; JusteUne : Boolean ; DateDeb : TDateTime ; Action : TActionFiche ; SurEuro : boolean);
Procedure FicheChancelsur2dates(Quel : String ; DateDeb,DateFin : TDateTime ; Action : TActionFiche ; SurEuro : boolean) ;

//==================================================
// Definition de class
//==================================================
Type
    TOF_CPCHANCELL = Class (TOF_GRILLEFILTRE)
    public
        procedure OnLoadRecord(vTOBLignes : TOB)                ; override ;
        function  OnUpdateRecord(laTOB : TOB)       : Integer   ; override ;
        function  OnAfterUpdateRecord(laTOB : TOB)  : Integer   ; override ;
        function  OnDeleteRecord(laTOB : TOB)       : Integer   ; override ;
        procedure ParametrageFiche                              ; override ;
        procedure ParametrageGrille(vGrille : THGrid)           ; override ;
        function  GenererConditionPlus              : String    ; override ;
        procedure OnCreateTOB ( vTOB : TOB )                    ; override ;
        procedure FListeOnExit(Sender: TObject)                 ; override ; {JP 18/05/06 : FQ 17546}
        procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState) ; override; {JP 20/08/03}
        procedure OnNew                    ; override ;
        procedure OnDelete                 ; override ;
        procedure OnUpdate                 ; override ;
        procedure OnLoad                   ; override ;
        procedure OnArgument (S : String ) ; override ;
        procedure OnDisplay                ; override ;
        procedure OnClose                  ; override ;
        procedure OnCancel                 ; override ;
    private
        TValable,
        TQuotite,
        TValeurQ    : THLabel;
        BPurge      : TSpeedButton;
        RGSens      : THRadioGroup;
        DateD,DateF : THEdit;
        OkDate      : TCheckBox;
        Combo       : THValComboBox;
        FSur2Dates  : Boolean;

        {JP 20/05/03 : Date de saisie de l'écriture}
        Datedeb : TDateTime;
        DateFin : TDateTime;

        SurEuro : boolean;
        Action : TActionFiche;
        devise : string;

        procedure OnClickBPurge (Sender: TObject);
        procedure OnClickRGSens (Sender: TObject);
        procedure OnChangeOkDate(Sender: TObject);
        {JP 20/08/03 : Pour gérer la conversion Cotation/Taux}
        procedure FListeCellExit(Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);

        {JP 20/08/03 : Pour remettre à jour l'écran si l'on change de devise}
        procedure OnDeviseChange(Sender : TObject);

        {JP 18/08/04 : FQ 13909/10 : gestion  de la cohérence des dates en fonction des options autour de l'euro}
        procedure VerifCoherenceDate;

        function  QuotiteChange(devise : string) : Integer;
        {JP 18/05/06 : FQ 17546 : Je sors le traitement de FListeCellExit afin de  pouvoir l'appeler
                       sur d'autres évènements}
        procedure MajQuotiteTaux(Col, Row : Integer);
    public
      {JP 18/08/04 : On s'assure que les critères de recherche sont valides : par défaut renvoie True}
      function  IsCritereValide : Boolean; override;
      procedure OnDateChange(Sender : TObject);
      procedure OnDateExit  (Sender : TObject);
    end;

//================================================================================
// Implementation
//================================================================================
Implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  HPanel;

//==================================================
// Definition des Constante
//==================================================
const
     msc0  = '0;Chancellerie;Voulez-vous enregistrer les modifications?;Q;YNC;Y;C;';
     msc1  = '1;Chancellerie;Confirmez-vous la suppression de l''enregistrement?;Q;YNC;N;C;';
     msc2  = '2;Chancellerie;Vous devez renseigner un code.;W;O;O;O;';
     msc3  = '3;Chancellerie;Vous devez renseigner un libellé.;W;O;O;O;';
     msc4  = '4;Chancellerie;La date de valeur que vous avez saisie est déjà renseignée.;W;O;O;O;';
     msc5  = '5;Chancellerie;Vous devez renseigner une date de valeur.;W;O;O;O;';
     msc6  = '6;Chancellerie;Vous devez renseigner un taux de change pour chaque date de valeur.;W;O;O;O;';
     msc7  = 'L''enregistrement est inaccessible';
     msc8  = '8;Chancellerie;L''enregistrement n''a pas pu être sauvegardé. La date est vide ou elle existe déjà.;W;O;O;O;';
     msc9  = '9;Chancellerie;Vous devez renseigner un taux de change supérieur à zéro.;W;O;O;O;';
     msc10 = 'Valable jusqu''au';
     msc11 = 'Valable à partir du';
     msc12 = 'par rapport à la devise pivot';
     msc13 = 'par rapport à l''euro';
     msc14 = '14;Chancellerie;La date saisie est supérieure à la date d''entrée en vigueur de l''Euro.;W;O;O;O;';
     msc15 = '15;Chancellerie;La date saisie est inférieure à la date d''entrée en vigueur de l''Euro.;W;O;O;O;';
     msc16 = '1 $1 = xx,xxx $2';
     msc17 = '1 $2 = xx,xxx $1';

//==================================================
// fonctions hors class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 18/03/2003
Modifié le ... : 18/03/2003
Description .. : Point d'entrée
Mots clefs ... :
*****************************************************************}
Procedure FicheChancel(Quel : String ; JusteUne : Boolean ; DateDeb : TDateTime ; Action : TActionFiche ; SurEuro : boolean);
var
  Ok : string;
begin
  {Ce Boolean ne servant à rien, je l'utilise pour savoir si le combo de choix des devises
   sera actif ou non}
  if JusteUne then OK := ';OUI;'
              else OK := ';NON;';
  {JP 20/08/03 : Utilisation de la fonction ActionToString}
  if (SurEuro) then AGLLanceFiche('CP','CPCHANCELL','',Quel,ActionToString(Action) + ';' + Quel + ';SUREURO;' + DateToStr(DateDeb) + Ok)
  else AGLLanceFiche('CP','CPCHANCELL','',Quel,'ACTION=' + ActionToString(Action) + ';' + Quel + ';PASEURO;' + DateToStr(DateDeb) + Ok);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/06/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure FicheChancelsur2dates(Quel : String ; DateDeb,DateFin : TDateTime ; Action : TActionFiche ; SurEuro : boolean) ;
var lStSurEuro : string;
begin
  if SurEuro then lStSurEuro := 'SUREURO'
  else lStSurEuro := 'PASEURO';

  AGLLanceFiche('CP','CPCHANCELL','',Quel,ActionToString(Action) + ';' + Quel + ';' + lStSurEuro + ';' + DateToStr(DateDeb) + ';NON;' + 'TRUE;' + DateToStr(DateFin));
end;

//==================================================
// Evenements par default de la TOF_GRILLECOMBO
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCHANCELL.OnLoadRecord(vTOBLignes : TOB);
var
  st1,st2 : string;
  Quotite : Integer;
begin
    // recup de la devise et init du caption
    {$IFDEF TRESO}
    Devise := combo.Value;
    {$ELSE}
    if (vTOBLignes.Detail.count > 0) then
      Devise := vTOBLignes.Detail[0].GetValue('H_DEVISE')
    else
      Devise := '';  
    {$ENDIF}
    Ecran.Caption:=TraduireMemoire(Ecran.Caption) ;
    InitCaption(Ecran,combo.Value,combo.text);

    // change la quotité et la monnaieIn
    if (devise <> '') then Quotite := QuotiteChange(devise)
                      else Quotite := 0;

    // set de visibilité
    TQuotite.Visible := (Quotite <> 1);
    TValeurQ.Visible := (Quotite <> 1);

    if (devise = V_PGI.DevisePivot) then
    begin
        RGSens.ItemIndex := 1;
        RGSens.Visible := false;
    end
    else
    begin
        RGSens.ItemIndex := 0;
        RGSens.Visible := true;

        if ((SurEuro)) then
        begin
            st1 := FindEtReplace(msc16,'$1',V_PGI.DevisePivot,true);
            st2 := FindEtReplace(msc17,'$1',V_PGI.DevisePivot,true);
        end
        else
        begin
            st1 := FindEtReplace(msc16,'$1',V_PGI.DeviseFongible,true);
            st2 := FindEtReplace(msc17,'$1',V_PGI.DeviseFongible,true);
        end;

        RGsens.Items[0] := FindEtReplace(st1,'$2',devise,true);
        RGsens.Items[1] := FindEtReplace(st2,'$2',devise,true);
    end;


    {$IFDEF TRESO}
    TValable.Visible := False;
    TValeurQ.Visible := False;
    TQuotite.Caption := 'La quotité est de 1 Euro ' + ' pour ' + IntToStr(Quotite) + ' ' + combo.text;
    {$ELSE}
    TValable.Visible := VH^.TenueEuro;    
    // st des text
    if (SurEuro) then
      TValable.Caption := TraduireMemoire(msc11) + '  ' + DateToStr(V_PGI.DateDebutEuro) + '  ' + TraduireMemoire(msc13)
    else
      TValable.Caption := TraduireMemoire(msc10) + '  ' + DateToStr(V_PGI.DateDebutEuro-1)+'  ' + TraduireMemoire(msc12) ;
    TValeurQ.Caption := IntToStr(Quotite);
    {$ENDIF}


    // enable des bouton
    if (Action = taConsult) then BPurge.Enabled := false;

    // sens
    OnClickRGSens(nil);

end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPCHANCELL.OnUpdateRecord(laTOB : TOB) : Integer;
begin
    result := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPCHANCELL.OnAfterUpdateRecord(laTOB : TOB) : Integer;
begin
    result := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPCHANCELL.OnDeleteRecord(laTOB : TOB) : Integer;
begin
    result := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCHANCELL.ParametrageFiche;
begin
    // init de champs du parent
    NomTable        := 'CHANCELL';
    ChampTri        := 'H_DATECOURS';
    ListeChamps     := 'H_DATECOURS;H_COTATION;H_TAUXREEL;H_TAUXCLOTURE;H_COMMENTAIRE';
    CodeEtat        := '';
    NatureEtat      := '';
    TitreEcran      := 'Chancellerie : ';
    ListeColsUniq   := '';
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCHANCELL.ParametrageGrille(vGrille : THGrid);
begin
    // gestion du sens :)
    RGSens.ItemIndex := 0;
    OnClickRGSens(nil);

//    ListeChamps     := 'H_DATECOURS;H_COTATION;H_TAUXREEL;H_TAUXCLOTURE;H_COMMENTAIRE';

    // type des champs
    vGrille.ColTypes[1] := 'D';
    vGrille.ColTypes[2] := 'R';
    vGrille.ColTypes[3] := 'R';
    vGrille.ColTypes[4] := 'R';

    // format du contenu des colonnes
    vGrille.ColAligns[2] := taRightJustify ;
    vGrille.ColAligns[3] := taRightJustify ;
    vGrille.ColAligns[4] := taRightJustify ;

    // FQ17720 : tauxreel base sur 9 décimales, cotation sur 6
    vGrille.ColFormats[2] := '# ##0.000000';        // H_Cotation
    vGrille.ColFormats[3] := '# ##0.000000000';     // H_TAUXREEL
    vGrille.ColFormats[4] := '# ##0.000000000';     // H_TAUXCLOTURE

    // Longueur de saisie des champs
    vGrille.ColLengths[1] := 10;
    vGrille.ColLengths[4] := 20;
    vGrille.ColLengths[5] := 17;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPCHANCELL.GenererConditionPlus : String;
begin
    // recuperation de la condition de date
  Result := '';
//  if (OkDate.checked) then
    Result := '(H_DATECOURS>="' + USDate(DateD) +'" AND H_DATECOURS<="' + USDate(DateF) + '")';
end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... :   /  /
Modifié le ... : 26/03/2003
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCHANCELL.OnCreateTOB(vTOB: TOB);
var
  dt : TDateTime;
begin
    vTOB.PutValue('H_SOCIETE', V_PGI.CodeSociete); {JP 05/05/04 : !!!!!}
    vTOB.PutValue('H_DEVISE',Combo.value);
    {JP 20/08/03 : Initialisation avec la date de saisie s'il n'y a pas de ligne, sinon avec la date de la ligne précédante}
    if Fliste.RowCount < 3 then
    begin
        vTOB.PutValue('H_DATECOURS',DateDeb)
    end
    else
    begin
        dt := StrToDate(Fliste.CellValues[1, Fliste.RowCount - 2]);
        dt :=  dt + 1;
        if (DateDeb < dt) then vTOB.PutValue('H_DATECOURS',dt)
        else vTOB.PutValue('H_DATECOURS',DateDeb);
    end;
end;

//==================================================
// Evenements par default de la TOF
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCHANCELL.OnNew;
begin
    inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCHANCELL.OnDelete;
begin
    inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCHANCELL.OnUpdate;
begin
  {JP 18/05/06 : FQ 17546 : Mise à jour des quotités / Taux}
  MajQuotiteTaux(Fliste.Col, Fliste.Row);
  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCHANCELL.OnLoad;
begin
    // init devise
    if devise <> ''
      then SetControlText('H_DEVISE', devise)
      else Combo.ItemIndex := 0;

    {JP 20/08/03 : Pour remettre à jour l'écran si l'on change de devise}
    Combo.OnChange := OnDeviseChange;
    inherited;
    DateD.OnChange := OnDateChange;
    DateF.OnChange := OnDateChange;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCHANCELL.OnArgument(S : String);
var
    s1,s2 : string;
begin
    Action := taConsult;
    SurEuro := false;

    DateDeb := iDate1900;
    DateFin := iDate2099;

    // recup des parametre
    s1 := uppercase(S);
    // Mode
    s2 := ReadTokenSt(s1);
    if (s2 = 'ACTION=CREATION') then Action := taCreat
    else if (s2 = 'ACTION=MODIFICATION') then Action := taModif;
    // quelle devise
    s2 := ReadTokenSt(s1); if (s2 <> '') then devise := s2;
    // sur euro ??
    s2 := ReadTokenSt(s1); if (s2 = 'SUREURO') then SurEuro := true;
    // date de debut de selection
    s2 := ReadTokenSt(s1); if (s2 <> '') then DateDeb := StrToDate(s2);
    // choix des devise
    s2 := ReadTokenSt(s1); SetControlEnabled('H_DEVISE',(not (s2 = 'OUI')));
    // Sur deux date ??
    s2 := ReadTokenSt(s1); if (s2 <> '') then FSur2Dates := (S2 = 'TRUE') else FSur2Dates := false;
    // date de fin de selection
    s2 := ReadTokenSt(s1); if (s2 <> '') then DateFin := StrToDate(s2);

    // recup des control utils
    TQuotite := THLabel(GetControl('TQuotite',true)); if (not assigned(TQuotite)) then exit;
    TValeurQ := THLabel(GetControl('TValeurQ',true)); if (not assigned(TValeurQ)) then exit;
    TValable := THLabel(GetControl('TValable',true)); if (not assigned(TValable)) then exit;
    RGSens := THRadioGroup(GetControl('RGSens',true)); if (not assigned(RGSens)) then exit;
    BPurge := TSpeedButton(GetControl('BPurge',true)); if (not assigned(BPurge)) then exit;
    DateD := THEdit(GetControl('DATEDU',true)); if (not assigned(DateD)) then exit;
    DateF := THEdit(GetControl('DATEDU_',true)); if (not assigned(DateF)) then exit;
    OkDate := TCheckBox(GetControl('OKDATE',true)); if (not assigned(OkDate)) then exit;
    Combo := THValComboBox(GetControl('H_DEVISE',true)); if (not assigned(Combo)) then exit;

    // set des fonction
    RGSens.OnClick := OnClickRGSens;
    BPurge.OnClick := OnClickBPurge;
    OkDate.OnClick := OnChangeOkDate;

    // init des dates
    {JP 20/08/03 : Il me semble mieux de filtrer les taux sur une période restreinte et proche du jour}
    if (DateDeb = idate1900) then DateDeb := DebutDeMois(Date-5);
    if (DateFin = idate2099) then DateFin := FinDeMois(Date+5);
    
{$IFNDEF TRESO}
    // selon euro
    if not SurEuro then
    begin
        if (DateFin > v_pgi.datedebuteuro) then DateFin := v_pgi.datedebuteuro-1;
        if (DateDeb > DateFin) then DateDeb := idate1900;
    end
    else 
    begin
        if (DateDeb < v_pgi.datedebuteuro) then DateDeb := v_pgi.datedebuteuro;
        if (DateDeb > DateFin) then DateFin := idate2099;
    end;
{$ENDIF}

    DateD.Text := DateToStr(DateDeb);
    DateF.Text := DateToStr(DateFin);

    inherited;

    {JP 20/08/03 : Pour gérer la conversion Cotation/Taux}
    FListe.OnCellExit := FListeCellExit;
    DateF.OnExit := OnDateExit;
    DateD.OnExit := OnDateExit;

    // 14371
    if SurEuro then Ecran.HelpContext := 1150200
               else Ecran.HelpContext := 1150100;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCHANCELL.OnDisplay;
begin
    inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCHANCELL.OnClose;
begin
    inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCHANCELL.OnCancel;
begin
    inherited;
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCHANCELL.OnClickRGSens(Sender: TObject);
begin
    // pb de focus et de contenu !
    FListe.col := 1;

    // affichage de certaine collone selon le sens
    if (RGSens.ItemIndex = 1) then
    begin
        if (not (FListe.ColWidths[2] = -1)) then
        begin
            FListe.ColLengths[3]  := 20;
            FListe.ColWidths[3]   := FListe.ColWidths[2];
        end ;
        FListe.ColLengths[2] := -1 ;
        FListe.ColWidths[2]   := -1 ;
    end
    else
    begin
        if (not (FListe.ColWidths[3] = -1)) then
        begin
            FListe.ColLengths[2]  := 20;
            FListe.ColWidths[2]   := FListe.ColWidths[3];
        end ;
        FListe.ColLengths[3]  := -1 ;
        FListe.ColWidths[3]   := -1 ;
    end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCHANCELL.OnClickBPurge(Sender: TObject);
{$IFDEF EAGLCLIENT}
{$ELSE}
var
    LesDatesASup : TDateTime;
    Depuis, svg : String;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
{$ELSE}

    svg := GetField('H_DATECOURS');
    //22/06/06    YMO FQ18297 Plantage lors de la comparaison directe du variant retourné par Getfield
    if (svg <> '') then
    begin
        Depuis := GetField('H_DATECOURS');
        if (PurgeOui(Depuis,combo.Value,LesDatesASup)) then ExecuteSQl('DELETE FROM CHANCELL WHERE H_DEVISE="'+combo.Value+'" '+'AND H_DATECOURS<="'+USDatetime(LesDatesASup)+'" ');
    end;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/03/2003
Modifié le ... : 18/08/2004
Description .. : JP : FQ 13909 -13910 : on autorise la gestion des filtres
Suite ........ : dans tous les cas mais on fait un contrôle sur les date en
Suite ........ : fonction de l'euro
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCHANCELL.OnChangeOkDate(Sender: TObject);
begin
  {On commence par désactiver les évènements OnChange pour ne pas lancer x fois RemplitGrille}
  DateD.OnChange := nil;
  DateF.OnChange := nil;
  try
    {Les contrôles ne sont actifs que si on demande le filtre sur les dates}
    SetControlEnabled('TDATEDU', OkDate.Checked);
    SetControlEnabled('TDATEAU', OkDate.Checked);
    DateD.Enabled := OkDate.Checked;
    DateF.Enabled := OkDate.Checked;
    {Vérifie les dates et lance RemplitGrille}
    VerifCoherenceDate;
  finally
    DateD.OnChange := OnDateChange;
    DateF.OnChange := OnDateChange;
  end;
end;

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPCHANCELL.QuotiteChange(devise : string) : Integer;
var
    Q : TQuery;
begin
    // selection de la quotité
    Result := 1;
    Q := OpenSQL('SELECT D_QUOTITE FROM DEVISE WHERE D_DEVISE="'+devise+'"', true);
    if (not Q.Eof) then Result := Q.Fields[0].AsInteger;
    Ferme(Q);
end;

{JP 20/08/03 : Calcule (1/montant saisi dans la cellule) et le formate en fonction du paramétrage
               de la devise. Le montant calculé sera affecté au taux si l'on saisit une cotation
               et inversement si l'on saisit un taux
{---------------------------------------------------------------------------------------}
procedure TOF_CPCHANCELL.FListeCellExit(Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);
{---------------------------------------------------------------------------------------}
begin
  MajQuotiteTaux(ACol, ARow);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCHANCELL.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if (Key = VK_DOWN) and (Shift = []) and (Fliste.Row = Fliste.RowCount - 1) then begin
    Key := 0 ;
    BInsertClick(Sender);
  end ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCHANCELL.OnDeviseChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
    RemplitGrille;
end;

{JP 18/08/04 : FQ 13909/10 : gestion  de la cohérence des dates en fonction des options autour de l'euro
{---------------------------------------------------------------------------------------}
procedure TOF_CPCHANCELL.VerifCoherenceDate;
{---------------------------------------------------------------------------------------}
begin
  {S'il n'y a pas de filtre sur les dates}
  if not OkDate.Checked then begin
    {et qu'il s'agit d'une devise non européennes sur l'€uro}
    if SurEuro then begin
      {On positionne la date de début sur la date de passage à l'euro ...}
      DateD.Text := DateToStr(V_PGI.DateDebutEuro);
      {... et celle de fin sur la date maximum}
      DateF.Text := StDate2099;
    end
    {Il s'agit d'une devise européenne ou de chacellerie sur Franc}
    else if not SurEuro then begin
      {On positionne la date de fin à la veille du passage à l'euro ...}
      DateF.Text := DateToStr(V_PGI.DateDebutEuro - 1);
      {... et celle de début sur la date minimum}
      DateD.Text := StDate1900;
    end;
  end

  {On filtre les dates, On s'assure que les dates sont cohérentes avec l'€uro}
  else begin
    {et qu'il s'agit d'une devise non européennes sur l'€uro}
    if SurEuro then begin
      {On positionne la date de début sur la date de passage à l'euro si besoin}
      if (StrToDateTime(DateD.Text) < V_PGI.DateDebutEuro) then
        DateD.Text := DateToStr(V_PGI.DateDebutEuro);
    end
    {Il s'agit d'une devise européenne ou de chacellerie sur Franc}
    else if not SurEuro then begin
      {On positionne la date de fin à la veille du passage à l'euro si besoin}
      if (StrToDateTime(DateF.Text) >= V_PGI.DateDebutEuro) then
        DateF.Text := DateToStr(V_PGI.DateDebutEuro - 1);
    end;
  end;
  {Lancement de la recherche}
  RemplitGrille;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCHANCELL.OnDateChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if not CIsValideDate(THEdit(Sender).Text) then Exit;
  {On commence par désactiver les évènements OnChange pour ne pas lancer x fois RemplitGrille}
  DateD.OnChange := nil;
  DateF.OnChange := nil;
  try
   {Vérifie les dates et lance RemplitGrille}
    VerifCoherenceDate;
  finally
    DateD.OnChange := OnDateChange;
    DateF.OnChange := OnDateChange;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCHANCELL.OnDateExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if (StrToDate(DateF.Text) < StrToDate(DateD.Text)) {
     JP 12/01/06 : FQ 17089 : J'avais bu, je suppose, lorsque j'ai écrit cet évènement
      or (StrToDate(DateF.Text) > StrToDate(DateD.Text))} then begin
    HShowMessage('0;' + Ecran.Caption + ';Les dates ne sont pas cohérentes.;I;O;O;O;', '', '');
  end;
end;

{18/08/04 : On s'assure que les critères de recherche sont valides : par défaut renvoie True}
{---------------------------------------------------------------------------------------}
function TOF_CPCHANCELL.IsCritereValide : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited IsCritereValide;
  Result := Result and CIsValideDate(DateF.Text) and CIsValideDate(DateD.Text);
end;

{JP 18/05/06 : FQ 17546 : Je sors le traitement de FListeCellExit afin de  pouvoir l'appeler sur d'autres évènements
{---------------------------------------------------------------------------------------}
procedure TOF_CPCHANCELL.MajQuotiteTaux(Col, Row : Integer);
{---------------------------------------------------------------------------------------}
var
  Mnt,Inv,svg : Double;
  Tax,Rst : string;
  lInDecTaux : integer ;
  lInDecInv  : integer ;
begin
  if not(Col in [2, 3]) then Exit;

  // FQ17720 : tauxreel base sur 9 décimales, cotation sur 6
  if RGSens.ItemIndex = 0 then
    begin
    lInDecTaux := 6 ;
    lInDecInv  := 9 ;
    end
  else
    begin
    lInDecTaux := 9 ;
    lInDecInv  := 6 ;
    end ;

  // formatage du taux
  Mnt := Arrondi(Valeur(FListe.Cells[Col, Row]),lInDecTaux);
  Tax := StrfMontant(Mnt,8,lInDecTaux, '', true);

  // recup de la donnée
  if RGSens.ItemIndex = 0
    then svg := GetField('H_COTATION', Row-1)
    else svg := GetField('H_TAUXREEL', Row-1);
  // si pas de changement on touche pas
  if (Mnt = svg) then exit;

  // calcul du taux inverse
  if (Mnt = 0)
    then Inv := 0
    else Inv := Arrondi(1/Mnt,lInDecInv);
  Rst := StrfMontant(Inv,8+lInDecInv,lInDecInv,'',true);

  // set ...
  if RGSens.ItemIndex = 0 then
  begin
    FListe.Cells[2, Row] := Tax; SetField('H_COTATION', Mnt, Row - 1);
    FListe.Cells[3, Row] := Rst; SetField('H_TAUXREEL', Inv, Row - 1);
  end
  else
  begin
    FListe.Cells[2, Row] := Rst; SetField('H_COTATION', Inv, Row - 1);
    FListe.Cells[3, Row] := Tax; SetField('H_TAUXREEL', Mnt, Row - 1);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCHANCELL.FListeOnExit(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {JP 18/05/06 : FQ 17546 : Mise à jour des quotités / Taux}
  MajQuotiteTaux(Fliste.Col, Fliste.Row);
end;


//================================================================================
// Initialization
//================================================================================

Initialization
    registerclasses([TOF_CPCHANCELL]);
end.
