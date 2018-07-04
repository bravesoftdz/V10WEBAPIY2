unit ASSISTCODEAPPEL;

interface

uses
  Windows, Messages, SysUtils, StdCtrls, Hctrls, Mask, Spin,
  ExtCtrls, ComCtrls, HSysMenu, hmsgbox, HTB97, Controls, Classes,  Graphics,  Forms, Dialogs
  ,Dicobtp,assist, Paramsoc, Buttons,HEnt1,UIUtil,Ent1,EntGC,LicUtil  ,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Utob
{$IFDEF EAGLCLIENT}
  ,eTablette
{$ELSE}
   ,Tablette,HDB, HPanel
{$ENDIF}
   ;
Procedure LanceAssistCodeAppel ;

type
  TFAssistCodeAppel = class(TFAssist)
    TabSheet1: TTabSheet;
    LIB1: THLabel;
    HLabel2: THLabel;
    libgene1: THLabel;
    BevelDecoupe: TBevel;
    LIBNBPART: THLabel;
    LibGene2: THLabel;
    LibGene3: THLabel;
    LibGene4: THLabel;
    PARTIE1: TTabSheet;
    TSO_APPCODENBPARTIE: THLabel;
    TSO_APPCO1LIB: THLabel;
    TSO_APPCO1TYPE: THLabel;
    TSO_APPCO1LNG: THLabel;
    Entetepartie1: THLabel;
    BevelPartie1: TBevel;
    SO_APPCO1LIB: THCritMaskEdit;
    SO_APPCO1TYPE: THValComboBox;
    SO_APPCODENBPARTIE: TSpinEdit;
    SO_APPCO1LNG: TSpinEdit;
    TSO_APPCO1VALEUR: THLabel;
    SO_APPCO1VALEUR: THCritMaskEdit;
    PARTIE2: TTabSheet;
    PARTIE3: TTabSheet;
    TSO_APPCO2LIB: THLabel;
    TSO_APPCO2TYPE: THLabel;
    TSO_APPCO2LNG: THLabel;
    LIBPARTIE2: THLabel;
    BevelPartie2: TBevel;
    SO_APPCO2LIB: THCritMaskEdit;
    SO_APPCO2TYPE: THValComboBox;
    SO_APPCO2LNG: TSpinEdit;
    TSO_APPCO2VALEUR: THLabel;
    SO_APPCO2VALEUR: THCritMaskEdit;
    TSO_APPCO3LIB: THLabel;
    TSO_APPCO3TYPE: THLabel;
    TSO_APPCO3LNG: THLabel;
    LibPartie3: THLabel;
    BevelPartie3: TBevel;
    SO_APPCO3LIB: THCritMaskEdit;
    SO_APPCO3TYPE: THValComboBox;
    SO_APPCO3LNG: TSpinEdit;
    TSO_APPCO3VALEUR: THLabel;
    SO_APPCO3VALEUR: THCritMaskEdit;
    SO_APPCO2VISIBLE: TCheckBox;
    SO_APPCO3VISIBLE: TCheckBox;
    PARAMAFFAIREPART1: TBitBtn;
    PARAMAFFAIREPART2: TBitBtn;
    PARAMAFFAIREPART3: TBitBtn;
    LIBLNG: THLabel;
    LNGDISPO: THLabel;
    LIBLNG2: THLabel;
    TSO_APPCO1VALEURPRO: THLabel;
    SO_APPCO1VALEURPRO: THCritMaskEdit;
    TSO_APPCO2VALEURPRO: THLabel;
    SO_APPCO2VALEURPRO: THCritMaskEdit;
    TSO_APPCO3VALEURPRO: THLabel;
    SO_APPCO3VALEURPRO: THCritMaskEdit;
    TSO_BTNATAFFAIRE: THLabel;
    TSO_BTNATPROPOSITION: THLabel;
    SO_BTNATAFFAIRE: THValComboBox;
    SO_BTNATPROPOSITION: THValComboBox;
    TSCompl: TTabSheet;
    HLabel3: THLabel;
    Bevel1: TBevel;
    LibGene5: THLabel;
    LibGene6: THLabel;
    HLabel1: THLabel;
    SO_APPPRODIFFERENT: TCheckBox;
    HLabel4: THLabel;
    HLabel5: THLabel;
    HLabel6: THLabel;
    SO_APPRECODIFICATION: TCheckBox;
    SO_AAMM3: TRadioButton;
    SO_AAAAMM3: TRadioButton;
    SO_AAAAMM1: TRadioButton;
    SO_AAMM1: TRadioButton;
    SO_AAAAMM2: TRadioButton;
    SO_AAMM2: TRadioButton;

    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure SO_APPCO1TYPEClick(Sender: TObject);
    procedure PARAMAFFAIREPART1Click(Sender: TObject);
    procedure SO_APPCODENBPARTIEChange(Sender: TObject);
    procedure SO_APPCOVALEURExit(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure SO_APPCO1LNGChange(Sender: TObject);
    procedure SO_APPPRODIFFERENTClick(Sender: TObject);
    procedure bAideClick(Sender: TObject);
    procedure SO_AAMM1Click(Sender: TObject);
    procedure SO_AAAAMM1Click(Sender: TObject);
  private
    { Déclarations privées }
    bCreation,Erreur : Boolean;
    ExisteCompteur : Boolean;
    LngMax , iLngDispo : integer;

    Procedure CalculLngDispo;
    procedure ChargementPartieCodeAppel;
    Procedure PartieEnter (NumPart : integer);
    Procedure FormatValeurPartie (NumPart,Lng : integer;Appel : Boolean);
    procedure FormateZoneCreation;
    procedure GestionPartie;
    Procedure RazCompteur (CC: THEDIT; Cache : Boolean);

    function  TestCptAff (Num : string; TypeaFF:string;Ii:integer): boolean;


  public
    { Déclarations publiques }
    function  PageCount : Integer ; override ;
  end;



implementation
{$R *.DFM}
Procedure LanceAssistCodeAppel ;
Var FAssistCodeAppel  : TFAssistCodeAppel ;
begin
if Blocage (['AffToutSeul'],True,'AffToutSeul') then  exit; // quelqu'un d'autre travaille sur la base

FAssistCodeAppel:=TFAssistCodeAppel.Create(Application) ;
//FAssistCodeAppel.SO_APPCODENBPARTIE.Value := VH_GC.CleAffaire.NbPartie;
FAssistCodeAppel.ShowModal ;

FAssistCodeAppel.Free;

Bloqueur ('AffToutSeul',False);

end;


procedure TFAssistCodeAppel.FormCreate(Sender: TObject);
begin
inherited;

if ExisteSQL ('SELECT AFF_AFFAIRE From AFFAIRE WHERE AFF_AFFAIRE0="W"') then
	 bCreation := False
else
	 bCreation := True;

//POUR TEST
//bCreation := true;
ExisteCompteur := false;
LngMax := 14;

P.Pages[1].TabVisible := False;
P.Pages[2].PageIndex := 1;
P.Pages[3].PageIndex := 2;
P.Pages[4].PageIndex := 3;

end;

procedure TFAssistCodeAppel.FormShow(Sender: TObject);
begin
  inherited;
bFin.Enabled := False;
Erreur:=false;

// Champs non modifiables si pas en création
SO_APPCODENBPARTIE.Enabled := False;

SO_APPCO1TYPE.Enabled := bCreation;
SO_APPCO3TYPE.Enabled := bCreation;
SO_APPCO2TYPE.Enabled := bCreation;

SO_AAMM1.Enabled := bCreation;
SO_AAMM2.Enabled := bCreation;
SO_AAMM3.Enabled := bCreation;

SO_AAAAMM1.Enabled := bCreation;
SO_AAAAMM2.Enabled := bCreation;
SO_AAAAMM3.Enabled := bCreation;

SO_APPCO1VALEUR.Enabled := bCreation;
SO_APPCO2VALEUR.Enabled := bCreation;
SO_APPCO3VALEUR.Enabled := bCreation;

// maj Champs général
SO_APPRECODIFICATION.Checked := GetParamSoc ('SO_APPRECODIFICATION');

//Chargement des zones en fonction du code appel
if bcreation then
   Begin
   FormateZoneCreation;
   ChargementPartieCodeAppel;
   exit;
   End
else
	 Begin
	 SO_APPCODENBPARTIE.Value := GetParamSoc('SO_APPCODENBPARTIE');
   SO_APPCO1LNG.Value := GetParamSoc('SO_APPCO1LNG');
   SO_APPCO2LNG.Value := GetParamSoc('SO_APPCO2LNG');
   SO_APPCO3LNG.Value := GetParamSoc('SO_APPCO3LNG');
   end;

SO_APPCO1LNG.Enabled := False;
SO_APPCO2LNG.Enabled := False;
SO_APPCO3LNG.Enabled := False;

SO_APPCO1TYPE.Enabled := False;
SO_APPCO2TYPE.Enabled := False;
SO_APPCO3TYPE.Enabled := False;

   ChargementPartieCodeAppel;

   CalculLngDispo;

// afficher les champs nature de pièce / mot de passe du jour
if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
   begin
   TSO_BTNATAFFAIRE.Visible := True;
   SO_BTNATAFFAIRE.Visible := True;
   TSO_BTNATPROPOSITION.Visible := True;
   SO_BTNATPROPOSITION.Visible := True;
   end;

end;

Procedure TFAssistCodeAppel.ChargementPartieCodeAppel;
begin

// Partie 1 du code affaire
SO_APPCO1LIB.Text := GetParamSoc ('SO_APPCO1LIB');
SO_APPCO1TYPE.Value := GetParamSoc ('SO_APPCO1TYPE');
SO_APPCO1VALEUR.Text := GetParamSoc ('SO_APPCO1VALEUR');
SO_APPCO1VALEURPRO.Text := GetParamSoc ('SO_APPCO1VALEURPRO');

// Partie 2 du code affaire
SO_APPCO2LIB.Text := GetParamSoc('SO_APPCO2LIB');
SO_APPCO2TYPE.Value := GetParamSoc('SO_APPCO2TYPE');
SO_APPCO2VALEUR.Text := GetParamSoc('SO_APPCO2VALEUR');
SO_APPCO2VISIBLE.Checked := GetParamSoc('SO_APPCO2VISIBLE');
SO_APPCO2VALEURPRO.Text := GetParamSoc('SO_APPCO2VALEURPRO');

// Partie 3 du code affaire
SO_APPCO3LIB.Text := GetParamSoc('SO_APPCO3LIB');
SO_APPCO3TYPE.Value := GetParamSoc('SO_APPCO3TYPE');
SO_APPCO3VALEUR.Text := GetParamSoc('SO_APPCO3VALEUR');
SO_APPCO3VALEURPRO.Text := GetParamSoc('SO_APPCO3VALEURPRO');
SO_APPCO3VISIBLE.Checked := GetParamSoc('SO_APPCO3VISIBLE');

// modif brl 160707
//if SO_APPCO1TYPE.Value = 'FIX' then SO_APPCO1VALEUR.enabled := False;
//if SO_APPCO2TYPE.Value = 'FIX' then SO_APPCO2VALEUR.enabled := False;
//if SO_APPCO3TYPE.Value = 'FIX' then SO_APPCO3VALEUR.enabled := False;

end;

Procedure TFAssistCodeAppel.FormateZoneCreation;
Begin

   // test pour ne pas recharger les valeurs à partir de celles du paramétrage affaire
   if (GetParamSoc('SO_APPCODENBPARTIE')= GetParamSoc('SO_AFFCODENBPARTIE')) AND
      (GetParamSoc('SO_APPCO1LNG') = GetParamSoc('SO_AFFCO1LNG')) AND
      (GetParamSoc('SO_APPCO1LNG') = GetParamSoc('SO_AFFCO1LNG')) AND
      (GetParamSoc('SO_APPCO1LNG') = GetParamSoc('SO_AFFCO1LNG')) Then
       Begin
    	 SO_APPCODENBPARTIE.Value := GetParamSoc('SO_APPCODENBPARTIE');
       SO_APPCO1LNG.Value := GetParamSoc('SO_APPCO1LNG');
       SO_APPCO2LNG.Value := GetParamSoc('SO_APPCO2LNG');
       SO_APPCO3LNG.Value := GetParamSoc('SO_APPCO3LNG');
       exit;
       end;

	 SO_APPCODENBPARTIE.Value := GetParamSoc('SO_AFFCODENBPARTIE');

   SO_APPCO1LNG.Value := GetParamSoc('SO_AFFCO1LNG');
   SO_APPCO2LNG.Value := GetParamSoc('SO_AFFCO2LNG');
   SO_APPCO3LNG.Value := GetParamSoc('SO_AFFCO3LNG');

   SO_APPCO1VALEUR.enabled := True;
   SO_APPCO2VALEUR.enabled := True;
   SO_APPCO3VALEUR.enabled := True;

   if (GetParamSoc('SO_AFFCO1TYPE') = 'LIS') OR (GetParamSoc('SO_AFFCO1TYPE') = 'SAI') then
      Begin
      SO_APPCO1TYPE.Value := 'FIX';
			SO_APPCO1LIB.Text := 'Valeur Fixe';
			SO_APPCO1VALEUR.MaxLength :=  SO_APPCO1LNG.Value;
			SO_APPCO1VALEUR.Text := '';
      end
   else if GetParamSoc('SO_AFFCO1TYPE') = 'CPT' then
      Begin
      SO_APPCO1TYPE.Value := 'CPT';
  		SO_APPCO1LIB.Text := 'Compteur';
     	SO_APPCO1VALEUR.Text := '';
      end
   else if GetParamSoc('SO_AFFCO1TYPE') = 'CPM' then
      Begin
      SO_APPCO1TYPE.Value := 'CPT';
  		SO_APPCO1LIB.Text := 'Compteur';
     	SO_APPCO1VALEUR.Text := '';
      end;

   if (GetParamSoc('SO_AFFCO2TYPE') = 'LIS') OR (GetParamSoc('SO_AFFCO2TYPE') = 'SAI') then
			Begin
      SO_APPCO2TYPE.Value := 'FIX';
			SO_APPCO2LIB.Text := 'Valeur Fixe';
			SO_APPCO2VALEUR.Text := '';
      end
   else if GetParamSoc('SO_AFFCO2TYPE') = 'CPT' then
      Begin
      SO_APPCO2TYPE.Value := 'CPT';
  		SO_APPCO2LIB.Text := 'Compteur';
     	SO_APPCO2VALEUR.Text := '';
      end
   else if GetParamSoc('SO_AFFCO2TYPE') = 'CPM' then
      Begin
      SO_APPCO2TYPE.Value := 'CPT';
  		SO_APPCO2LIB.Text := 'Compteur';
     	SO_APPCO2VALEUR.Text := '';
      end;

   if (GetParamSoc('SO_AFFCO3TYPE') = 'LIS') OR (GetParamSoc('SO_AFFCO3TYPE') = 'SAI') then
			Begin
      SO_APPCO3TYPE.Value := 'FIX';
			SO_APPCO3LIB.Text := 'Valeur Fixe';
			SO_APPCO3VALEUR.Text := '';
      end
   else if GetParamSoc('SO_AFFCO3TYPE') = 'CPT' then
      Begin
      SO_APPCO3TYPE.Value := 'CPT';
  		SO_APPCO3LIB.Text := 'Compteur';
     	SO_APPCO3VALEUR.Text := '';
      end
   else if GetParamSoc('SO_AFFCO3TYPE') = 'CPM' then
      Begin
      SO_APPCO3TYPE.Value := 'CPT';
  		SO_APPCO3LIB.Text := 'Compteur';
     	SO_APPCO3VALEUR.Text := '';
      end;


	 if (GetParamSoc('SO_AFFCO1TYPE') <> 'CPT') AND (GetParamSoc('SO_AFFCO2TYPE') <> 'CPT') AND (GetParamSoc('SO_AFFCO3TYPE') <> 'CPT') Then
      if GetParamSoc('SO_AFFCODENBPARTIE') = 1 then
         Begin
	       SO_APPCO1TYPE.Value := 'CPT';
    		 SO_APPCO1LIB.Text := 'Compteur';
     		 SO_APPCO1VALEUR.Text := '';
         end
      Else if GetParamSoc('SO_AFFCODENBPARTIE') = 2 then
         Begin
	       SO_APPCO2TYPE.Value := 'CPT';
    		 SO_APPCO2LIB.Text := 'Compteur';
     		 SO_APPCO2VALEUR.Text := '';
         end
      Else if GetParamSoc('SO_AFFCODENBPARTIE') = 3 then
         Begin
	       SO_APPCO3TYPE.Value := 'CPT';
    		 SO_APPCO3LIB.Text := 'Compteur';
     		 SO_APPCO3VALEUR.Text := '';
         end;

  SO_APPCO1LNG.Enabled := False;
  SO_APPCO2LNG.Enabled := False;
  SO_APPCO3LNG.Enabled := False;

  SO_APPCO1TYPE.Enabled := False;
  SO_APPCO2TYPE.Enabled := False;
  SO_APPCO3TYPE.Enabled := False;

// Mise à jour immédiate dans les paransoc
SetParamSoc('SO_APPCODENBPARTIE',SO_APPCODENBPARTIE.Value);
SetParamSoc ('SO_APPCO1LNG',SO_APPCO1LNG.Value);
SetParamSoc ('SO_APPCO1TYPE',SO_APPCO1TYPE.Value);
SetParamSoc ('SO_APPCO2TYPE',SO_APPCO2TYPE.Value);
SetParamSoc ('SO_APPCO2LNG',SO_APPCO2LNG.Value);
SetParamSoc ('SO_APPCO3TYPE',SO_APPCO3TYPE.Value);
SetParamSoc ('SO_APPCO3LNG',SO_APPCO3LNG.Value);

SetParamSoc ('SO_APPCO1VALEUR',SO_APPCO1VALEUR.Text);
SetParamSoc ('SO_APPCO1VALEURPRO',SO_APPCO1VALEURPRO.Text);
SetParamSoc ('SO_APPCO2VALEUR',SO_APPCO2VALEUR.Text);
SetParamSoc ('SO_APPCO2VALEURPRO',SO_APPCO2VALEURPRO.Text);
SetParamSoc ('SO_APPCO3VALEUR',SO_APPCO3VALEUR.Text);
SetParamSoc ('SO_APPCO3VALEURPRO',SO_APPCO3VALEURPRO.Text);

SetParamSoc ('SO_APPCO1LIB',SO_APPCO1LIB.Text);
SetParamSoc ('SO_APPCO2LIB',SO_APPCO2LIB.Text);
SetParamSoc ('SO_APPCO3LIB',SO_APPCO3LIB.Text);

SetParamSoc ('SO_APPCO2VISIBLE',SO_APPCO2VISIBLE.Checked);
SetParamSoc ('SO_APPCO3VISIBLE',SO_APPCO3VISIBLE.Checked);

end;

Procedure TFAssistCodeAppel.GestionPartie;
Begin

if SO_APPCODENBPARTIE.Value = 1 then
	 Begin
   SO_APPCO1LIB.Text := 'Compteur sur Appel';
	 SO_APPCO1TYPE.Enabled := False;
   SO_APPCO1TYPE.Text := '';
   SO_APPCO1TYPE.Value := 'CPT';
   SO_APPCO1VALEUR.Text :='';
   SO_APPCO1TYPE.Enabled := false;
   End;

if (SO_APPCODENBPARTIE.Value = 2) then
   Begin
   SO_APPCO2LIB.Text := 'Compteur sur Appel';
   SO_APPCO2TYPE.Enabled := false;
   SO_APPCO2TYPE.Text := '';
   SO_APPCO2TYPE.Value := 'CPT';
   SO_APPCO2VALEUR.Text :='';
   end;

if (SO_APPCODENBPARTIE.Value = 3) then
   Begin
   SO_APPCO3LIB.Text := 'Compteur sur Appel';
	 SO_APPCO3TYPE.Enabled := False;
   SO_APPCO3TYPE.Text := '';
   SO_APPCO3TYPE.Value := 'CPT';
   end;

if SO_APPCO1TYPE.Value = 'FIX' then SO_APPCO1VALEUR.enabled := bcreation;
if SO_APPCO2TYPE.Value = 'FIX' then SO_APPCO2VALEUR.enabled := bcreation;
if SO_APPCO3TYPE.Value = 'FIX' then SO_APPCO3VALEUR.enabled := bcreation;

end;

procedure TFAssistCodeAppel.bFinClick(Sender: TObject);
begin
  inherited;

// Valeurs par défaut
SetParamSoc ('SO_APPORMATEXER','AUC');

FormatValeurPartie (1,SO_APPCO1LNG.Value,False);
FormatValeurPartie (1,SO_APPCO1LNG.Value,True);
FormatValeurPartie (2,SO_APPCO2LNG.Value,False);
FormatValeurPartie (2,SO_APPCO2LNG.Value,True);
FormatValeurPartie (3,SO_APPCO3LNG.Value,False);
FormatValeurPartie (3,SO_APPCO3LNG.Value,True);

if bCreation then
   BEGIN
   // Certains éléments peuvent être maj en création uniquement.
   SetParamSoc('SO_APPCODENBPARTIE',SO_APPCODENBPARTIE.Value);
   SetParamSoc ('SO_APPCO1LNG',SO_APPCO1LNG.Value);
   SetParamSoc ('SO_APPCO1TYPE',SO_APPCO1TYPE.Value);
   SetParamSoc ('SO_APPCO2TYPE',SO_APPCO2TYPE.Value);
   SetParamSoc ('SO_APPCO2LNG',SO_APPCO2LNG.Value);
   SetParamSoc ('SO_APPCO3TYPE',SO_APPCO3TYPE.Value);
   SetParamSoc ('SO_APPCO3LNG',SO_APPCO3LNG.Value);
   END;

SetParamSoc ('SO_APPCO1VALEUR',SO_APPCO1VALEUR.Text);
SetParamSoc ('SO_APPCO1VALEURPRO',SO_APPCO1VALEURPRO.Text);
SetParamSoc ('SO_APPCO2VALEUR',SO_APPCO2VALEUR.Text);
SetParamSoc ('SO_APPCO2VALEURPRO',SO_APPCO2VALEURPRO.Text);
SetParamSoc ('SO_APPCO3VALEUR',SO_APPCO3VALEUR.Text);
SetParamSoc ('SO_APPCO3VALEURPRO',SO_APPCO3VALEURPRO.Text);

SetParamSoc ('SO_APPCO1LIB',SO_APPCO1LIB.Text);
SetParamSoc ('SO_APPCO2LIB',SO_APPCO2LIB.Text);
SetParamSoc ('SO_APPCO3LIB',SO_APPCO3LIB.Text);

SetParamSoc ('SO_APPCO2VISIBLE',SO_APPCO2VISIBLE.Checked);
SetParamSoc ('SO_APPCO3VISIBLE',SO_APPCO3VISIBLE.Checked);

// mcd 18/09/01 mais partie visible en nb parti gérées
if (GetParamSoc('SO_APPCODENBPARTIE') >=2) then
   SetParamSoc ('SO_APPCO2VISIBLE','X')
else
	 SetParamSoc ('SO_APPCO2VISIBLE','-');

if (GetParamSoc('SO_APPCODENBPARTIE') >=3) then
	 SetParamSoc ('SO_APPCO3VISIBLE','X')
else
	 SetParamSoc ('SO_APPCO3VISIBLE','-');

SetParamSoc ('SO_APPPRODIFFERENT',SO_APPPRODIFFERENT.Checked);
SetParamSoc ('SO_APPRECODIFICATION',SO_APPRECODIFICATION.Checked);

// Natures de pièces
SetParamSoc ('SO_BTNATAFFAIRE',SO_BTNATAFFAIRE.Value);
SetParamSoc ('SO_BTNATPROPOSITION',SO_BTNATPROPOSITION.Value);

//ChargeSocieteHalley; AB 14/01/2003
ChargeParamsGC ;
self.close ;
end;

procedure TFAssistCodeAppel.SO_APPCO1TYPEClick(Sender: TObject);
Var NumPart : integer;
    Sp : TSpinEdit;
    LibValeur : THLabel;
    Valeur,ValeurPro : THEdit;
    Combo : THValComboBox;
    Bt : TBitBtn;
    LibLongCode : THLabel;
    LongCode : TSpinEdit;
    AAMM : TRadioButton;
    AAAAMM : TRadioButton;
begin
  inherited;

  Combo := THValComboBox(Sender);

if Combo.name = 'SO_APPCO1TYPE' then NumPart := 1 else
if Combo.name = 'SO_APPCO2TYPE' then NumPart := 2 else
if Combo.name = 'SO_APPCO3TYPE' then NumPart := 3 else exit;

Sp := TSpinEdit(FindComponent ('SO_APPCO'+ IntToStr(NumPart) + 'LNG'));

LibValeur := THLabel(FindComponent( 'TSO_APPCO' + IntToStr(NumPart) + 'VALEUR'));
LongCode := TSpinEdit(FindComponent( 'SO_APPCO' + IntToStr(NumPart) + 'LNG'));
LibLongCode := THLabel(FindComponent( 'TSO_APPCO' + IntToStr(NumPart) + 'LNG'));
Valeur := THEdit(FindComponent( 'SO_APPCO' + IntToStr(NumPart) + 'VALEUR'));
ValeurPro := THEdit(FindComponent( 'SO_APPCO' + IntToStr(NumPart) + 'VALEURPRO'));
Bt := TBitBtn(FindComponent( 'PARAMAFFAIREPART'+IntToStr(NumPart)));
AAMM := TRadioButton(FindComponent( 'SO_AAMM' + IntToStr(NumPart)));
AAAAMM := TRadioButton(FindComponent( 'SO_AAAAMM' + IntToStr(NumPart)));

//initialisation à l'affichage des zones écran si Année-Mois
//LongCode.Enabled := True;
LongCode.left := 147;

LibValeur.Visible := true;
LibValeur.Enabled := True;
Valeur.Visible := true;
LibLongCode.Visible := True;
AAMM.checked := false;
AAAAMM.checked := false;
AAMM.Visible := False;
AAAAMM.Visible := False;

//if bCreation then
//   BEGIN
//	 LongCode.Value := 1;
//	 Valeur.Maxlength := 14;
//   END;

if Bt <> Nil then  Bt.Visible := false;

if (combo.value = 'LIS') then
   BEGIN
   if Sp <> Nil then
   	  BEGIN
      Sp.Enabled := False;
      Sp.Value := 3;
      END;
   if Valeur <> Nil then razCompteur(Valeur,True);
   if (ValeurPro <> Nil) And (SO_APPPRODIFFERENT.Checked) then razCompteur(ValeurPro,True);
   if Bt <> Nil then  Bt.Visible := true;
   END
else if (combo.value = 'FIX') then
   Begin
   if not bcreation then sp.enabled := false;
   end
Else if (combo.value = 'DAT') then
	 BEGIN
   LongCode.enabled := false;
   LongCode.left := 302;
   LibLongCode.Visible := false;
   LibValeur.Enabled := False;
   LibValeur.caption := 'Exemple de formatage';
   AAMM.Visible := True;
   AAAAMM.Visible := True;
   if bCreation then
	   BEGIN
  	 AAMM.checked :=  True;
	   AAMM.OnClick (AAMM);
  	 END
   END
Else if (combo.value = 'CPT') then
   Begin
   LibValeur.Caption := 'Compteur sur Appel';
   Valeur.Enabled := True;
   end
else   // Fixe + Cpt
   BEGIN
   if (Sp <> Nil) And (bCreation) then Sp.Enabled := True;
   if Valeur <> Nil then razCompteur(Valeur,True);
   if (ValeurPro <> Nil) And (SO_APPPRODIFFERENT.Checked) then razCompteur(ValeurPro,True);
   END;
end;

procedure TFAssistCodeAppel.PARAMAFFAIREPART1Click(Sender: TObject);
Var NumPart : integer;
begin
  inherited;
NumPart := TBitBtn(Sender).Tag;
ParamTable('AFFAIREPART'+IntToStr(NumPart),taCreat,0,Nil) ;
end;

procedure TFAssistCodeAppel.CalculLngDispo;
BEGIN
iLngDispo := LngMax - SO_APPCO1LNG.Value - SO_APPCO2LNG.Value - SO_APPCO3LNG.Value;

if iLngDispo <0 then iLNGDispo := 0;

LNGDispo.caption := IntToStr(iLngDispo);

END;

procedure TFAssistCodeAppel.PartieEnter(NumPart : integer);
Var  sp 			: TSpinEdit;
     LngAutre	: Integer;
     LngDispo : Integer;
begin
  inherited;

  CalculLngDispo;
  sp := TSpinEdit(FindComponent('SO_APPCO'+ IntToStr(NumPart) + 'LNG'));

if sp <> Nil then
   BEGIN
   if NumPart = 1 then
      LngAutre := SO_APPCO2LNG.Value + SO_APPCO3LNG.Value
   else if NumPart = 2 then
      LngAutre := SO_APPCO1LNG.Value + SO_APPCO3LNG.Value
   else if NumPart = 3 then
   		LngAutre := SO_APPCO1LNG.Value + SO_APPCO2LNG.Value
   else
   		LngAutre := 0;
   LngDispo := LngMax - LngAutre;
   sp.maxValue := LngDispo;
   if sp.Value > LngDispo then sp.Value := LngDispo;
   END;

end;

procedure TFAssistCodeAppel.SO_APPCODENBPARTIEChange(Sender: TObject);
begin
  inherited;

if SO_APPCODENBPARTIE.Value = 0 then GestionPartie;

if SO_APPCODENBPARTIE.Value < 2 then
    BEGIN
    SO_APPCO2LIB.Text := 'non utilisé';
    SO_APPCO2TYPE.Text := '';
    SO_APPCO2LNG.MinValue := 0;
    SO_APPCO2LNG.Value := 0;
    SO_APPCO2VALEUR.Text := '';
    SO_APPCO2VALEURPRO.Text := '';
    SO_APPCO2VISIBLE.Checked := false;
    END;

if SO_APPCODENBPARTIE.Value < 3 then
    BEGIN
    SO_APPCO3LIB.Text := 'non utilisé';
    SO_APPCO3TYPE.Text := '';
    SO_APPCO3LNG.MinValue := 0;
    SO_APPCO3LNG.Value := 0;
    SO_APPCO3VALEUR.Text := '';
    SO_APPCO3VALEURPRO.Text := '';
    SO_APPCO3VISIBLE.Checked := false;
    END;

CalculLngDispo;

end;

procedure TFAssistCodeAppel.bSuivantClick(Sender: TObject);
Var Onglet : TTabSheet;
    St_NomOnglet : String;
begin

if Erreur=true then exit;

inherited;
Onglet := P.ActivePage;  // Activation de la page suivante
st_NomOnglet := Onglet.Name;

if st_NomOnglet = 'PARTIE1' then
    BEGIN
    SO_APPCODENBPARTIEChange(Nil); // forcer l'init des zones non gérées PA le 13/12/2000
    SO_APPPRODIFFERENTClick(Nil);
    SO_APPCO1TYPEClick(SO_APPCO1TYPE);
    //PartieEnter(1);
    if SO_APPCODENBPARTIE.Value = 1 then BSuivant.Enabled := False;
    END
else if st_NomOnglet = 'PARTIE2' then
    BEGIN
    SO_APPCO1TYPEClick(SO_APPCO2TYPE);
    //PartieEnter(2);
    if SO_APPCODENBPARTIE.Value = 2 then BSuivant.Enabled := False;
    END
else if st_NomOnglet = 'PARTIE3' then
    BEGIN
    SO_APPCO1TYPEClick(SO_APPCO3TYPE);
    //PartieEnter(3);
    if SO_APPCODENBPARTIE.Value = 3 then BSuivant.Enabled := False;
    END;

if (bSuivant.Enabled) then
    bFin.Enabled := False
else
		bFin.Enabled := True;

end;

procedure TFAssistCodeAppel.bPrecedentClick(Sender: TObject);
Var Onglet : TTabSheet;
    St_NomOnglet : String;
begin
// cas ou erreur sur compteur, il ne faut pas passer sur page prec
If Erreur=true then exit;

  inherited;
Onglet := P.ActivePage;
st_NomOnglet := Onglet.Name;

if st_NomOnglet = 'PARTIE1' then PartieEnter(1) else
if st_NomOnglet = 'PARTIE2' then PartieEnter(2) else
if st_NomOnglet = 'PARTIE3' then PartieEnter(3) ;

bSuivant.Enabled := true;
bFin.Enabled := False;

end;

procedure TFAssistCodeAppel.SO_APPCO1LNGChange(Sender: TObject);
begin
  inherited;
CalculLngDispo;
FormatValeurPartie (TSpinEdit(Sender).Tag,TSpinEdit(Sender).Value,False);
FormatValeurPartie (TSpinEdit(Sender).Tag,TSpinEdit(Sender).Value,True);
end;

procedure TFAssistCodeAppel.SO_APPCOVALEURExit(Sender: TObject);
var
  ii,cpt:integer;
  typeAff:string;

begin
  // mcd 29/04/2002 pour ctrl modif code dans les valeurs permises
  inherited;

  Erreur:=False;

  if (THCritMaskEdit(Sender).name='SO_APPCO3VALEUR') then
     begin
     ii:=3;
     if GetParamSoc('SO_APPCO3TYPE') <> 'CPT' then Exit;
     cpt :=GetParamSoc('SO_APPCO3VALEUR');
     TypeAff:='AFF';
     end
  else if (THCritMaskEdit(Sender).name='SO_APPCO3VALEURPRO') then
     begin
     ii:=3;
     if GetParamSoc('SO_APPCO3TYPEPRO') <> 'CPT' then Exit;
     cpt :=GetParamSoc('SO_APPCO3VALEURPRO');
     TypeAff:='PRO';
     end
  else if (THCritMaskEdit(Sender).name='SO_APPCO2VALEURPRO') then
     begin
     ii:=2;
     if GetParamSoc('SO_APPCO2TYPE') <> 'CPT' then Exit;
     cpt :=GetParamSoc('SO_APPCO2VALEURPRO');
     TypeAff:='PRO';
     end
  else if (THCritMaskEdit(Sender).name='SO_APPCO2VALEUR') then
     begin
     ii:=2;
     if GetParamSoc('SO_APPCO2TYPE') <> 'CPT' then Exit;
     cpt :=GetParamSoc('SO_APPCO2VALEUR');
     TypeAff:='AFF';
     end
  else if (THCritMaskEdit(Sender).name='SO_APPCO1VALEURPRO') then
     begin
     ii:=1;
     if GetParamSoc('SO_APPCO1TYPE') <> 'CPT' then Exit;
     cpt :=GetParamSoc('SO_APPCO1VALEURPRO');
     TypeAff:='PRO';
     end
  else if (THCritMaskEdit(Sender).name='SO_APPCO1VALEUR') then
     begin
     ii:=1;
     if GetParamSoc('SO_APPCO1TYPE') <> 'CPT' then Exit;
     cpt :=GetParamSoc('SO_APPCO1VALEUR');
     TypeAff:='AFF';
     end
  else
     begin
     ii := 0;
     cpt := 0;
     end;

  if (TestCptAff (THCritMaskEdit(Sender).Text,TypeAff,ii))=false then
  begin
    if (V_PGI.PassWord <> CryptageSt(DayPass(Date))) then
    begin
      PGIBoxAf('Compteur inférieur à ce qui existe en fichier. Valeur interdite','');
      THCritMaskEdit(Sender).text:=InttoStr(cpt);
      THCritMaskEdit(Sender).SetFocus;
      Erreur:=True;
    end
    else
    begin
      // on peut forcer si on est en mot de passe du jour
      // a faire en connaissance de cause, pour réutiliser pe un code non utilisé
      if PGIAsk('Compteur inférieur à ce qui existe en fichier. Confirmez-vous le changement','')<> mrYes then
      begin
        THCritMaskEdit(Sender).Text:=IntToStr(cpt);
        THCritMaskEdit(Sender).SetFocus;
        Erreur:=true;
      end;
    end;
  end;
end;

function TFAssistCodeAppel.TestCptAff (Num : String; TypeAff:string;ii:integer): boolean;
Var QQ : TQuery ;
sql:string;
Begin
result:=true;
sql:=  'SELECT aff_Affaire FROM AFFAIRE';
//if (SO_APPPRODIFFERENT.Checked) then sql:=sql+ ' WHERE AFF_STATUTAFFAIRE="'+TypeAff+'"';
if (SO_APPPRODIFFERENT.Checked) then sql:=sql+ ' WHERE AFF_AFFAIRE0="W"';
if not(ExisteSQL(sql)) then exit;
sql:=  'SELECT MAX(AFF_AFFAIRE'+ InttoStr(ii)+') FROM AFFAIRE ';
//if (SO_APPPRODIFFERENT.Checked) then sql:=sql+' WHERE AFF_STATUTAFFAIRE="'+TypeAff+'"';
if (SO_APPPRODIFFERENT.Checked) then sql:=sql+' WHERE AFF_AFFAIRE0="W"';
QQ:=OpenSQL(sql,TRUE,-1,'',true);
if Not QQ.EOF then
   if (StrtoInt(Num))> QQ.Fields[0].AsInteger
      then result:=true
      else result:=false;
Ferme(QQ);
end;


procedure TFAssistCodeAppel.FormatValeurPartie (NumPart,Lng : integer; Appel : Boolean);
Var CC : THEdit;
    st,suff : String;
    Combo : THValComboBox;
    Cpt : integer;
BEGIN

if Appel then
	 suff := 'APP'
else
	 suff := '';

if Appel And Not(SO_APPPRODIFFERENT.Checked) then Exit;

Combo := THValComboBox(findComponent('SO_APPCO'+ IntToStr(NumPart)+'TYPE'));
CC := THEdit(findComponent('SO_APPCO'+ IntToStr(NumPart)+'VALEUR'+Suff));

if CC <> Nil then
	 CC.MaxLength := Lng
else
	 exit;

if Combo = Nil then Exit;

if CC.Text = '' then CC.Text := '0';

if Combo.Value = 'CPT' then // Formattage avec des 0
    BEGIN
    st := Copy (CC.Text,1,Lng);
    if length(st) < lng then
        BEGIN
        Cpt := StrToInt(st);
        st := Format('%*.*d',[lng,lng,Cpt]);
        CC.Text := st;
        END;
    END;
END;

procedure TFAssistCodeAppel.SO_APPPRODIFFERENTClick(Sender: TObject);
Var ProVisible : Boolean;
begin
  inherited;
ProVisible := SO_APPPRODIFFERENT.Checked;
SO_APPCO1VALEURPRO.Visible := ProVisible;
TSO_APPCO1VALEURPRO.Visible := ProVisible;
SO_APPCO2VALEURPRO.Visible := ProVisible;
TSO_APPCO2VALEURPRO.Visible := ProVisible;
SO_APPCO3VALEURPRO.Visible := ProVisible;
TSO_APPCO3VALEURPRO.Visible := ProVisible;
end;

function TFAssistCodeAppel.PageCount : Integer ;
BEGIN
result := SO_APPCODENBPARTIE.Value + 2;  //gm 13/01/03
END ;

Procedure TFAssistCodeAppel.razCompteur (CC: THEDIT;Cache : Boolean);
BEGIN
if CC = nil then Exit;
if Cache then
   BEGIN
   CC.Enabled := False;
   CC.Color := clBtnFace;
   CC.Text := '';
   END
else
   BEGIN
   CC.Enabled := True;
   CC.Color := clwindow;
   END;
END;

procedure TFAssistCodeAppel.bAideClick(Sender: TObject);
begin
  inherited;
CallHelpTopic(Self) ;
end;

procedure TFAssistCodeAppel.SO_AAMM1Click(Sender: TObject);
Var NumPart    : integer;
    Valeur     : THEdit;
    LongCode   : TSpinEdit;
    FormatDate : TRadioButton;
begin
  inherited;
//Recherche du N° d'écran et donc du nom de la zone
FormatDate := TRadioButton(Sender);
if FormatDate.name = 'SO_AAMM1' then NumPart := 1 else
if FormatDate.name = 'SO_AAMM2' then NumPart := 2 else
if FormatDate.name = 'SO_AAMM3' then NumPart := 3 else exit;

LongCode := TSpinEdit(FindComponent( 'SO_APPCO' + IntToStr(NumPart)+ 'LNG'));
Valeur := THEdit(FindComponent( 'SO_APPCO' + IntToStr(NumPart) + 'VALEUR'));

LongCode.VALUE := 4;

Valeur.Maxlength := LongCode.value;
Valeur.Text := FormatDateTime('yymm', now);

CalculLngDispo;

end;

procedure TFAssistCodeAppel.SO_AAAAMM1Click(Sender: TObject);
Var NumPart    : integer;
    Valeur     : THEdit;
    LongCode   : TSpinEdit;
    FormatDate : TRadioButton;
begin
  inherited;

//Recherche du N° d'écran et donc du nom de la zone
FormatDate := TRadioButton(Sender);
if FormatDate.name = 'SO_AAAAMM1' then NumPart := 1 else
if FormatDate.name = 'SO_AAAAMM2' then NumPart := 2 else
if FormatDate.name = 'SO_AAAAMM3' then NumPart := 3 else exit;

LongCode := TSpinEdit(FindComponent( 'SO_APPCO' + IntToStr(NumPart)+ 'LNG'));
Valeur := THEdit(FindComponent( 'SO_APPCO' + IntToStr(NumPart) + 'VALEUR'));

LongCode.VALUE := 6;
Valeur.Maxlength := LongCode.value;
Valeur.Text := FormatDateTime('yyyymm', now);

CalculLngDispo;

end;

end.
