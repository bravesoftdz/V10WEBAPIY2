unit AssistCodeAffaire;

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
Procedure LanceAssistCodeAffaire ;

type
  TFAssistCodeaffaire = class(TFAssist)
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
    TSO_AFFCODENBPARTIE: THLabel;
    TSO_AFFCO1LIB: THLabel;
    TSO_AFFCO1TYPE: THLabel;
    TSO_AFFCO1LNG: THLabel;
    Entetepartie1: THLabel;
    BevelPartie1: TBevel;
    SO_AFFCO1LIB: THCritMaskEdit;
    SO_AFFCO1TYPE: THValComboBox;
    SO_AFFCODENBPARTIE: TSpinEdit;
    SO_AFFCO1LNG: TSpinEdit;
    TSO_AFFCO1VALEUR: THLabel;
    SO_AFFCO1VALEUR: THCritMaskEdit;
    PARTIE2: TTabSheet;
    PARTIE3: TTabSheet;
    TSO_AFFCO2LIB: THLabel;
    TSO_AFFCO2TYPE: THLabel;
    TSO_AFFCO2LNG: THLabel;
    LIBPARTIE2: THLabel;
    BevelPartie2: TBevel;
    SO_AFFCO2LIB: THCritMaskEdit;
    SO_AFFCO2TYPE: THValComboBox;
    SO_AFFCO2LNG: TSpinEdit;
    TSO_AFFCO2VALEUR: THLabel;
    SO_AFFCO2VALEUR: THCritMaskEdit;
    TSO_AFFCO3LIB: THLabel;
    TSO_AFFCO3TYPE: THLabel;
    TSO_AFFCO3LNG: THLabel;
    LibPartie3: THLabel;
    BevelPartie3: TBevel;
    SO_AFFCO3LIB: THCritMaskEdit;
    SO_AFFCO3TYPE: THValComboBox;
    SO_AFFCO3LNG: TSpinEdit;
    TSO_AFFCO3VALEUR: THLabel;
    SO_AFFCO3VALEUR: THCritMaskEdit;
    SO_AFFCO2VISIBLE: TCheckBox;
    SO_AFFCO3VISIBLE: TCheckBox;
    PARAMAFFAIREPART1: TBitBtn;
    PARAMAFFAIREPART2: TBitBtn;
    PARAMAFFAIREPART3: TBitBtn;
    LIBLNG: THLabel;
    LNGDISPO: THLabel;
    LIBLNG2: THLabel;
    TSO_AFFCO1VALEURPRO: THLabel;
    SO_AFFCO1VALEURPRO: THCritMaskEdit;
    TSO_AFFCO2VALEURPRO: THLabel;
    SO_AFFCO2VALEURPRO: THCritMaskEdit;
    TSO_AFFCO3VALEURPRO: THLabel;
    SO_AFFCO3VALEURPRO: THCritMaskEdit;
    TSO_AFNATAFFAIRE: THLabel;
    TSO_AFNATPROPOSITION: THLabel;
    SO_AFNATAFFAIRE: THValComboBox;
    SO_AFNATPROPOSITION: THValComboBox;
    TSCompl: TTabSheet;
    HLabel3: THLabel;
    Bevel1: TBevel;
    LibGene5: THLabel;
    LibGene6: THLabel;
    HLabel1: THLabel;
    SO_AFFPRODIFFERENT: TCheckBox;
    HLabel4: THLabel;
    HLabel5: THLabel;
    HLabel6: THLabel;
    SO_AFFRECODIFICATION: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure SO_AFFCO1TYPEClick(Sender: TObject);
    procedure PARAMAFFAIREPART1Click(Sender: TObject);
    procedure SO_AFFCODENBPARTIEChange(Sender: TObject);
    procedure SO_AFFCOVALEURExit(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure SO_AFFCO1LNGChange(Sender: TObject);
    procedure SO_AFFPRODIFFERENTClick(Sender: TObject);
    procedure bAideClick(Sender: TObject);
  private
    { Déclarations privées }
    bCreation,Erreur : Boolean;
    LngMax , iLngDispo : integer;
    Procedure CalculLngDispo;
    Procedure PartieEnter (NumPart : integer);
    Procedure FormatValeurPartie (NumPart,Lng : integer;Propos : Boolean);
    Procedure razCompteur (CC: THEDIT; Cache : Boolean);
    function TestCptAff (Num : string; TypeaFF:string;Ii:integer): boolean;
  public
    { Déclarations publiques }
    function  PageCount : Integer ; override ;
  end;



implementation
{$R *.DFM}
Procedure LanceAssistCodeAffaire ;
Var FAssistCodeaffaire  : TFAssistCodeaffaire ;
begin
if Blocage (['AffToutSeul'],True,'AffToutSeul') then  exit; // quelqu'un d'autre travaille sur la base

FAssistCodeaffaire:=TFAssistCodeaffaire.Create(Application) ;
FAssistCodeaffaire.SO_AFFCODENBPARTIE.Value := VH_GC.CleAffaire.NbPartie;
FAssistCodeaffaire.ShowModal ;

FAssistCodeaffaire.Free;
Bloqueur ('AffToutSeul',False);
end;


procedure TFAssistCodeaffaire.FormCreate(Sender: TObject);
begin
inherited;

{$IFDEF BTP}
if ExisteSQL ('SELECT AFF_AFFAIRE From AFFAIRE WHERE AFF_AFFAIRE0="A"') then bCreation := False
                                                                        else bCreation := True;
{$ELSE}
if ExisteSQL ('SELECT AFF_AFFAIRE From AFFAIRE') then bCreation := False
                                                 else bCreation := True;
{$ENDIF}
////// bCreation := true; // POUR TEST
LngMax := 14;
{$IFDEF BTP}
P.Pages[1].TabVisible := False;
P.Pages[2].PageIndex := 1;
P.Pages[3].PageIndex := 2;
P.Pages[4].PageIndex := 3;
{$ENDIF}
end;

procedure TFAssistCodeaffaire.FormShow(Sender: TObject);
Var Combo : THValComboBox;
begin
  inherited;
bFin.Enabled := False;
Erreur:=false;
// Champs non modifiables si pas en création
SO_AFFCODENBPARTIE.Enabled := bCreation;
SO_AFFCO1TYPE.Enabled := bCreation; SO_AFFCO1LNG.Enabled := bCreation;
SO_AFFCO2TYPE.Enabled := bCreation; SO_AFFCO2LNG.Enabled := bCreation;
SO_AFFCO3TYPE.Enabled := bCreation; SO_AFFCO3LNG.Enabled := bCreation;

// maj Champs général
SO_AFFCODENBPARTIE.Value   := VH_GC.CleAffaire.NbPartie;
SO_AFFPRODIFFERENT.Checked := VH_GC.CleAffaire.ProDifferent;
SO_AFFRECODIFICATION.Checked := GetParamSoc ('SO_AFFRECODIFICATION');
// Partie 1 du code affaire
SO_AFFCO1LIB.Text := GetParamSoc ('SO_AFFCO1LIB');
SO_AFFCO1TYPE.Value := GetParamSoc ('SO_AFFCO1TYPE');
SO_AFFCO1LNG.Value := GetParamSoc ('SO_AFFCO1LNG');
SO_AFFCO1VALEUR.Text := GetParamSoc ('SO_AFFCO1VALEUR');
SO_AFFCO1VALEURPRO.Text := GetParamSoc ('SO_AFFCO1VALEURPRO');
// Partie 2 du code affaire
SO_AFFCO2LIB.Text := GetParamSoc ('SO_AFFCO2LIB');
SO_AFFCO2TYPE.Value := GetParamSoc ('SO_AFFCO2TYPE');
SO_AFFCO2LNG.Value := GetParamSoc ('SO_AFFCO2LNG');
SO_AFFCO2VALEUR.Text := GetParamSoc ('SO_AFFCO2VALEUR');
SO_AFFCO2VISIBLE.Checked := GetParamSoc ('SO_AFFCO2VISIBLE');
SO_AFFCO2VALEURPRO.Text := GetParamSoc ('SO_AFFCO2VALEURPRO');
// Partie 3 du code affaire
SO_AFFCO3LIB.Text := GetParamSoc ('SO_AFFCO3LIB');
SO_AFFCO3TYPE.Value := GetParamSoc ('SO_AFFCO3TYPE');
SO_AFFCO3LNG.Value := GetParamSoc ('SO_AFFCO3LNG');
SO_AFFCO3VALEUR.Text := GetParamSoc ('SO_AFFCO3VALEUR');
SO_AFFCO3VALEURPRO.Text := GetParamSoc ('SO_AFFCO3VALEURPRO');
SO_AFFCO3VISIBLE.Checked := GetParamSoc ('SO_AFFCO3VISIBLE');
// Natures de pièces
SO_AFNATAFFAIRE.Value := GetParamSoc ('SO_AFNATAFFAIRE');
SO_AFNATPROPOSITION.Value := GetParamSoc ('SO_AFNATPROPOSITION');

Combo := THValComboBox(FindComponent('SO_AFFCO1TYPE'));
SO_AFFCO1TYPEClick(Combo);
Combo := THValComboBox(FindComponent('SO_AFFCO2TYPE'));
SO_AFFCO1TYPEClick(Combo);
Combo := THValComboBox(FindComponent('SO_AFFCO3TYPE'));
SO_AFFCO1TYPEClick(Combo);
CalculLngDispo;
// afficher les champs nature de pièce / mot de passe du jour
if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
   begin
   TSO_AFNATAFFAIRE.Visible := True; SO_AFNATAFFAIRE.Visible := True;
   TSO_AFNATPROPOSITION.Visible := True; SO_AFNATPROPOSITION.Visible := True;
   end;
end;

procedure TFAssistCodeaffaire.bFinClick(Sender: TObject);
begin
  inherited;
//NextControl(self); AB 14/01/2003
// Valeurs par défaut
SetParamSoc ('SO_AFFORMATEXER','AUC');
// SetParamsoc ('SO_AFFGESTIONAVENANT',False);

FormatValeurPartie (1,SO_AFFCO1LNG.Value,False); FormatValeurPartie (1,SO_AFFCO1LNG.Value,True);
FormatValeurPartie (2,SO_AFFCO2LNG.Value,False); FormatValeurPartie (2,SO_AFFCO2LNG.Value,True);
FormatValeurPartie (3,SO_AFFCO3LNG.Value,False); FormatValeurPartie (3,SO_AFFCO3LNG.Value,True);
if bCreation then
   BEGIN
   // Certains éléments peuvent être maj en création uniquement.
   SetParamSoc('SO_AFFCODENBPARTIE',SO_AFFCODENBPARTIE.Value);
   SetParamSoc ('SO_AFFCO1TYPE',SO_AFFCO1TYPE.Value);
   SetParamSoc ('SO_AFFCO1LNG',SO_AFFCO1LNG.Value);
   SetParamSoc ('SO_AFFCO2TYPE',SO_AFFCO2TYPE.Value);
   SetParamSoc ('SO_AFFCO2LNG',SO_AFFCO2LNG.Value);
   SetParamSoc ('SO_AFFCO3TYPE',SO_AFFCO3TYPE.Value);
   SetParamSoc ('SO_AFFCO3LNG',SO_AFFCO3LNG.Value);
   END;

SetParamSoc ('SO_AFFCO1VALEUR',SO_AFFCO1VALEUR.Text);
SetParamSoc ('SO_AFFCO1VALEURPRO',SO_AFFCO1VALEURPRO.Text);
SetParamSoc ('SO_AFFCO2VALEUR',SO_AFFCO2VALEUR.Text);
SetParamSoc ('SO_AFFCO2VALEURPRO',SO_AFFCO2VALEURPRO.Text);
SetParamSoc ('SO_AFFCO3VALEUR',SO_AFFCO3VALEUR.Text);
SetParamSoc ('SO_AFFCO3VALEURPRO',SO_AFFCO3VALEURPRO.Text);

SetParamSoc ('SO_AFFCO1LIB',SO_AFFCO1LIB.Text);
SetParamSoc ('SO_AFFCO2LIB',SO_AFFCO2LIB.Text);
SetParamSoc ('SO_AFFCO3LIB',SO_AFFCO3LIB.Text);

SetParamSoc ('SO_AFFCO2VISIBLE',SO_AFFCO2VISIBLE.Checked);
SetParamSoc ('SO_AFFCO3VISIBLE',SO_AFFCO3VISIBLE.Checked);
            // mcd 18/09/01 mais partie visible en nb parti gérées
if (GetParamSoc('SO_AFFCODENBPARTIE') >=2) then SetParamSoc ('SO_AFFCO2VISIBLE','X')
   else SetParamSoc ('SO_AFFCO2VISIBLE','-');
if (GetParamSoc('SO_AFFCODENBPARTIE') >=3) then SetParamSoc ('SO_AFFCO3VISIBLE','X')
   else SetParamSoc ('SO_AFFCO3VISIBLE','-');
SetParamSoc ('SO_AFFPRODIFFERENT',SO_AFFPRODIFFERENT.Checked);
SetParamSoc ('SO_AFFRECODIFICATION',SO_AFFRECODIFICATION.Checked);
// Natures de pièces
SetParamSoc ('SO_AFNATAFFAIRE',SO_AFNATAFFAIRE.Value);
SetParamSoc ('SO_AFNATPROPOSITION',SO_AFNATPROPOSITION.Value);

//ChargeSocieteHalley; AB 14/01/2003
ChargeParamsGC ;
self.close ;
end;

procedure TFAssistCodeaffaire.SO_AFFCO1TYPEClick(Sender: TObject);
Var NumPart : integer;
    Sp : TSpinEdit;
    LibValeur : THLabel;
    Valeur,ValeurPro : THEdit;
    Combo : THValComboBox;
    Bt : TBitBtn;
begin
  inherited;
Combo := THValComboBox(Sender);
if Combo.name = 'SO_AFFCO1TYPE' then NumPart := 1 else
if Combo.name = 'SO_AFFCO2TYPE' then NumPart := 2 else
if Combo.name = 'SO_AFFCO3TYPE' then NumPart := 3 else exit;
Sp := TSpinEdit(FindComponent ('SO_AFFCO'+ IntToStr(NumPart) + 'LNG'));
LibValeur := THLabel(FindComponent( 'TSO_AFFCO' + IntToStr(NumPart) + 'VALEUR'));
Valeur := THEdit(FindComponent( 'SO_AFFCO' + IntToStr(NumPart) + 'VALEUR'));
ValeurPro := THEdit(FindComponent( 'SO_AFFCO' + IntToStr(NumPart) + 'VALEURPRO'));
Bt := TBitBtn(FindComponent( 'PARAMAFFAIREPART'+IntToStr(NumPart)));
if Bt <> Nil then  Bt.Visible := false;
if (combo.value = 'LIS') then
   BEGIN
   if Sp <> Nil then BEGIN Sp.Enabled := False;  Sp.Value := 3; END;
   if Valeur <> Nil then razCompteur(Valeur,True);
   if (ValeurPro <> Nil) And (SO_AFFPRODIFFERENT.Checked) then razCompteur(ValeurPro,True);
   if Bt <> Nil then  Bt.Visible := true;
   END else
if (combo.value = 'SAI') then
   BEGIN
   if (Sp <> Nil) And (bCreation) then Sp.Enabled := True;
   if Valeur <> Nil then razCompteur(Valeur,True);
   if (ValeurPro <> Nil) And (SO_AFFPRODIFFERENT.Checked) then razCompteur(ValeurPro,True);
   END
else   // Fixe + Cpt
   BEGIN
   if (Sp <> Nil) And (bCreation) then Sp.Enabled := True;
   if Valeur <> Nil then razCompteur(Valeur,False);
   if (ValeurPro <> Nil) And (SO_AFFPRODIFFERENT.Checked) then razCompteur(ValeurPro,False);
   if (combo.value = 'FIX') then LibValeur.Caption := 'Valeur fixe' else
   if (combo.value = 'CPT') then LibValeur.Caption := 'Compteur sur affaires' else
   if (combo.value = 'CPM') then LibValeur.Caption := 'Compteur sur affaires modifiable';
   END;
end;

procedure TFAssistCodeaffaire.PARAMAFFAIREPART1Click(Sender: TObject);
Var NumPart : integer;
begin
  inherited;
NumPart := TBitBtn(Sender).Tag;
ParamTable('AFFAIREPART'+IntToStr(NumPart),taCreat,0,Nil) ;
end;

procedure TFAssistCodeaffaire.CalculLngDispo;
BEGIN
iLngDispo := LngMax - SO_AFFCO1LNG.Value - SO_AFFCO2LNG.Value - SO_AFFCO3LNG.Value;
if iLngDispo <0 then iLNGDispo := 0;
LNGDispo.caption := IntToStr(iLngDispo);
END;


procedure TFAssistCodeaffaire.PartieEnter(NumPart : integer);
Var  sp : TSpinEdit;
     LngAutre,LngDispo : Integer;

begin
  inherited;
CalculLngDispo;
sp := TSpinEdit(FindComponent('SO_AFFCO'+ IntToStr(NumPart) + 'LNG'));
if sp <> Nil then
   BEGIN
   if NumPart = 1 then BEGIN LngAutre := SO_AFFCO2LNG.Value + SO_AFFCO3LNG.Value; END
   else if NumPart = 2 then BEGIN LngAutre := SO_AFFCO1LNG.Value + SO_AFFCO3LNG.Value; END
   else if NumPart = 3 then BEGIN LngAutre := SO_AFFCO1LNG.Value + SO_AFFCO2LNG.Value; END
   else LngAutre := 0;
   LngDispo := LngMax - LngAutre;
   sp.maxValue := LngDispo; if sp.Value > LngDispo then sp.Value := LngDispo;
   END;
end;

procedure TFAssistCodeaffaire.SO_AFFCODENBPARTIEChange(Sender: TObject);
Var ModifLng : Boolean;
begin
  inherited;
ModifLng := False;
if SO_AFFCODENBPARTIE.Value < 2 then
    BEGIN
    SO_AFFCO2LIB.Text := 'non utilisé'; SO_AFFCO2TYPE.Text := '';
    SO_AFFCO2LNG.MinValue := 0; SO_AFFCO2LNG.Value := 0;
    SO_AFFCO2VALEUR.Text := ''; SO_AFFCO2VALEURPRO.Text := '';
    SO_AFFCO2VISIBLE.Checked := false;
    ModifLng := True;
    END;
if SO_AFFCODENBPARTIE.Value < 3 then
    BEGIN
    SO_AFFCO3LIB.Text := 'non utilisé'; SO_AFFCO3TYPE.Text := '';
    SO_AFFCO3LNG.MinValue := 0; SO_AFFCO3LNG.Value := 0;
    SO_AFFCO3VALEUR.Text := ''; SO_AFFCO3VALEURPRO.Text := '';
    SO_AFFCO3VISIBLE.Checked := false;
    ModifLng := True;
    END;
if ModifLng then CalculLngDispo;
end;

procedure TFAssistCodeaffaire.bSuivantClick(Sender: TObject);
Var Onglet : TTabSheet;
    St_NomOnglet : String;
begin
//nextControl(self);  AB 14/01/2003
if Erreur=true then exit;
  inherited;
Onglet := P.ActivePage;  // Activation de la page suivante
st_NomOnglet := Onglet.Name;

if st_NomOnglet = 'PARTIE1' then
    BEGIN
    SO_AFFCODENBPARTIEChange(Nil); // forcer l'init des zones non gérées PA le 13/12/2000
    SO_AFFPRODIFFERENTClick(Nil);
    SO_AFFCO1TYPEClick(SO_AFFCO1TYPE);
    PartieEnter(1);
    if SO_AFFCODENBPARTIE.Value = 1 then BSuivant.Enabled := False;
    END else
if st_NomOnglet = 'PARTIE2' then
    BEGIN
    SO_AFFCO1TYPEClick(SO_AFFCO2TYPE);
    PartieEnter(2);
    if SO_AFFCODENBPARTIE.Value = 2 then BSuivant.Enabled := False;
    END else
if st_NomOnglet = 'PARTIE3' then
    BEGIN
    SO_AFFCO1TYPEClick(SO_AFFCO3TYPE);
    PartieEnter(3);
    if SO_AFFCODENBPARTIE.Value = 3 then BSuivant.Enabled := False;
    END;
if (bSuivant.Enabled) then   bFin.Enabled := False else  bFin.Enabled := True;
end;



procedure TFAssistCodeaffaire.bPrecedentClick(Sender: TObject);
Var Onglet : TTabSheet;
    St_NomOnglet : String;
begin
//NextControl(self);AB 14/01/2003
If Erreur=true then exit; // cas ou erreur sur compteur, il ne faut pas passer sur page prec
  inherited;
Onglet := P.ActivePage;
st_NomOnglet := Onglet.Name;
if st_NomOnglet = 'PARTIE1' then PartieEnter(1) else
if st_NomOnglet = 'PARTIE2' then PartieEnter(2) else
if st_NomOnglet = 'PARTIE3' then PartieEnter(3) ;

bSuivant.Enabled := true; bFin.Enabled := False;
end;

procedure TFAssistCodeaffaire.SO_AFFCO1LNGChange(Sender: TObject);
begin
  inherited;
CalculLngDispo;
FormatValeurPartie (TSpinEdit(Sender).Tag,TSpinEdit(Sender).Value,False);
FormatValeurPartie (TSpinEdit(Sender).Tag,TSpinEdit(Sender).Value,True);
end;

procedure TFAssistCodeaffaire.SO_AFFCOVALEURExit(Sender: TObject);
var
  ii,cpt:integer;
  typeAff:string;

begin
  // mcd 29/04/2002 pour ctrl modif code dans les valeurs permises
  inherited;
  Erreur:=False;
  if (THCritMaskEdit(Sender).name='SO_AFFCO3VALEUR') then begin ii:=3; cpt :=GetParamSoc('SO_AFFCO3VALEUR');TypeAff:='AFF'; end
  else  if (THCritMaskEdit(Sender).name='SO_AFFCO3VALEURPRO') then begin ii:=3; cpt :=GetParamSoc('SO_AFFCO3VALEURPRO');TypeAff:='PRO'; end
  else  if (THCritMaskEdit(Sender).name='SO_AFFCO2VALEURPRO') then begin ii:=2; cpt :=GetParamSoc('SO_AFFCO2VALEURPRO'); TypeAff:='PRO';end
  else  if (THCritMaskEdit(Sender).name='SO_AFFCO2VALEUR') then begin ii:=2; cpt :=GetParamSoc('SO_AFFCO2VALEUR');TypeAff:='AFF'; end
  else  if (THCritMaskEdit(Sender).name='SO_AFFCO1VALEURPRO') then begin ii:=1; cpt :=GetParamSoc('SO_AFFCO1VALEURPRO'); TypeAff:='PRO';end
  else  if (THCritMaskEdit(Sender).name='SO_AFFCO1VALEUR') then begin ii:=1; cpt :=GetParamSoc('SO_AFFCO1VALEUR'); TypeAff:='AFF';end
  else begin ii := 0; cpt := 0; end;

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

function TFAssistCodeaffaire.TestCptAff (Num : String; TypeAff:string;ii:integer): boolean;
Var QQ : TQuery ;
sql:string;
Begin
result:=true;
sql:=  'SELECT aff_Affaire FROM AFFAIRE';
if (SO_AFFPRODIFFERENT.Checked) then sql:=sql+ ' WHERE AFF_STATUTAFFAIRE="'+TypeAff+'"';
if not(ExisteSQL(sql)) then exit;
sql:=  'SELECT MAX(AFF_AFFAIRE'+ InttoStr(ii)+') FROM AFFAIRE ';
if (SO_AFFPRODIFFERENT.Checked) then sql:=sql+' WHERE AFF_STATUTAFFAIRE="'+TypeAff+'"';
QQ:=OpenSQL(sql,TRUE,-1,'',true);
if Not QQ.EOF then
   if (StrtoInt(Num))> QQ.Fields[0].AsInteger
      then result:=true
      else result:=false;
Ferme(QQ);
end;


procedure TFAssistCodeaffaire.FormatValeurPartie (NumPart,Lng : integer; Propos : Boolean);
Var CC : THEdit;
    st,suff : String;
    Combo : THValComboBox;
    Cpt : integer;
BEGIN
if Propos then suff := 'PRO' else suff := '';
if Propos And Not(SO_AFFPRODIFFERENT.Checked) then Exit;
Combo := THValComboBox(findComponent('SO_AFFCO'+ IntToStr(NumPart)+'TYPE'));
CC := THEdit(findComponent('SO_AFFCO'+ IntToStr(NumPart)+'VALEUR'+Suff));
if CC <> Nil then CC.MaxLength := Lng else exit;
if Combo = Nil then Exit;
if CC.Text = '' then CC.Text := '0';
if pos(Combo.Value, 'CPT;CPM')>0 then // Formattage avec des 0
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

procedure TFAssistCodeaffaire.SO_AFFPRODIFFERENTClick(Sender: TObject);
Var ProVisible : Boolean;
begin
  inherited;
ProVisible := SO_AFFPRODIFFERENT.Checked;
SO_AFFCO1VALEURPRO.Visible := ProVisible; TSO_AFFCO1VALEURPRO.Visible := ProVisible;
SO_AFFCO2VALEURPRO.Visible := ProVisible; TSO_AFFCO2VALEURPRO.Visible := ProVisible;
SO_AFFCO3VALEURPRO.Visible := ProVisible; TSO_AFFCO3VALEURPRO.Visible := ProVisible;
end;

function TFAssistCodeAffaire.PageCount : Integer ;
BEGIN
result := SO_AFFCODENBPARTIE.Value + 2;  //gm 13/01/03
END ;

Procedure TFAssistCodeAffaire.razCompteur (CC: THEDIT;Cache : Boolean);
BEGIN
if CC = nil then Exit;
if Cache then
   BEGIN
   CC.Enabled := False; CC.Color := clBtnFace; CC.Text := '';
   END
else
   BEGIN
   CC.Enabled := True;  CC.Color := clwindow;
   END;
END;

procedure TFAssistCodeaffaire.bAideClick(Sender: TObject);
begin
  inherited;
CallHelpTopic(Self) ;
end;

end.
