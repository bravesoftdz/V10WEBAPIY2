unit UTofBTEtatsAffaire;

interface
uses  StdCtrls,Controls,Classes,db,forms,sysutils,dbTables,ComCtrls,AGLInit,SaisUtil,HPanel,
      HCtrls,HEnt1,HMsgBox,UTOF, mul,FactUtil,AffaireUtil,ActiviteUtil,utob,grids, DicoAF, HSysMenu,HTB97,utofbaseetats;
Type
     TOF_BTETATS_AFF = Class (TOF_BASE_ETATS)
      procedure OnArgument(stArgument : String ) ; override ;
      procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
      procedure ComboRupture1Change(Sender: TObject); override ;
      procedure ComboRupture2Change(Sender: TObject); override ;
      procedure ComboRupture3Change(Sender: TObject); override ;
      procedure ComboRupture4Change(Sender: TObject); override ;
      procedure ComboRupture5Change(Sender: TObject); override ;
      procedure TableauObjectsInvisibles(Crit:string; var iNbChamps:integer; var tbChamps:PString); override;
      procedure DateTraitementChange(Sender: TObject);
     public
      TMoisCloture   :  THEdit;
      TMoisCloture_  :  THEdit;
      TTMoisCloture  :  THLabel;
      TTMoisCloture_ :  THLabel;
      DateDebExer    :  THEdit;
      DateDebExer_   :  THEdit;
      TDateDebExer   :  THLabel;
      TDateDebExer_  :  THLabel;
      DateFinExer    :  THEdit;
      DateFinExer_   :  THEdit;
      TDateFinExer   :  THLabel;
      TDateFinExer_  :  THLabel;
      TabSheetChampParam : TTabSheet;
      DateTraitement :  THEdit;
      DateFinGener   :  THEdit;
     END ;

const
NbTbChRessInv = 13 ;
TbChampsRessInvisibles : array[1..NbTbChRessInv] of string 	= (
          {1}        'ARS_LIBELLE',
          {2}        'TARS_LIBELLE',
          {3}        'ARS_LIBELLE_',
          {4}        'TARS_LIBELLE_',
          {5}        'ARS_FONCTION1',
          {6}        'TARS_FONCTION1',
          {7}        'ARS_FONCTION1_',
          {8}        'TARS_FONCTION1_',
          {9}        'ARS_TYPERESSOURCE',
          {10}       'TARS_TYPERESSOURCE',
          {11}       'STATISTIQUES',
          {12}       'COMPLEMENT',
          {13}       'RUPTURE'
          );

NbTbChValoInv = 10 ;
TbChampsValoInvisibles : array[1..NbTbChValoInv] of string 	= (
          {1}        'ARS_PVHT',
          {2}        'TARS_PVHT',
          {3}        'TARS_PVHT_',
          {4}        'ARS_PVHT_',
          {5}        'TARS_TAUXREVIENTUN',
          {6}        'ARS_TAUXREVIENTUN',
          {7}        'TARS_TAUXREVIENTUN_',
          {8}        'ARS_TAUXREVIENTUN_',
          {9}        'TXX_VARIABLE1',
          {10}       'XX_VARIABLE1'
          );

implementation

procedure TOF_BTETATS_AFF.OnArgument(stArgument : String );
begin
// Suppression de l'onglet Champs paramétrables en Gestion d'affaire
TabSheetChampParam  := TTabSheet(GetControl('COL_VARIABLE'));

// Suppression des champs mois de cloture en Gestion d'Affaire
TMoisCloture := THEdit(GetControl('T_MOISCLOTURE'));
TMoisCloture_ := THEdit(GetControl('T_MOISCLOTURE_'));
TTMoisCloture := THLabel(GetControl('TT_MOISCLOTURE'));
TTMoisCloture_ := THLabel(GetControl('TT_MOISCLOTURE_'));

if Not (ctxScot in V_PGI.PGIContexte) then
   begin
   if (TMoisCloture<>nil) then TMoisCloture.visible:=false;
   if (TMoisCloture_<>nil) then TMoisCloture_.visible:=false;
   if (TTMoisCloture<>nil) then TTMoisCloture.visible:=false;
   if (TTMoisCloture_<>nil) then TTMoisCloture_.visible:=false;
   if (TabSheetChampParam<>nil) then TabSheetChampParam.TabVisible:=false;
   end;


// Suppression des champs date debut d'exercice en Gestion d'Affaire
DateDebExer := THEdit(GetControl('AFF_DATEDEBEXER'));
DateDebExer_ := THEdit(GetControl('AFF_DATEDEBEXER_'));
TDateDebExer := THLabel(GetControl('TAFF_DATEDEBEXER'));
TDateDebExer_ := THLabel(GetControl('TAFF_DATEDEBEXER_'));

if Not (ctxScot in V_PGI.PGIContexte) then
   begin
   if (DateDebExer<>nil) then DateDebExer.visible:=false;
   if (DateDebExer_<>nil) then DateDebExer_.visible:=false;
   if (TDateDebExer<>nil) then TDateDebExer.visible:=false;
   if (TDateDebExer_<>nil) then TDateDebExer_.visible:=false;
   end;

// Suppression des champs date fin d'exercice en Gestion d'Affaire
DateFinExer := THEdit(GetControl('AFF_DATEFINEXER'));
DateFinExer_ := THEdit(GetControl('AFF_DATEFINEXER_'));
TDateFinExer := THLabel(GetControl('TAFF_DATEFINEXER'));
TDateFinExer_ := THLabel(GetControl('TAFF_DATEFINEXER_'));

if Not (ctxScot in V_PGI.PGIContexte) then
   begin
   if (DateFinExer<>nil) then DateFinExer.visible:=false;
   if (DateFinExer_<>nil) then DateFinExer_.visible:=false;
   if (TDateFinExer<>nil) then TDateFinExer.visible:=false;
   if (TDateFinExer_<>nil) then TDateFinExer_.visible:=false;
   end;




Inherited;

// On traite la suppression de T_MOISCLOTURE dans les combo associées aux ruptures
if Not (ctxScot in V_PGI.PGIContexte) then
   begin
   DeleteItemDansCombo(ComboRupture1, 'T_MOISCLOTURE');
   DeleteItemDansCombo(ComboRupture2, 'T_MOISCLOTURE');
   DeleteItemDansCombo(ComboRupture3, 'T_MOISCLOTURE');
   DeleteItemDansCombo(ComboRupture4, 'T_MOISCLOTURE');
   DeleteItemDansCombo(ComboRupture5, 'T_MOISCLOTURE');
   end;

// Gestion du cochage des champs RuptAff si la combo XX_RUPTURE = GP_AFFAIRE
CocherSiValeurDansCombo(RuptAff1, 'GP_AFFAIRE', ComboRupture1);
CocherSiValeurDansCombo(RuptAff2, 'GP_AFFAIRE', ComboRupture2);
CocherSiValeurDansCombo(RuptAff3, 'GP_AFFAIRE', ComboRupture3);
CocherSiValeurDansCombo(RuptAff4, 'GP_AFFAIRE', ComboRupture4);
CocherSiValeurDansCombo(RuptAff5, 'GP_AFFAIRE', ComboRupture5);

// Gestion de la synchronisation entre le champ DateTraitement et DateFinGener
DateTraitement := THEdit(GetControl('DATETRAITEMENT'));
DateFinGener := THEdit(GetControl('AFF_DATEFINGENER'));
if (DateTraitement<>nil) and (DateFinGener<>nil) then
    begin
    DateTraitementChange(nil);
    DateTraitement.OnChange:=DateTraitementChange;
    end;

End;

procedure TOF_BTETATS_AFF.DateTraitementChange(Sender: TObject);
begin
DateFinGener.Text := DateTraitement.Text;
end;

procedure TOF_BTETATS_AFF.ComboRupture1Change(Sender: TObject);
begin
Inherited;
CocherSiValeurDansCombo(RuptAff1, 'GP_AFFAIRE', ComboRupture1);
end;

procedure TOF_BTETATS_AFF.ComboRupture2Change(Sender: TObject);
begin
Inherited;
CocherSiValeurDansCombo(RuptAff2, 'GP_AFFAIRE', ComboRupture2);
end;

procedure TOF_BTETATS_AFF.ComboRupture3Change(Sender: TObject);
begin
Inherited;
CocherSiValeurDansCombo(RuptAff3, 'GP_AFFAIRE', ComboRupture3);
end;

procedure TOF_BTETATS_AFF.ComboRupture4Change(Sender: TObject);
begin
Inherited;
CocherSiValeurDansCombo(RuptAff4, 'GP_AFFAIRE', ComboRupture4);
end;

procedure TOF_BTETATS_AFF.ComboRupture5Change(Sender: TObject);
begin
Inherited;
CocherSiValeurDansCombo(RuptAff5, 'GP_AFFAIRE', ComboRupture5);
end;

procedure TOF_BTETATS_AFF.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('AFF_AFFAIRE'));
Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
Aff_:=THEdit(GetControl('AFF_AFFAIRE_'));
Aff0_:=THEdit(GetControl('AFF_AFFAIRE0_'));
Aff1_:=THEdit(GetControl('AFF_AFFAIRE1_'));
Aff2_:=THEdit(GetControl('AFF_AFFAIRE2_'));
Aff3_:=THEdit(GetControl('AFF_AFFAIRE3_'));
Aff4_:=THEdit(GetControl('AFF_AVENANT_'));
Tiers:=THEdit(GetControl('AFF_TIERS'));
Tiers_:=THEdit(GetControl('AFF_TIERS_'));
end;

procedure TOF_BTETATS_AFF.TableauObjectsInvisibles(Crit:string; var iNbChamps:integer; var tbChamps:PString);
begin
if (Crit='RESS') then
    begin
    iNbChamps := NbTbChRessInv;
    tbChamps := @TbChampsRessInvisibles;
    end;

if (Crit='VALO') then
    begin
    iNbChamps := NbTbChValoInv;
    tbChamps := @TbChampsValoInvisibles;
    end;
end;

Initialization
registerclasses([TOF_BTETATS_AFF]);
end.
