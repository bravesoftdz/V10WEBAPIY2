unit AssistCreationAffaire;

interface

// mcd 02/01 revue pour utiliser les mêmes fct lors de la duplication N en N+1

uses
  Windows, Messages,SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
{$IFDEF EAGLCLIENT}
	emul,MaineAGL,
{$ELSE}
 {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB,   Fe_Main,Mul, // gmeagl
{$ENDIF}
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
  Buttons, Mask, AffaireUtil,M3FP, UTOB,Dicobtp,UTOF,AGLInit, FactTOB,
  HEnt1,Hstatus,EntGC,Facture,FactComm,UtilGc,TiersUtil,FactUtil,UtilPGI,SaisUtil,
  UTOFAFTRADUCCHAMPLIBRE,AffaireDuplic,ParamSoc , Hqry,UtofAftiers_mul,UtilMulTrt,AGLInitGc,
  HPanel,uEntCommun,UtilTOBPiece;

type
  TFAssistCreationAffaire = class(TFAssist)
    TSCODIFICATION: TTabSheet;
    BevelDecoupe: TBevel;
    LibSelectModele: THLabel;
    libselectModele2: THLabel;
    LibWelcome1: THLabel;
    AFFAIREMODELE: THCritMaskEdit;
    HLabel1: THLabel;
    HLabel2: THLabel;
    LibSelectClient: THLabel;
    LibSelectClient2: THLabel;
    AFF_TIERS: THCritMaskEdit;
    CBMultiClient: TCheckBox;
    BtSelectMultiClient: THBitBtn;
    TSCaracteristique: TTabSheet;
    NBMULTICLIENT: THLabel;
    LibSelectDates: THLabel;
    TAFF_DATEDEBUT: THLabel;
    AFF_DATEDEBUT: THCritMaskEdit;
    TAFF_DATEFIN: THLabel;
    AFF_DATEFIN: THCritMaskEdit;
    HLabel3: THLabel;
    Bevel1: TBevel;
    Bevel2: TBevel;
    Bevel3: TBevel;
    HLabel4: THLabel;
    HLabel5: THLabel;
    LibModeFact: THLabel;
    LibLModeFact2: THLabel;
    TAFF_MONTANTECHE: THLabel;
    LibLModeFact3: THLabel;
    AFF_GENERAUTO: TLabel;
    AFF_PERIODICITE: TLabel;
    AFF_MONTANTECHE: THCritMaskEdit;
    AFF_AFFAIRE0: THCritMaskEdit;
    AFF_AFFAIRE1: THCritMaskEdit;
    AFF_AFFAIRE2: THCritMaskEdit;
    AFF_AFFAIRE3: THCritMaskEdit;
    AFF_AVENANT: THCritMaskEdit;
    BRechAffaire: TToolbarButton97;
    CalculMoisCloture: TCheckBox;
    PAvancee: TTabSheet;
    Bevel4: TBevel;
    Bevel5: TBevel;
    Escompteclient: TCheckBox;
    LibRepriseclientdestination: THLabel;
    ModeRegleClient: TCheckBox;
    LibParamavance: THLabel;
    LibInfoComplement: THLabel;
    TAFF_RESPONSABLE: THLabel;
    AFF_RESPONSABLE: THCritMaskEdit;
    AFF_DEPARTEMENT: THValComboBox;
    TAFF_DEPARTEMENT: THLabel;
     CalculExercice: TCheckBox;
    AFF_TIERSMODELE: THCritMaskEdit;
    LibModeleTiers: THLabel;
    SurAffModele: TCheckBox;
    TAFF_DATEDEBGENER: THLabel;
    AFF_DATEDEBGENER: THCritMaskEdit;
    TAFF_DATEFINGENER: THLabel;
    AFF_DATEFINGENER: THCritMaskEdit;
    bDateGener: TCheckBox;
    LA_Taches: THLabel;
    CB_DupliquerTache: TCheckBox;
    RepFactAff: TCheckBox;
    RepEtatAff: TCheckBox;

    procedure CBMultiClientClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BtSelectMultiClientClick(Sender: TObject);
    procedure AFF_TIERSElipsisClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure BRechAffaireClick(Sender: TObject);
    procedure AFF_AFFAIRE1Exit(Sender: TObject);
    procedure CalculMoisClotureClick(Sender: TObject);
    procedure AFF_TIERSMODELEElipsisClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure bDateGenerClick(Sender: TObject);
    procedure RepFactAffClick(Sender: TObject);
    procedure bAideClick(Sender: TObject);
    procedure AFF_RESPONSABLEElipsisClick(Sender: TObject);
  private
    { Déclarations privées }
    StatutAffaire,StatutReduit : string;
    ListeClients : TStringList;
    bProposition, bMoisCloture,bExercice : boolean;

    Function CreationAffaires : Boolean;
    Procedure PrechargementAffaireModele;
    Function AlimAffaireModele(AvecPrecharge : Boolean): Boolean;
    procedure AffichePlanning;
  public
    { Déclarations publiques }
             //mcd mis en public pour duplication N en N+1, ab Listeclients en tob
    TobAffModele,TobListeClients : TOB;
    CleDocModele : R_CLEDOC;
    LastAffaire : String;
                // fin mcd
  end;


// Appel de l'assistant de création
Function LanceAssistCreationAffaire(TypeAff : string): string;
// Fonction de gestion des pieces et tob associées
procedure ToutAllouerPiece (var TOBPiece, TOBBases, TOBEches,TobNomenclature, TobPorcs : TOB);
procedure ToutLibererPiece (var TOBPiece, TOBBases, TOBEches,TobNomenclature, TobPorcs : TOB);
Function LoadLesTobPiece (Cledoc : R_CLEDOC; TOBPiece, TOBBases, TOBEches, TobPorcs : TOB) : Boolean;


// libellés des messages de la TOM Affaire
Const	MsgCreat: array[1..14] of string 	= (
          {1}        'Aucune affaire modèle sélectionnée',
          {2}        'Aucun client sélectionné',
          {3}        'La création d''affaire(s) a échoué',
          {4}        'La création d''affaire(s) a réussi',
          {5}        'Assistant de création d''affaires',
          {6}        'Affaire modèle invalide',
          {7}        'Date de début supérieure à la date de fin',
          {8}        'Nombre de pièces associées à l''affaire incorrect',
          {9}        'Chargement de la pièce modèle impossible',
          {10}       'Renumérotation de la pièce impossible',
          {11}       'Code affaire déja utilisé pour ce client',
          {12}       'Problème rencontré lors de la validation',
          {13}       'Nouveau code affaire incorrect',
          {14}       'Tiers non valide'
          );

// TOF de sélection Multi-Tiers
type
     TOF_TIERS_MULTI = Class (TOF_AFTRADUCCHAMPLIBRE)
        procedure OnArgument(Arguments : String) ; override ;
        Procedure SelectionMultiClient ;
        Procedure SelectionListeClients ;
     END ;

implementation
{$DEFINE ABTEST} // modif du 13/12/2002
{$R *.DFM}
Function LanceAssistCreationAffaire(TypeAff : string): string;
Var FAssistCreationAffaire  : TFAssistCreationAffaire ;
begin
FAssistCreationAffaire:=TFAssistCreationAffaire.Create(Application) ;
if TypeAff = '' then TypeAff := 'AFF';
FAssistCreationAffaire.StatutAffaire := TypeAff;
FAssistCreationAffaire.StatutReduit  := StatutCompletToReduit(TypeAff);
if (TypeAff='AFF') then FAssistCreationAffaire.bProposition := False
                   else FAssistCreationAffaire.bProposition := True;

FAssistCreationAffaire.ShowModal ;
Result := FAssistCreationAffaire.LastAffaire;
FAssistCreationAffaire.Free;
end;



Function AGLAssistCreationAffaire( parms: array of variant; nb: integer ): variant;
begin
result := LanceAssistCreationAffaire(Parms[0]);
end;

// *****************************************************************************
// ***************************** Gestion Evènements ****************************
// *****************************************************************************
procedure TFAssistCreationAffaire.CBMultiClientClick(Sender: TObject);
Var VoirMono, VoirMulti : Boolean;
begin
  inherited;
if CBMulticlient.Checked then BEGIN VoirMono := False; VoirMulti := True;  END
                         else BEGIN VoirMono := True;  VoirMulti := False; END;

LibSelectClient2.Visible := VoirMono; AFF_TIERS.Visible := VoirMono;
BtSelectMultiClient.Visible := VoirMulti; NBMULTICLIENT.Visible := VoirMulti;
if VoirMulti then AFF_TIERS.Text := '';
// On vide la liste des clients
ListeClients.Clear;
end;

procedure TFAssistCreationAffaire.FormShow(Sender: TObject);
begin
  inherited;
  AffichePlanning;

  {$IFDEF CCS3}
  SurAffModele.Visible := false; SurAffModele.Checked := false;
  {$ENDIF}
  Bfin.enabled:=False;
  ListeClients := TStringList.Create;
  CBMultiClientClick (Nil);
  ChargeCleAffaire(AFF_AFFAIRE0,AFF_AFFAIRE1,AFF_AFFAIRE2,AFF_AFFAIRE3,AFF_AVENANT,nil,taModif,AFFAIREMODELE.Text,False);
     // mcd 16/10/01if ctxScot in V_PGI.PGIContexte then
  if (GetParamSoc('SO_AfFormatExer') <> 'AUC') then
     BEGIN
     LibSelectDates.caption := '"Dates de l''exercice client" de la fiche mission';
     CalculMoisCloture.Visible := True; CalculMoisCloture.Checked := True;
     CalculExercice.Visible := True; CalculExercice.Checked := True;
     bDateGener.Visible := False;
      // mcd 17/10/01 on cahce les date de factures si pas calcul date
     Aff_DAteFInGener.visible:=False;
     TAff_DAteFInGener.visible:=False;
     Aff_DAteDebGener.visible:=False;
     TAff_DAteDebGener.visible:=False;
     LibSelectDates.visible:=False;   //mcd 27/03/02 libellés à cacher aussi
     END
  else
     BEGIN
     CalculMoisCloture.Visible := False;  CalculExercice.Visible := False;
     bDateGener.Checked := True;
     bDateGenerClick(Nil);
     END;
  LastAffaire :='';
end;

procedure TFAssistCreationAffaire.bSuivantClick(Sender: TObject);
Var Onglet : TTabSheet;
    St_NomOnglet : string;
    Q : Tquery;
begin
Onglet := P.ActivePage;  // Activation de la page en cours
st_NomOnglet := Onglet.Name;
if st_NomOnglet = 'TSCODIFICATION' then
    BEGIN
    Bfin.enabled:=False;
    if Not(CBMulticlient.Checked) then
       BEGIN
       ListeClients.Clear;
       if AFF_TIERS.Text <> '' then
         begin
         // test client saisi valide
{$IFDEF ABTEST}
       if (TobListeClients<>nil) then TobListeClients.cleardetail
       else TobListeClients := Tob.create ('Liste Clients',nil,-1);
       Q := OpenSQL ('SELECT T_TIERS FROM TIERS WHERE T_TIERS="'+AFF_TIERS.Text+'" AND '+FabriqueWhereNatureAuxiAff(StatutAffaire),True,-1,'',true);
       if Not (Q.EOF) then TobListeClients.LoadDetailDB('LIGNE CLIENTS','','',Q,False,False) ;
       Ferme(Q);
       if (TobListeClients.detail.count = 0) then
            begin
            PgiBoxAf (MsgCreat[14],'Sélection des tiers');
            AFF_TIERS.SetFocus; Exit;
            end;
{$ELSE}
         if not(ExisteSQL('SELECT T_TIERS FROM TIERS WHERE T_TIERS="'+AFF_TIERS.Text+'"')) then
            begin
            PgiBoxAf (MsgCreat[14],'Sélection des tiers');
            AFF_TIERS.SetFocus; Exit;
            end;
         ListeClients.Add (AFF_TIERS.Text);
{$ENDIF}
         end;
       END;
    // Test que l'affaire modèle est bien renseignée
    if (AFF_AFFAIRE1.text = '') then
       BEGIN
       PgiBoxAf (MsgCreat[1],'Sélection d''une affaire modèle');
       AFF_AFFAIRE1.SetFocus; Exit;
       END;
    // Test que le client est renseigné
{$IFDEF ABTEST}
    if (TobListeClients = nil) or ((TobListeClients <> nil) and (TobListeClients.detail.Count = 0)) then
{$ELSE}
    if ListeClients.Count = 0 then
{$ENDIF}
       BEGIN
       PgiBoxAf (MsgCreat[2],'Sélection d''un client de destination'); Exit;
       END;
    // Test que l'affaire modèle est correcte
    if Not(AlimAffaireModele(True)) then
       BEGIN
       PgiBoxAf (MsgCreat[6],MsgCreat[5]); Exit;
       END
    END
else if( st_NomOnglet = 'TSCaracteristique' )then  Bfin.enabled:=True
else     Bfin.enabled:=False;
inherited; // si pas de problème on continue ...
end;

procedure TFAssistCreationAffaire.PrechargementAffaireModele;
Var Lib,TypeGener,Periode : String;
    NbMois : integer;
BEGIN
// Préchargement des zones modifiables / affaire modèle.
AFF_RESPONSABLE.Text    := TobAFFModele.GetValue('AFF_RESPONSABLE');
AFF_DEPARTEMENT.Value   := TobAFFModele.GetValue('AFF_DEPARTEMENT');
if (GetParamSoc('SO_AfFormatExer') <> 'AUC') then begin
  AFF_DATEDEBUT.Text      := TobAFFModele.GetValue('AFF_DATEDEBEXER');
  AFF_DATEFIN.Text        := TobAFFModele.GetValue('AFF_DATEFINEXER');
   end
else begin
  AFF_DATEDEBUT.Text      := TobAFFModele.GetValue('AFF_DATEDEBUT');
  AFF_DATEFIN.Text        := TobAFFModele.GetValue('AFF_DATEFIN');
  end;

  AFF_DATEDEBGENER.Text   := TobAFFModele.GetValue('AFF_DATEDEBGENER');
AFF_DATEFINGENER.Text   := TobAFFModele.GetValue('AFF_DATEFINGENER');

    // gestion des modes de facturation ...
TypeGener := TobAFFModele.GetValue('AFF_GENERAUTO');
AFF_GENERAUTO.Caption:=RechDom('AFTGENERAUTO',TypeGener,False);

TAFF_MONTANTECHE.Visible := True; AFF_MONTANTECHE.Visible := True; // par défaut
if TypeGener = 'FOR' then TAFF_MONTANTECHE.Caption := 'Montant forfaitaire' else
if TypeGener = 'POU' then TAFF_MONTANTECHE.Caption := 'Pourcentage' else
if TypeGener = 'POT' then
   BEGIN TAFF_MONTANTECHE.Visible := False; AFF_MONTANTECHE.Visible := false END
else TAFF_MONTANTECHE.Caption := 'Montant forfaitaire';

NbMois := TobAFFModele.GetValue('AFF_INTERVALGENER');
Periode := TobAFFModele.GetValue('AFF_PERIODICITE');
if NbMois = 0 then Lib := 'Non défini' else
if NbMois = 1 then
   BEGIN
   if Periode ='A' then Lib := 'Annuelle' else
   if Periode ='M' then Lib := 'Mensuelle' else
   if Periode ='P' then Lib := 'Ponctuelle' else  // 03/05/02
   if Periode ='S' then Lib := 'Hebdomadaire';
   END
else // > 1 mois
     BEGIN
     if Periode ='A' then Lib := Format('Tous les %d ans',[NbMois]) else
     if Periode ='M' then Lib := Format('Tous les %d mois',[NbMois]) else
     if Periode ='S' then Lib := Format('Toutes les %d semaines',[NbMois]);
     END;
AFF_PERIODICITE.Caption := Lib;
if TypeGener = 'FOR' then
   AFF_MONTANTECHE.Text:=TobAffModele.GetValue('AFF_MONTANTECHEDEV') else
if TypeGener = 'POU' then
   AFF_MONTANTECHE.Text:=TobAffModele.GetValue('AFF_POURCENTAGE') else
if TypeGener = 'POT' then    // ras toujours à 100 % non modifiable
else AFF_MONTANTECHE.Text:=TobAffModele.GetValue('AFF_MONTANTECHEDEV');
END;

procedure TFAssistCreationAffaire.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  inherited;
ListeClients.Free;
if TobAffModele <> Nil then TobAffModele.Free;
if TobListeClients <> Nil then TobListeClients.Free;
end;

procedure TFAssistCreationAffaire.BtSelectMultiClientClick(
  Sender: TObject);
begin
  inherited;
{$IFDEF ABTEST}
  if (TobListeClients<>nil) then TobListeClients.free;
  AFLanceFiche_Mul_tiers_Multi('','WHERE='+FabriqueWhereNatureAuxiAff(StatutAffaire));
  TobListeClients := TheTob;
  TheTob := Nil;
 if (TobListeClients <> Nil) and (TobListeClients.detail.Count> 0) then
 NBMultiClient.Caption := Format(' %d Client(s) sélectionné(s)',[TobListeClients.detail.Count])
 else NBMultiClient.Caption := 'Aucun client sélectionné';
{$ELSE}
// Sélection multi-tiers
// mc d12/06/02 AGLLanceFiche('AFF','AFTIERS_MULTI','','','WHERE='+FabriqueWhereNatureAuxiAff(StatutAffaire));
AFLanceFiche_Mul_tiers_Multi('','WHERE='+FabriqueWhereNatureAuxiAff(StatutAffaire));
if TheTOB <> Nil then
   BEGIN
   Listeclients.Clear;
   Listeclients.Free;
   ListeClients := TStringList(TheTOB.Data);
   if ListeClients<> Nil then
      BEGIN
      if ListeClients.Count <> 0 then NBMultiClient.Caption := Format(' %d Client(s) sélectionné(s)',[ListeClients.Count])
                                 else NBMultiClient.Caption := 'Aucun client sélectionné';
      END;
   TheTob.free;
   TheTob := Nil;
   END;
{$ENDIF}
end;

procedure TFAssistCreationAffaire.BRechAffaireClick(Sender: TObject);
begin
  inherited;
GetAffaireEntete(AFFAIREMODELE,AFF_AFFAIRE1,AFF_AFFAIRE2,AFF_AFFAIRE3,AFF_AVENANT,AFF_TIERSMODELE,bProposition,false,SurAffModele.Checked,True,true);
ChargeCleAffaire(AFF_AFFAIRE0,AFF_AFFAIRE1,AFF_AFFAIRE2,AFF_AFFAIRE3,AFF_AVENANT,NIL,taModif,AFFAIREMODELE.Text,False);
end;

procedure TFAssistCreationAffaire.AFF_AFFAIRE1Exit(Sender: TObject);
Var ip : integer;
    CCError : THCritMaskEdit;
begin
  inherited;
AFFAIREMODELE.Text := DechargeCleAffaire(AFF_AFFAIRE0,AFF_AFFAIRE1,AFF_AFFAIRE2,AFF_AFFAIRE3,AFF_AVENANT,AFF_TIERSMODELE.Text,taModif,False,True,bProposition,iP);
if ip <> 0 then
   BEGIN
   CCError := THCritMaskEdit(FindComponent ('AFF_AFFAIRE'+ intToStr(ip)));
   if CCError <> Nil then
      if CCerror.Visible then CCError.SetFocus;
   END
else
   BEGIN
   TestCleAffaire(AFFAIREMODELE,AFF_AFFAIRE1,AFF_AFFAIRE2,AFF_AFFAIRE3,AFF_AVENANT,AFF_TIERSMODELE, AFF_AFFAIRE0.Text, True, True,bProposition);
   END;
end;

procedure TFAssistCreationAffaire.CalculMoisClotureClick(Sender: TObject);
Var bvoirDates : Boolean;
begin
  inherited;
bVoirDates := Not(CalculMoisCloture.Checked);
AFF_DATEDEBUT.Visible := bVoirDates; TAFF_DATEDEBUT.Visible := bVoirDates;
AFF_DATEFIN.Visible := bVoirDates; TAFF_DATEFIN.Visible := bVoirDates;
LibSelectDates.Visible:=BvoirDates;//mcd 27/03/02
CalculExercice.Visible := Not(bVoirdates);
if (VH_GC.AFFORMATEXER <> 'AM') then CalculExercice.Visible := False;
end;

procedure TFAssistCreationAffaire.bFinClick(Sender: TObject);
BEGIN
inherited;
NextControl(self);
CreationAffaires;
if ExisteSql ('SELECT GOB_CODE FROM PARAMOBLIG WHERE GOB_CODE="A08" AND GOB_OBLIGATOIRE="X"')
  then PgiBoxAf ('Pensez à modifier les affaires créees si vous utilisez l''option de champs obligatoires,surtout si vos modèles ne contiennent pas tous les champs obligatoires','Création affaires'); // mcd 17/03/03
Close;
END;


// *****************************************************************************
// ***************************** Traitement de création ************************
// *****************************************************************************
Function TFAssistCreationAffaire.CreationAffaires : Boolean;
Var i : integer;
    CodeClient,Arg : String;
    TraitementOK: Boolean;
    //NbPiece : Integer;
    DateDebut, DateFin, DateDebGener, DateFinGener : TDateTime;
BEGIN
  Result := False;
  TraitementOk := True;
  arg := '';
  if Not(AlimAffaireModele(False)) then Exit;
  // Préparation des arguments
  // Traitement des dates d'affaires
  if (AFF_DATEDEBUT.Text <> '') then DateDebut := StrToDate(AFF_DATEDEBUT.Text)
                                else DateDebut := iDate1900;
  if (AFF_DATEFIN.Text <> '')   then DateFin   := StrToDate(AFF_DATEFIN.Text)
                                else DateFin   := iDate2099;

  //DateDebGener := iDate1900;  DateFinGener   := iDate2099;
  //if Not(bDateGener.Checked) then
  //   begin
     DateDebGener := StrToDate(AFF_DATEDEBGENER.Text);
     DateFinGener := StrToDate(AFF_DATEFINGENER.Text);
  //   end

  bMoisCloture :=  CalculMoisCloture.Checked; bExercice := CalculExercice.Checked;
  if DateDebut <> idate1900 then arg:=arg + 'AFF_DATEDEBUT:'+ DateTostr (DateDebut)+';';
  if DateFin   <> idate2099 then arg:=arg + 'AFF_DATEFIN:'  + DateTostr (DateFin)  +';';

  if DateDebGener <> idate1900 then arg:=arg + 'AFF_DATEDEBGENER:'+ DateTostr (DateDebGener)+';';
  if DateFinGener <> idate2099 then arg:=arg + 'AFF_DATEFINGENER:'+ DateTostr (DateFinGener)+';';

  if bMoisCloture then arg := arg + 'DATEAFFSURMOISCLOT:X;' ;
  if bExercice    then arg := arg + 'EXERSURMOISCLOT:X;' ;
  if Escompteclient.Checked  then arg := arg + 'ESCOMPTETIERS:X;' ;
  if ModeRegleClient.Checked then arg := arg + 'MODEREGLETIERS:X;' ;
  if (AFF_MONTANTECHE.Text <> '') then
     if strTofloat(AFF_MONTANTECHE.Text)<>0 then
        arg:=arg+'AFF_MONTANTECHEDEV:'+AFF_MONTANTECHE.Text+';';
  arg:= arg + 'AFF_RESPONSABLE:'+ AFF_RESPONSABLE.Text +';';
  arg:= arg + 'AFF_DEPARTEMENT:'+ AFF_DEPARTEMENT.Value+';';

{$IFDEF ABTEST}
  InitMove(TobListeClients.detail.count,'');
  // Boucle sur tous les clients pour création des affaires
  for i := 0 to TobListeClients.detail.count-1 do
      BEGIN
      TraitementOk := True;
      CodeClient := TobListeClients.detail[i].getvalue('T_TIERS');
{$ELSE}
  InitMove(ListeClients.count,'');
  // Boucle sur tous les clients pour création des affaires

  for i := 0 to ListeClients.count-1 do
      BEGIN
      TraitementOk := True;
      CodeClient := ListeClients[i];
{$ENDIF}
      // Ajout argument du nouveau client
      if TobAffModele.GetValue('AFF_TIERS') <> CodeClient then
        arg:=arg + 'AFF_TIERS:'+CodeClient+';';
      LastAffaire := DuplicationAffaire (tdaAssistCreat, TobAffModele.GetValue('AFF_AFFAIRE'), Arg, TobAffModele,True,RepFactaff.Checked,RepEtataff.Checked);

      // C.B duplication des taches
      if VH_GC.GAPlanningSeria and (LastAffaire <> '') and (CB_DupliquerTache.Checked) then
        DuplicationTache(TobAffModele.GetValue('AFF_AFFAIRE'), LastAffaire);

      MoveCur (False);
      END;
  FiniMove ();
  Result := TraitementOK;
  // On retourne le code affaire généré pour rentrer en auto sur cette affaire
{$IFDEF ABTEST}
  if TobListeClients.detail.count<>1 then LastAffaire := '';
{$ELSE}
  if ListeClients.count<>1 then LastAffaire := '';
{$ENDIF}


END;

// *****************************************************************************
// ***************************** Gestion Affaire modèle ************************
// *****************************************************************************
Function TFAssistCreationAffaire.AlimAffaireModele(Avecprecharge : Boolean) : Boolean;
Var stSQL : string;
    Q : Tquery;
    Recharge : Boolean;
BEGIN
Result := False;
Recharge := true;
// Vérification si la TObaffaireModèle est déjà chargée
if TobAffModele <> Nil then
   BEGIN
   if TobAffModele.getValue('AFF_AFFAIRE') = AFFAIREMODELE.Text then
      BEGIN
      Recharge := False; Result := true;
      END;
   if Recharge then BEGIN TobAffModele.Free; TobAffModele := nil; END;
   END;

// Chargement de la tob affaire modèle
if Recharge then
  BEGIN       // Création d'une affaire, il faut tout prendre**
  stSQL := 'SELECT * From AFFAIRE where AFF_AFFAIRE="'+ AFFAIREMODELE.Text+'"' ;
  Q := OpenSQL (stSQL,True,-1,'',true);
  if Not (Q.EOF) then
     BEGIN
     Result := True;
     TobAffModele := TOB.Create('AFFAIRE',Nil,-1);
     TobAffModele.SelectDB ('',Q);
     if AvecPrecharge then PrechargementAffaireModele;
     END;
  Ferme(Q);
  END;
END;


procedure TFAssistCreationAffaire.AFF_TIERSElipsisClick(Sender: TObject);
begin
  inherited;
GetTiersRecherche(AFF_TIERS,FabriqueWhereNatureAuxiAff(StatutAffaire),'','');
end;


//******************************************************************************
//******************* fonction externe de gestion des pieces  ******************
//******************************************************************************
procedure ToutAllouerPiece (var TOBPiece, TOBBases, TOBEches, TobNomenclature,TobPorcs : TOB);
BEGIN
TOBPiece:=TOB.Create('PIECE',Nil,-1) ;
TOBBases:=TOB.Create('Les Bases',Nil,-1) ;  //mcd 28/01/03 chgmt nom tob
TOBEches:=TOB.Create('Les Echeances',Nil,-1) ;
TobNomenclature := TOB.Create('les Nomenclatures',Nil,-1) ; //mcd 28/01/03 chgmt nom tob
TOBPorcs:=TOB.Create('Les Ports',Nil,-1) ;     //mcd 28/01/03 chgmt nom tob
END;

procedure ToutLibererPiece(var TOBPiece, TOBBases, TOBEches, TobNomenclature,TOBPorcs : TOB);
BEGIN
if TobPiece <> Nil then BEGIN TOBPiece.Free ; TOBPiece:=Nil ; END;
if TOBEches <> Nil then BEGIN TOBEches.Free ; TOBEches:=Nil ; END;
if TOBBases <> Nil then BEGIN TOBBases.Free ; TOBBases:=Nil ; END;
if TOBPorcs <> Nil then BEGIN TOBPorcs.Free ; TOBPorcs:=Nil ; END;
if TobNomenclature <> Nil then
   BEGIN TobNomenclature.Free ; TobNomenclature:=Nil ; END;
END;

Function LoadLesTobPiece (Cledoc : R_CLEDOC; TOBPiece, TOBBases, TOBEches, TobPorcs : TOB) : Boolean;
Var Q : TQuery ;
//    i : integer ;
BEGIN
  Result := True;
         // Création d'une affaire, il faut tout prendre**
  Q:=OpenSQL('SELECT * FROM PIECE WHERE '+WherePiece(CleDoc,ttdPiece,False),True,-1,'',true) ;
  if not Q.EOF then TOBPiece.SelectDB('',Q)
               else BEGIN Result := False; Ferme(Q) ; exit; END;
  Ferme(Q) ;
  // Lecture Lignes Création d'une affaire, il faut tout prendre**
  Q:=OpenSQL('SELECT * FROM LIGNE WHERE '+WherePiece(CleDoc,ttdLigne,False)+' ORDER BY GL_NUMLIGNE',True,-1,'',true) ;
  TOBPiece.LoadDetailDB('LIGNE','','',Q,False,False) ;
  Ferme(Q) ;

  PieceAjouteSousDetail(TOBPiece);  // GM 08/02/02

  // Lecture bases  Création d'une affaire, il faut tout prendre**
  Q:=OpenSQL('SELECT * FROM PIEDBASE WHERE '+WherePiece(CleDoc,ttdPiedBase,False),True,-1,'',true) ;
  TOBBases.LoadDetailDB('PIEDBASE','','',Q,False,False) ;
  Ferme(Q) ;
  // Lecture Echéances Création d'une affaire, il faut tout prendre**
  Q:=OpenSQL('SELECT * FROM PIEDECHE WHERE '+WherePiece(CleDoc,ttdEche,False),True,-1,'',true) ;
  TOBEches.LoadDetailDB('PIEDECHE','','',Q,False,false) ;
  Ferme(Q) ;
  // Lecture Ports  Création d'une affaire, il faut tout prendre**
  Q:=OpenSQL('SELECT * FROM PIEDPORT WHERE '+WherePiece(CleDoc,ttdPorc,False),True,-1,'',true) ;
  TOBPorcs.LoadDetailDB('PIEDPORT','','',Q,False,False) ;
  Ferme(Q) ;
  // Lecture Nomenclatures  LoadLesNomen(TOBPiece,TOBNomenclature,TOBArticles,CleDoc) ;
  // Lecture Lot  LoadLesLots(TOBPiece,TOBDesLots,CleDoc) ;
  // Lecture ACompte  LoadLesAcomptes(TOBPiece,TOBAcomptes,CleDoc) ;
  // Lecture Analytiques LoadLesAna(TOBPiece,TOBAnaP) ;
END;


//******************************************************************************
//******************* Gestion du mul de sélection multi-tiers ******************
//******************************************************************************
procedure TOF_TIERS_MULTI.OnArgument(Arguments : String) ;
Var stmp:string;
       tsArgsReception : TStringList;
Begin
inherited ;
sTmp := StringReplace(Arguments, ';', chr(VK_RETURN), [rfReplaceAll]);
tsArgsReception := TStringList.Create;
tsArgsReception.Text := sTmp;
{$IFDEF ABTEST}
if pos('PRO',Arguments) > 0 then
begin
  SetControlVisible ('T_NATUREAUXI',True);
  SetControlVisible ('TT_NATUREAUXI',True);
  SetControlText('T_NATUREAUXI','CLI;PRO');
end else
{$ENDIF}
  SetCOntrolText('XX_WHERE',tsArgsReception.Values['WHERE']);

if not(ctxscot In V_PGI.PGIContexte) then
   BEGIN
   SetControlVisible ('T_MOISCLOTURE',False); SetControlVisible ('T_MOISCLOTURE_',False);
   SetControlVisible ('TT_MOISCLOTURE',False); SetControlVisible ('TT_MOISCLOTURE_',False);
   END;
End ;

Procedure TOF_TIERS_MULTI.SelectionListeClients ;
var TobListeClients :Tob;
    stWhere : string;
begin
  TobListeClients := Tob.Create('liste Clients',NIL,-1);

{$IFDEF EAGLCLIENT}
if TFMul(Ecran).bSelectAll.Down then
   if not TFMul(Ecran).Fetchlestous then
     begin
     TFMul(Ecran).bSelectAllClick(Nil);
     TFMul(Ecran).bSelectAll.Down := False;
     exit;
     end;
{$ENDIF}

  if TFMul(Ecran).FListe.AllSelected then
  begin
    stWhere := RecupWhereCritere (TPageControl(TFMul(Ecran).Pages));
    TobListeClients.LoadDetailFromSQL ('SELECT T_TIERS FROM GCVTIERSCOMPL ' + stWhere);
    TFMul(Ecran).FListe.AllSelected:=False;
  end else
    TraiteEnregMulListe (TFMul(Ecran), 'T_TIERS','GCVTIERSCOMPL', TobListeClients, True);
  TheTob := TobListeClients ;
  Ecran.Close;
end;

 Procedure TOF_TIERS_MULTI.SelectionMultiClient ;
var  F : TFMul ;
 //    L:THDBGrid;
 {$IFDEF EAGLCLIENT}  // gmeagl
       L : THGrid;
{$ELSE}
       L : THDBGrid;
{$ENDIF}
//     Q:TQuery;  //gmeagl
     Q:THQuery;
     i :integer;
     stClient : String;  
     ZListeclients : TStringList;
     TOBCom : TOB;
begin
F:=TFMul(Ecran);
L:= F.FListe; Q:= F.Q;

if(F.FListe.NbSelected=0) and (not F.FListe.AllSelected) then
   begin
   PGIInfoAf('Aucun Client sélectionné','Assistant de création d''affaires'); Exit;
   end;
// InitMove(RecordsCount(Q),'');  /gmeag
ZListeClients := TStringList.create;




{$IFDEF EAGLCLIENT}
if F.bSelectAll.Down then
   if not F.Fetchlestous then
     begin
     F.bSelectAllClick(Nil);
     F.bSelectAll.Down := False;
     exit;
     end;
{$ENDIF}

if L.AllSelected then   // Tous sélectionnés
   BEGIN
 // attention NUL, il faut passer par une tob pour ne pas faire les q.next...
   Q.First; // InitMove(RecordsCount(Q),'');   gmeagl
   while Not Q.EOF do
     BEGIN
     stClient := Q.FindField('T_TIERS').asString;
     if stClient <> '' then ZListeclients.add(stClient);
     Q.NEXT; // moveCur(False); gmeagl                          // ##PLPBNEXT ?
     END;
     L.AllSelected:=False;
   END
Else                   // Sélection multiple effectuée
  Begin
  for i:=0 to L.NbSelected-1 do
   BEGIN
 //  L.GotoLeBOOKMARK(i);      gmeagl
        F.FListe.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
        F.Q.TQ.Seek(F.FListe.Row-1) ;
{$ENDIF}

   stClient := Q.FindField('T_TIERS').asString;
   if stClient <> '' then ZListeclients.add(stClient);
   // MoveCur(False);   gmeagl
   END;
  L.ClearSelected;
  End;
// FiniMove;     gmeagl
// Envoi de la string list par l'objet global The TOB
TobCom := TOB.Create ('tob communication', nil,-1);
TheTob:=TobCom ;
TobCom.data:=ZListeClients;
Ecran.Close;
END;


procedure AGLSelectionMultiClient(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then MaTOF:=TFMul(F).LaTOF else exit;
{$IFDEF ABTEST}
if (MaTOF is TOF_TIERS_MULTI) then TOF_TIERS_MULTI(MaTOF).SelectionListeClients else exit;
{$ELSE}
if (MaTOF is TOF_TIERS_MULTI) then TOF_TIERS_MULTI(MaTOF).SelectionMultiClient else exit;
{$ENDIF}
end;




procedure TFAssistCreationAffaire.AFF_TIERSMODELEElipsisClick(Sender: TObject);
begin
  inherited;
GetTiersRecherche(AFF_TIERSMODELE,FabriqueWhereNatureAuxiAff(StatutAffaire),'','');
end;

procedure TFAssistCreationAffaire.bPrecedentClick(Sender: TObject);
begin
Bfin.enabled:=False;
inherited; // si pas de problème on continue ...

end;

procedure TFAssistCreationAffaire.bDateGenerClick(Sender: TObject);
Var VoirDate : Boolean;
begin
  inherited;
VoirDate := Not(bDateGener.Checked);
AFF_DATEDEBGENER.Visible := VoirDate; TAFF_DATEDEBGENER.Visible := VoirDate;
AFF_DATEFINGENER.Visible := VoirDate; TAFF_DATEFINGENER.Visible := VoirDate;
if  VoirDate then  // gm 08/04/02
Begin
	AFF_DATEDEBGENER.text  := AFF_DATEDEBUT.text;
  AFF_DATEFINGENER.text  := AFF_DATEFIN.text;
End;

end;
procedure TFAssistCreationAffaire.RepFactAffClick(Sender: TObject);
begin  //mcd 16/09/03 car dans ce cas, le mtt forfaitaire n'est pas pris en compte..; donc on ne permet pas de le changer
Aff_MontantEche.enabled := Not(RepFactAff.Checked);
end;

procedure TFAssistCreationAffaire.bAideClick(Sender: TObject);
begin
  inherited;
CallHelpTopic(Self) ;
end;

procedure TFAssistCreationAffaire.AffichePlanning;
begin
  if not VH_GC.GAPlanningSeria then
    begin
      LA_Taches.visible := false;
      CB_DupliquerTache.visible := false;
    end
  else
    begin
      LA_Taches.visible := true;
      CB_DupliquerTache.visible := true;
    end;
end;
    
procedure TFAssistCreationAffaire.AFF_RESPONSABLEElipsisClick(
  Sender: TObject);
begin
  inherited;
DispatchRecherche(AFF_RESPONSABLE,3,'ARS_RESSOURCE='+AFF_RESPONSABLE.text,'','');
end;

Initialization
RegisterAglFunc('AssistCreationAffaire', False ,1, AGLAssistCreationAffaire);
RegisterAglProc('AFSelectionMultiClient',TRUE,0,AGLSelectionMultiClient);
registerclasses([TOF_TIERS_MULTI]);

end.
