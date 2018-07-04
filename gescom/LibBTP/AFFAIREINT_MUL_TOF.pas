{***********UNITE*************************************************
Auteur  ...... : SANTUCCI Lionel
Créé le ...... : 14/02/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFFAIREINT_MUL ()
Mots clefs ... : TOF;AFFAIREINT_MUL
*****************************************************************}
Unit AFFAIREINT_MUL_TOF ;

Interface

uses  StdCtrls,Controls,Classes,M3FP,HTB97,HPanel,windows,messages,Ent1,
{$IFDEF EAGLCLIENT}
      eMul,Maineagl,HQry,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DBGrids,mul,FE_Main,
{$ENDIF}
      forms,sysutils,HDB,ParamSoc,utob,
      ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,ConfidentAffaire,
      AffaireUtil,AffaireRegroupeUtil,Dicobtp,SaisUtil,EntGC,
      utofAfBaseCodeAffaire,utilpgi,AglInit,UtilGc,
      UtofAfAvenant, // à laisser car du scirpt, appel d'une fiche avec cette tof. Pour Ok dans projet
      AssistCreationAffaire, // mcd 10/06/02 appelé du script. a laisser
      TraducAffaire;

Type
  TOF_AFFAIREINT_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (StArgument : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private

    Affaire0		: String;
    Fonction 		: string;
    Statut			: String;
    sPlus       : String;
    EtatAffaire : THMultiValComboBox;

    BOuvrir1   	: TToolBarButton97;

    Top 				: Integer;
    GNbLoads 		: integer;

    VUserInvite : boolean;

    Fliste 			: THDbGrid;

    bPasSiUn		: Boolean;
    bItDatesMax : Boolean;

    procedure AffectationRessource(Sender: TOBJect);
  	procedure BinsertClick (Sender: TObject);
		procedure ControleCritere(Critere, Valeur : string);
    procedure ControleStatut;
    procedure FlisteDblClick (Sender : TObject);

  end ;

Implementation

procedure TOF_AFFAIREINT_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFFAIREINT_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFFAIREINT_MUL.OnUpdate ;
begin
  Inherited ;
{$IFDEF EAGLCLIENT}
if (Ecran is TFMul) then ModifColAffaireGrid ( TFMul(Ecran).FListe,TFMul(Ecran).Q.tq);
{$ELSE}
//if (Ecran is TFMul) then ModifColAffaireGrid ( TFMul(Ecran).FListe);
if (Ecran is TFMul) then ModifColAffaireGrid ( TFMul(Ecran));
{$ENDIF}
end ;

procedure TOF_AFFAIREINT_MUL.OnLoad ;
//Var //RazAffRef : Boolean;
    //Affaire,Affaire0, Affaire1  ,Affaire2 ,Affaire3,Avenant : string;
begin
inherited;
  Inc (GNbLoads);

end ;

procedure TOF_AFFAIREINT_MUL.OnArgument (StArgument : String ) ;
var
  Critere			: string;
  Affaire			: string;
  {sValeursEtatsAffaire : string;}
  i 					: integer;
  champ				: string;
  valeur			: string;
  X 					: integer;
  bPasSiUn		: boolean;
  bItDatesMax : boolean;
  CC 					: THValCOmboBox;
Begin
Inherited;
GNbLoads := 0;
bPasSiUn := false;
bItDatesMax := false;
top :=0; Statut :=''; sPlus := '';
EtatAffaire := THMultiValComboBox (GetControl ('AFF_ETATAFFAIRE'));
Fliste := THDbGrid (GetControl('FLISTE'));
VUserInvite:=False;

//Récupération des critéres de lancement
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
    X:=pos(':',Critere);
    if x<>0 then
       begin
       Champ:=copy(Critere,1,X-1);
       Valeur:=Copy (Critere,X+1,length(Critere)-X);
       end
    else
       begin
       X:=pos('=',Critere);
       if x<>0 then
          begin
          Champ:=copy(Critere,1,X-1);
          Valeur:=Copy (Critere,X+1,length(Critere)-X);
          end
       end;
       ControleCritere(Critere, Valeur);
       Critere:=(Trim(ReadTokenSt(stArgument)));
    END;

//Controle de la zone statut pour affichage
ControleStatut;

if Not(ModifAffaireAutorise) and (GetControlText('XXAction') <> 'ACTION=RECH') then SetControlText('XXAction','ACTION=CONSULTATION');

If GetControl ('XXACTION') <>Nil then
   begin
   if GetControlText('XXAction') <> '' then
      begin
      if (StringToAction(GetControlText('XXAction')) = taconsult) then
         begin
         {$IFDEF BTP}
         SetControlVisible ('BInsert',true); // on autorise la creation d'affaire en recherche
         {$ELSE}
         SetControlVisible ('BInsert',False);
         {$ENDIF}
         SetControlVisible ('BInsert1',False);
         end;
      end;
   end;

{$IFNDEF BTP}
// PL le 08/11/02
// Si on a un champ Etat Affaire et qu'il n'est pas préinitialisé, il faut adapter la sélection du mul
// en mettant toutes les valeurs présentes dans la liste dans la propriété .Text
if (EtatAffaire <> nil)  then
   if (EtatAffaire.text = '') then
      EtatAffaire.Text := StringReplace(EtatAffaire.Values.CommaText, ',', ';', [rfReplaceAll]);
{$ENDIF}

UpdateCaption(Ecran);

// on peut modifier le statut, il faut donc avoir les 2 boutons d'insert
if (top = 0) then
   begin
   SetControlVisible ('BInsert',True);
   end;

// Gestion des sous affaires
{$IFDEF BTP}
//SetcontrolVisible('PCOMPLEMENT',False);
{$ENDIF}

// Test droit d'accès en création
if not(CreationAffaireAutorise) then
   begin
   SetControlVisible('bInsert',False);
   end;

ChargeCleAffaire(THEdit(GetControl('AFF_AFFAIRE0')),THEdit(GetControl('AFF_AFFAIRE1')),THEdit(GetControl('AFF_AFFAIRE2')),THEdit(GetControl('AFF_AFFAIRE3')),
 								 THEdit(GetControl('AFF_AVENANT')),Nil,TaCreat,Affaire,false);

{$IFDEF EAGLCLIENT}
//SetControlVisible ('BASSISTANTCREATION',False); portage en Eagl de l'assistant création affaire
TraduitAFLibGridSt(TFMul(Ecran).FListe);
{$ELSE}
TraduitAFLibGridDB(TFMul(Ecran).FListe);
{$ENDIF}

if ((Ecran.FindComponent ('AFF_LIBREAFF1'))<>nil) then
    begin
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'AFF_LIBREAFF', 10, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'AFF_VALLIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'AFF_DATELIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'AFF_CHARLIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'AFF_BOOLLIBRE', 3, '_');
    end;

// gestion Etablissement (BTP)
CC:=THValComboBox(GetControl('AFF_ETABLISSEMENT')) ; if CC<>Nil then PositionneEtabUser(CC) ;
Fliste.OnDblClick := FlisteDblClick;
TToolbarButton97 (GetControl('Binsert')).OnClick := BInsertClick;

end ;

//controle du statut de l'affaire
procedure TOF_AFFAIREINT_MUL.ControleStatut;
Begin

    if statut = 'PRO' then
       begin
       SetControlVisible ('BInsert',False);
       sPlus := sPlus + GetPlusEtatAffaire (False);
       SetControlProperty('AFF_STATUTAFFAIRE','Value','PRO');
       SetControlText ('TAFF_AFFAIRE', TraduitGa ('Appel d''offre') ) ;
       Ecran.Caption := TraduitGA('Etudes Appel d''offres');
       SetControlProperty('BInsert','Hint', TraduitGA('Nouvel appel d''offre')) ;
       end
    else if statut = 'INT' then
       begin
       SetControlVisible ('BInsert',False);
       sPlus := sPlus + GetPlusEtatAffaire (False);
       Affaire0 := 'I';
       SetControlProperty('AFF_STATUTAFFAIRE','Value','INT');
       SetControlText ('TAFF_AFFAIRE', TraduitGa ('Contrat') ) ;
       Ecran.Caption := TraduitGA('Interventions - Contrats');
       SetControlProperty('BInsert','Hint', TraduitGA('Nouveau contrat'));
       end
		else if statut = 'APP' then
       begin
       sPlus := sPlus + ' AND (ISNUMERIC(CC_LIBRE)=1)';
       Affaire0 := 'W';
       SetControlText ('TAFF_AFFAIRE', TraduitGa ('Appel') ) ;
       SetControlProperty('AFF_STATUTAFFAIRE','Value',Statut);
       SetControlProperty('AFF_ETATAFFAIRE','Value','ECO');
       SetControlProperty('AFF_PRIOCONTRAT','Value','');
       SetControlProperty ('AFF_ETATAFFAIRE','Plus', sPlus);
       Ecran.Caption := TraduitGA('Interventions - Appels');
       SetControlProperty('BInsert','Hint', TraduitGA('Nouvel Appel')) ;
       end
    else if  statut = 'AFF' then
       Begin
       sPlus := sPlus + GetPlusEtatAffaire (True);
       Affaire0 := 'A';
       Statut := 'AFF';
//uniquement en line
{*       SetControlText ('TAFF_AFFAIRE', TraduitGa ('Affaire') ) ;
       SetControlProperty ('AFF_ETATAFFAIRE','Plus', sPlus);
       Ecran.Caption := TraduitGA('Affaires');
       SetControlProperty('BInsert','Hint', TraduitGA('Nouvelle Affaire'));
*}
       SetControlText ('TAFF_AFFAIRE', TraduitGa ('Chantier') ) ;
       SetControlProperty ('AFF_ETATAFFAIRE','Plus', sPlus);
       Ecran.Caption := TraduitGA('Chantiers');
       SetControlProperty('BInsert','Hint', TraduitGA('Nouveau Chantier'));
       end
    else
       begin
       sPlus := sPlus + GetPlusEtatAffaire (True);
       Affaire0 := 'A';
       Statut := 'AFF';
       SetControlText ('TAFF_AFFAIRE', TraduitGa ('Affaire') ) ;
       SetControlProperty ('AFF_ETATAFFAIRE','Plus', sPlus);
       Ecran.Caption := TraduitGA('Chantiers');
       SetControlProperty('BInsert','Hint', TraduitGA('Nouveau Chantier')) ;
       end;

end;

//controle du critere d'appel de la fiche
procedure TOF_AFFAIREINT_MUL.ControleCritere(Critere, Valeur : string);
Var TobEtatsBloques : TOB;
    i								: integer;
begin

   if (copy(Critere,1,9) ='AFF_TIERS')then
      begin
      SetCOntrolText('XXLEQUEL' ,Critere);
      SetControltext ('AFF_TIERS', valeur);
      exit;
      end;

   if (copy(Critere,1,6) ='ACTION') then
      SetControlText('XXACTION',Critere);

   if (copy(Critere,1,4) ='ETAT') then
      Begin
      SetControlProperty ('AFF_ETATAFFAIRE','value',Valeur);
      if Statut = 'APP' then
         Begin
		     EtatAffaire.text := StringReplace(Valeur, ',', ';', [rfReplaceAll]);
         Exit;
         End
      else
         Begin
	       EtatAffaire.text := Valeur;
         exit;
         end;
      end;

   // Réduction de la multicombo des états de l'affaire en fonction des blocages affaire
   if (Copy(Critere,1,3) = 'CTX')  and (EtatAffaire <> nil) then
      begin
      tobEtatsBloques := nil;
      try
        tobEtatsBloques := TOBEtatsBloquesAffaire (copy (critere, 5, 3), '-', V_PGI.groupe);
        sPlus := EtatAffaire.plus;
        if (tobEtatsBloques <> nil) then
           for i:=0 to tobEtatsBloques.Detail.count - 1 do
               begin
               sPlus := sPlus  + ' AND CC_CODE<>"' + tobEtatsBloques.Detail[i].GetValue('ABA_ETATAFFAIRE') + '"';
               end;
           SetControlProperty('AFF_ETATAFFAIRE','Plus', sPlus);
      finally
        tobEtatsBloques.Free;
      end;
      end;

    if (Copy(Critere,1,6) = 'STATUT') and (EtatAffaire <> nil) then
       begin
       sPlus  := EtatAffaire.plus;
       Statut := copy(critere,8,3);
       exit;
       end;

    if critere = 'APPOFFACCEPT' then
       begin
       SetControlVisible ('BInsert',False);
       SetControlProperty('AFF_ETATAFFAIRE','Plus',GetPlusEtatAffaire(False));
       SetControlProperty('AFF_ETATAFFAIRE','Value','ENC');
{$IFDEF BTP}
       SetControlProperty('AFF_STATUTAFFAIRE','Value','PRO');
{$ENDIF}
       Statut := 'PRO';
       SetControlProperty('Fliste','Multiselection',true);
       exit;
       end;

		if critere = 'NOCHANGESTATUT' then
       BEGIN
       SetControlVisible ('AFF_STATUTAFFAIRE',False);
       SetControlVisible ('TAFF_STATUTAFFAIRE',False);
       top :=1;
       exit;
       END;

    if critere = 'AFFECTATION' then
       begin
       Fonction := critere;
       SetControlVisible ('BOuvrir',False);
       SetControlVisible ('BSelectAll',True);
       SetControlVisible ('BOuvrir1',True);
       SetControlProperty('AFF_ETATAFFAIRE','Value','ECO');
       SetControlenabled('AFF_ETATAFFAIRE', false);
       SetControlproperty('FListe', 'Multiselection', true);
       BOuvrir1 := TToolbarButton97(ecran.FindComponent('BOuvrir1'));
		   BOuvrir1.OnClick := AffectationRessource;
       exit;
       end;

    if critere = 'NOCHANGETIERS' then BEGIN SetControlEnabled('AFF_TIERS',False); END;

    if critere = 'NOFILTRE' then   Tfmul(Ecran).FiltreDisabled:=true;

    if critere = 'PASSIUN' then   bPasSiUn := true;

    if critere = 'REDUIT' then
       if Not(GetParamsoc('SO_AfSaisAffInterdit')) then VUserInvite := True; //mcd 18/10/02 ajout test paramsoc

end;

procedure TOF_AFFAIREINT_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AFFAIREINT_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFFAIREINT_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_AFFAIREINT_MUL.FlisteDblClick(Sender: TObject);
var Affaire,Tiers,StatutAffaire,EtatAffaire,Action : string;
begin

  Tiers:=Fliste.datasource.dataset.FindField('AFF_TIERS').AsString;

  //Affaire:=Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString;
  //StatutAffaire:=Fliste.datasource.dataset.FindField('AFF_STATUTAFFAIRE').AsString;

  //FV1 : 17/11/2015 -  FS#1786 - POUCHAIN : Pb sur liste appels depuis devis
  Affaire:=Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString;
  StatutAffaire := Copy(Affaire, 1,1);

  EtatAffaire:=Fliste.datasource.dataset.FindField('AFF_ETATAFFAIRE').AsString;
  Action:=GetControlText('XXAction');

  if Action='ACTION=RECH' then
	begin
	  TFMul(Ecran).Retour :=Affaire+';'+Tiers;
    TFMUL(Ecran).Close;
	end
  else
	begin
    //FV1 : 17/11/2015 -  FS#1786 - POUCHAIN : Pb sur liste appels depuis devis
    if StatutAffaire = 'W' then
       AGLLanceFiche ('BTP','BTAPPELINT','','','CODEAPPEL:' + Affaire + ';ETAT:'+ EtatAffaire + ';ACTION=MODIFICATION')
    else
  	   AGLLanceFiche ('BTP','BTAFFAIRE','',Affaire,'STATUT:'+StatutAffaire+';ETAT:'+ EtatAffaire +';'+Action);
		TtoolBarButton97(GetCOntrol('Bcherche')).Click;
	end;

end;

procedure TOF_AFFAIREINT_MUL.BinsertClick(Sender: TObject);
var NatPiece,StatutAffaire,EtatAffaire,Tiers : string;
    Statut  : string;
    Affaire : string;
begin

  //FV1 : 17/11/2015 -  FS#1786 - POUCHAIN : Pb sur liste appels depuis devis
  //StatutAffaire := THValComboBox(GetControl('AFF_STATUTAFFAIRE')).Value ;

  EtatAffaire := THValComboBox(GetControl('AFF_ETATAFFAIRE')).Value ;
  Tiers       := 'AFF_TIERS:' + GetControlText('AFF_TIERS');

  //FV1 : 17/11/2015 -  FS#1786 - POUCHAIN : Pb sur liste appels depuis devis
  Affaire     :=THEdit(GetControl('AFF_AFFAIRE')).Text;
  Statut      := Copy(Affaire, 1,1);

  if (StatutAffaire = 'PRO') then
    natpiece := 'ETU'
  else
    natpiece := 'DBT';

  //FV1 : 17/11/2015 -  FS#1786 - POUCHAIN : Pb sur liste appels depuis devis
  if Statut = 'W' then
     AglLanceFiche ('BTP','BTAPPELINT','','','ACTION=CREATION')
  else
     AGLLanceFiche ('BTP','BTAFFAIRE','','','ACTION=CREATION;STATUT:'+StatutAffaire+';' + Tiers + ';'+ EtatAffaire + ';NATURE:'+natpiece);

TtoolBarButton97(GetCOntrol('Bcherche')).Click;
end;

Procedure TOF_AFFAIREINT_MUL.AffectationRessource(Sender : TOBJect);
var Arg,Affaire      : String;
    TOBparam : Tob;
    TOB1Appel: Tob;
    TOBDAppel: Tob;
    F 			 : TFMul;
    i 			 : integer;
    QQ			 : TQuery ;
{$IFDEF EAGLCLIENT}
    L 			 : THGrid;
{$ELSE}
    Ext 		 : String;
    L 			 : THDBGrid;
{$ENDIF}

begin  Inherited ;

	F:=TFMul(Ecran);

	if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
  	 begin
	   PGIInfo('Aucun élément sélectionné','');
  	 exit;
   	 end;

  TOBParam := TOB.Create ('AFFECT ',nil,-1);
  TOBParam.AddChampSupValeur ('RESSOURCE','');
	TOBParam.AddChampSupValeur ('RETOUR','-');

  TheTOB := TOBparam;
  AGLLanceFiche ('BTP','BTAFFECTATION','','','');
  TheTOB := nil;

  if TOBparam.GetValue('RETOUR')='-' then
		 Begin
  	 TobParam.free;
     exit ;
     end;

	if PGIAsk('Confirmez-vous le traitement ?','') <> mrYes then
     Begin
     TobParam.free;
  	 exit ;
     end;

  TOBDAppel := tob.create('LES APPELS', nil, -1);

	L:= F.FListe;
  SourisSablier;

  try
	if L.AllSelected then
  	 begin
  	 QQ:=F.Q;
  	 QQ.First;
	   while Not QQ.EOF do
     Begin
     	 Affaire := QQ.FindField('AFF_AFFAIRE').AsString;
       TOB1Appel := TOB.Create ('AFFAIRE',TOBDappel,-1);
       TOB1Appel.putValue('AFF_AFFAIRE',Affaire);
       if TOB1Appel.LoadDb (false) then
       begin
       	 TOB1Appel.PutValue('AFF_RESPONSABLE',TOBParam.getValue ('RESSOURCE'));
         TOB1Appel.putValue('AFF_ETATAFFAIRE','AFF');
       end else TOB1Appel.free;
     	 QQ.Next;
   	 end;
	   L.AllSelected:=False;
  	 end
  else
     begin
     for i:=0 to L.NbSelected-1 do
         begin
	       L.GotoLeBookmark(i);
         Affaire := TFMul(F).Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString;
         TOB1Appel := TOB.Create ('AFFAIRE',TOBDappel,-1);
         TOB1Appel.putValue('AFF_AFFAIRE',Affaire);
         if TOB1Appel.LoadDb (false) then
         	 begin
           TOB1Appel.PutValue('AFF_RESPONSABLE',TOBParam.getValue ('RESSOURCE'));
         	 TOB1Appel.putValue('AFF_ETATAFFAIRE','AFF');
         	 end
         else TOB1Appel.free;
      	 end;
	   		 L.ClearSelected;
	   end;
  if TOBDappel.detail.count > 0 then TOBDappel.UpdateDB (true);
	finally
  TOBDappel.free;
  TobParam.free;
	SourisNormale;
  F.BChercheClick(ecran);
  end;

end;

Initialization
  registerclasses ( [ TOF_AFFAIREINT_MUL ] ) ;
end.

