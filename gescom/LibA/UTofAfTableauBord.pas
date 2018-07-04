unit UTofAfTableauBord;                              
interface                                                           
 
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,UTOF,AFTableauBord,FactUtil,AffaireUtil,Dicobtp,HEnt1,
      EntGC,HTB97,utofbaseetats ,ParamSoc,
      UTofAfBaseCodeAffaire, UTOB, Stat,M3FP,utobview,ActiviteUtil,AFActivite,
{$IFDEF EAGLCLIENT}
      MaineAGL, eQRS1,
{$else}
      QRS1,Fe_Main, db,  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF BTP}
	  CalcOleGenericBTP,
{$ENDIF}
      Windows,UtofAfModifCutOff_Mul, AffaireRegroupeUtil,UtilGa,UtofAfsaisDAte ,Menus ,Ent1 ,AglInitAff
      ,UtofAfPlanning_mul, UtofAfbaseLigne_Mul ,UtofAfActiviteMul, UtilGc
{$IFDEF GIGI}
      ,UtofAppreciation, UtofAppreCon_Mul ,UtofAfModifBoni_Mul
{$ENDIF}
      ;

Type
     T_NiveauVoir     = (VRien,VTB,VActivite,VAffaire,VFacture,VBoni,VCutOff,VEngage,VPla,VBudg,VArt,Vapp) ;

Type
   TOF_TBAffaire = Class (TOF_BASE_ETATS)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnUpdate ; override ;
        procedure OnLoad ; override ;
        procedure ImpactRegroupementEdit (Sender: TObject);
        procedure NomsChampsAffaire(var Aff, Aff0,Aff1, Aff2, Aff3,Aff4, Aff_, Aff0_,Aff1_, Aff2_, Aff3_,Aff4_, Tiers, Tiers_:THEdit);override ;
        procedure NomsChampsAffaireString(var sAff0, sAff1, sAff2, sAff3, sAff4 : string); override;
        procedure TableauObjectsInvisibles(Crit:string; var iNbChamps:integer; var tbChamps:PString);override;
    Private
        procedure ReloadComboEtat;

     END ;

     TOF_TBAffViewer = Class (TOF_AFBASECODEAFFAIRE)
        procedure OnArgument (stArgument : string); override;
        procedure OnLoad; override;
        procedure OnClose; override;
        procedure ImpactRegroupementView (Sender: TObject);
        procedure NomsChampsAffaire(var Aff, Aff0,Aff1, Aff2, Aff3,Aff4, Aff_, Aff0_,Aff1_, Aff2_, Aff3_,Aff4_, Tiers, Tiers_:THEdit);override ;
        procedure RecupFormuleEcart(valeur ,zone : string);  // Gestion des écarts
        procedure AlimEcartTB;
        Procedure MenuPopOnClick(Sender:TobJect);
        private
        TV:TTobViewer;
        Voir : T_NiveauVoir;
        ParamTB : TParamTBAffaire;
        TypeAppelTB : String;
        Zbut : String;
        TOBTB : TOB;
        procedure ImpactTBTiers (Sender: TObject);
        procedure ImpactEcart   (Sender: TObject);
        procedure ImpactTotalChamps (Sender: TObject);
        procedure ImpactAchat   (Sender: TObject);
        procedure LanceAffaireTB;
        procedure OnDblClickTV (Sender: TObject);
        Procedure RecupAffaireLigneTB (Var Affaire, Tiers : string);
        procedure AppelAppreciation(zaff : string);
     END;

   TOF_AFTBMENU = Class (TOF)
        procedure OnArgument(stArgument : String ) ; override ;
    end;

Procedure AFLanceFiche_Balance(Argument:String);
Function  AfBuildAfTableauBord (Zone : T_PARAMTB):boolean;
Procedure AFLanceFiche_GrandLivre(Argument:String);
Procedure AFLanceFiche_TableauBord;
Procedure AFLanceFiche_TbAffaire(Argument:String);
Procedure AFLanceFiche_TBMultiAff(Argument:String);
Procedure AFLanceFiche_EtatAppreciation(Argument:String);
Procedure AFLanceFiche_EtatContrProd(Argument:String);
Function GetValoActDefaut (LesDeuxPossible : Boolean) : string;
//Procedure ComboPlusAchat (Var Pluscombo: string) ;

const
NbTbChInv = 2 ;
TbChampsInvisibles : array[1..NbTbChInv] of string 	= (
          {1}        'ATB_TYPERESSOURCE',
          {2}       'TATB_TYPERESSOURCE'
          );

implementation
Var         bMulti : Boolean;


 procedure TOF_TBAffaire.ReloadComboEtat;
Var Q : TQuery ;
    i,idef : integer;
    Dernier,StSQL : string;
begin
TFQRS1(ECRAN).FEtat.Items.Clear ; TFQRS1(ECRAN).FEtat.Values.Clear ; idef:=-1;
StSQL:='SELECT MO_TYPE, MO_NATURE, MO_CODE, MO_LIBELLE, MO_LANGUE, MO_DEFAUT FROM MODELES WHERE MO_TYPE="'+TFQRS1(ECRAN).FTypeEtat+'" AND MO_NATURE="'+TFQRS1(ECRAN).FNatEtat+'"' ;
if Not TFQRS1(ECRAN).FChoixEtat then StSQL:=StSQL+' AND MO_CODE="'+TFQRS1(ECRAN).FCodeEtat+'"' ;
Q:=OpenSQL(StSQL,TRUE) ;
While Not Q.EOF do
  BEGIN
  i:=TFQRS1(ECRAN).FEtat.Items.Add(Q.FindField('MO_LIBELLE').AsString) ;
  if Q.FindField('MO_DEFAUT').AsString='X' then Idef:=i ;
  TFQRS1(ECRAN).FEtat.Values.Add(Q.FindField('MO_CODE').AsString) ;
  Q.Next
  END ;
Ferme(Q) ;
TFQRS1(ECRAN).FEtat.ItemIndex:=-1;
if TFQRS1(Ecran).FEtat.Value='' then
  begin
  Dernier:=GetFromRegistry(HKEY_LOCAL_MACHINE,'Software\'+Apalatys+'\'+NomHalley+'\QRS1Def',GetLeNom(TFQRS1(Ecran).FNomFiltre),'') ;
  if Dernier<>'' then TFQRS1(ECRAN).FEtat.ItemIndex:= TFQRS1(ECRAN).FEtat.values.IndexOf(Dernier);
  if TFQRS1(ECRAN).FEtat.ItemIndex=-1 then TFQRS1(ECRAN).FEtat.ItemIndex:=idef;
  end else TFQRS1(ECRAN).FEtat.ItemIndex:= TFQRS1(ECRAN).FEtat.values.IndexOf(TFQRS1(Ecran).FEtat.Value);
TFQRS1(ECRAN).FEtatClick(Nil) ;
end;

//******************************************************************************
//**************** TOB fiche de lancement d'édition   **************************
//******************************************************************************
procedure TOF_TBAffaire.OnArgument(stArgument : String );
Var Combo   : THValComboBox;
Zone   : THMultiValComboBox;
    Critere, Champ, valeur  : String;
    x : integer;
    QEX : TQuery;
    TypeAppelTb, Part1, PArt0, Part2, Part3, Avenant : string;
Begin
Inherited;
TypeAppelTb :='';    //mcd ajout du traitement 12/02/03
          // on récupère les paramètres éventuels
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN
        X:=pos(':',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;
        if Champ = 'ETAT' then begin
           TFQRS1(Ecran).CodeEtat :=valeur;
           TFQRS1(Ecran).NatureEtat :=valeur;
           end
        else if Champ = 'TYPETB'      then TypeAppelTB := Valeur
        else if Champ = 'AFFAIRE' then
          begin
          SetControlText('ATBXAFFAIRE',Valeur); //mcd 12/02/03
                //mcd 10/06/02 ajout champs enabled si passé en paramètre
          SetControlEnabled ('ATBXAFFAIRE1',False);
          SetControlEnabled ('ATBXAFFAIRE2',False);
          SetControlEnabled ('ATBXAFFAIRE3',False);
          SetControlEnabled ('ATBXAVENANT',False);
          SetControlEnabled ('ATBXTIERS',False);
          setControlEnabled ('BSELECTAFF1',False);
          setControlEnabled ('BEFFACEAFF1',False);
          SetControlEnabled ('ATBXAFFAIRE1_',False);
          SetControlEnabled ('ATBXAFFAIRE2_',False);
          SetControlEnabled ('ATBXAFFAIRE3_',False);
          SetControlEnabled ('ATBXAVENANT_',False);
          SetControlEnabled ('ATBXTIERS_',False);
          setControlEnabled ('BSELECTAFF2',False);
          setControlEnabled ('BEFFACEAFF2',False);
          BchangeTiers :=False;
          end
        else if Champ = 'TIERS' then
         begin
         SetControlText('ATBXTIERS',Valeur);
         SetControlEnabled ('ATBXTIERS',False);
         SetControlText('ATBXTIERS_',Valeur);
         SetControlEnabled ('ATBXTIERS_',False);
         SetControlEnabled ('ATBXLIBELLETIERS_',False);
         SetControlEnabled ('ATBXLIBELLETIERS',False);
         BchangeTiers :=False;
         end      //mcd 12/02/03
        else if champ = 'NOFILTRE' then  TFQRS1(Ecran).FiltreDisabled:=true;  //mcd 12/02/03
        END;
    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;
// Suppression des champs mois de cloture en Gestion d'Affaire
if Not (ctxScot in V_PGI.PGIContexte) then
   begin
   SetControlVisible('ATBXMOISCLOTURE',False); SetControlVisible('ATBXMOISCLOTURE_',False);
   SetControlVisible('TATB_MOISCLOTURE',False); SetControlVisible('TATB_MOISCLOTURE_',False);
   if (THEdit(GetControl('ATBXAFFDATEDEBEXER'))<>nil) then SetControlVisible('ATBXAFFDATEDEBEXER',False);
   if (THEdit(GetControl('ATBXAFFDATEDEBEXER_'))<>nil) then SetControlVisible('ATBXAFFDATEDEBEXER_',False);
   if (THLabel(GetControl('TAFF_DATEDEBEXER'))<>nil) then SetControlVisible('TAFF_DATEDEBEXER',False);
   if (THLabel(GetControl('TAFF_DATEDEBEXER_'))<>nil) then SetControlVisible('TAFF_DATEDEBEXER_',False);
   if (THEdit(GetControl('ATBXAFFDATEFINEXER'))<>nil) then SetControlVisible('ATBXAFFDATEFINEXER',False);
   if (THEdit(GetControl('ATBXAFFDATEFINEXER_'))<>nil) then SetControlVisible('ATBXAFFDATEFINEXER_',False);
   if (THLabel(GetControl('TAFF_DATEFINEXER'))<>nil) then SetControlVisible('TAFF_DATEFINEXER',False);
   if (THLabel(GetControl('TAFF_DATEFINEXER_'))<>nil) then SetControlVisible('TAFF_DATEFINEXER_',False);
   end
else begin
   if (THEdit(GetControl('ATBXAFFDATELIMITE'))<>nil) then SetControlVisible('ATBXAFFDATELIMITE',False);
   if (THEdit(GetControl('ATBXAFFDATELIMITE_'))<>nil) then SetControlVisible('ATBXAFFDATELIMITE_',False);
   if (THLabel(GetControl('TAFF_DATELIMITE'))<>nil) then SetControlVisible('TAFF_DATELIMITE',False);
   if (THLabel(GetControl('TAFF_DATELIMITE_'))<>nil) then SetControlVisible('TAFF_DATELIMITE_',False);
  end;
// appel automatique du TB depuis l'affaire
if (TypeAppelTB = 'AFF') and (GetControlText('ATBXAFFAIRE') <> '') then
   BEGIN
   {$IFDEF BTP}
   BTPCodeAffaireDecoupe(GetControlText('ATBXAFFAIRE'),Part0,Part1,Part2,Part3,Avenant,taModif,false);   
   {$ELSE}
   CodeAffaireDecoupe(GetControlText('ATBXAFFAIRE'),Part0,Part1,Part2,Part3,Avenant,taModif,false);
   {$ENDIF}
   SetControlText('ATBXAFFAIRE1',Part1); SetControlText('ATBXAFFAIRE2',Part2);
   SetControlText('ATBXAFFAIRE3',Part3); SetControlText('ATBXAVENANT',avenant);
   SetControlText('ATBXAFFAIRE1_',GetControlText('ATBXAFFAIRE1'));
   SetControlText('ATBXAFFAIRE2_',GetControlText('ATBXAFFAIRE2'));
   SetControlText('ATBXAFFAIRE3_',GetControlText('ATBXAFFAIRE3'));
   SetControlText('ATBXAVENANT_',GetControlText('ATBXAVENANT'));
   END;

Zone :=ThMultiValCOmboBox(GetControl('NATUREPIECEGACHAT'));
If zone <>Nil then Zone.plus := Zone.plus + AfPlusNatureAchat(true);
Zone :=ThMultiValCOmboBox(GetControl('NATUREPIECEGACHATENG'));
If zone <>Nil then Zone.plus := Zone.plus + AfPlusNatureAchat(true);
if VH_GC.GAAchatSeria =False then
   begin
   SetControlVisible('NATUREPIECEGACHAT',False);
   SetControlVisible('TNATUREPIECEGACHAT',False);
   SetControlVisible('NATUREPIECEGACHATENG',False);
   SetControlVisible('TNATUREPIECEGACHATENG',False);
   end
else if CtxScot in V_PGI.PGIContexte then begin
     SetControlVisible('NATUREPIECEGACHATENG',False);
     SetControlVisible('TNATUREPIECEGACHATENG',False);
     Zone:=ThMultiValCOmboBox(Getcontrol('NATUREPIECEGACHAT'));
     if zone <>Nil then Zone.text:='FF;';     //pas de cde pour SCOT
     end;
Combo := THValComboBox(GetControl('REGROUPEMENT'));
Combo.OnClick := ImpactRegroupementEdit;
SetControlText ('XX_VARIABLE3', VH_GC.AFMesureActivite);
SetControlText ('ATB_UTILISATEUR',V_PGI.User);
if (VH_GC.AFDATEANALYSETB <>iDate1900) then
   begin
   SetControlText('ATBXDATE',DateTostr(VH_GC.AFDATEANALYSETB) );
   SetControltext('ATBXDATE_',DatetoStr(FinDeMois(PlusDate(VH_GC.AFDATEANALYSETB,11,'M'))));
   end;
if (ctxscot in V_PGI.PGIcontexte) then
   begin
   SetControlText('ATBXDATE',DateTostr(GetParamSoc('SO_AFDATEDEBCAB') ));
   SetControltext('ATBXDATE_',DatetoStr(GetParamSoc('SO_AFDATEFINCAB')));
   end;
// appel automatique du TB depuis l'affaire, ilf aut mettre les dates les plus larges
if (TypeAppelTB = 'AFF')  then
   BEGIN
   SetControlText('ATBXDATE',DatetoSTr(Idate1900 ));
   SetControltext('ATBXDATE_',DateToSTr(idate2099));
   END;
SetControlVisible('AVECSSAFF',GereSousAffaire);
if GetParamSoc('so_afVisaActivite') = False then begin
   SetControlVisible('alimvisa',False);
   SetControlChecked ('AlimVisa',False);
   end;
       // mcd 25/03/02
if GetParamSoc('so_afAppAvecBM') = False then begin
   if TCheckBox(GetControl('ALIMBONI')) <>NIl then begin
    SetControlVisible('alimBoni',False);
    SetControlChecked ('AlimBoni',False);
    end;
  end;
if GetParamSoc('so_afAppCutOff') = False then
   begin
   if TCheckBox(GetControl('ALIMCutOFF')) <>NIl then begin
      SetControlVisible('alimCutoff',False);
      SetControlChecked ('AlimCutoff',False);
      end;
   if TCheckBox(GetControl('CutOffMois')) <>NIl then begin  // mcd 03/06/02
      SetControlVisible('CutoffMois',False);
      SetControlChecked ('CutoffMois',False);
      end;
  end;
SetControlText('ATBXSTATUTAFFAIRE','AFF' );   //mcd 23/04/02 pour n'avoir que les affaire en std

// MCD 22/07/02 mis ici au lieu de le mettre dans la boucle sur stargument
// permet en fct de la nature état passé, de cacher ou non des zones des fiches de lancement identiques à plusieurs états
if (TFQRS1(Ecran).NatureEtat = 'ACU') then begin
          // cas cut off
  Ecran.caption :='Impressions Cut Off';
  UpdateCaption(Ecran);
  SetControlVisible ('ALIMBONI',False);
  SetControlChecked('ALIMBONI',false);
  SetControlChecked('ALIMPREVU',false);   //mcd 07/01/03 pas de prévu dans etat cut off.inutile de le traiter
  SetControlChecked('ALIMCUTOFF',true);
  SetControlVisible('CUTOFFMois',False);
  SetControlChecked('CUTOFFMois',False);
  end
Else if (TFQRS1(Ecran).NatureEtat = 'APF') then begin
          // Appréciation
  SetControlVisible ('REGROUPEMENT',False);
  SetControlVisible ('TREGROUPEMENT',False);
  SetControlText('REGROUPEMENT','CUM');
  SetControlVisible ('DETAIL',True);    //mcd 25/02/03
    //mcd 20/08/03
  if GetParamSoc ('SO_AFAPPANOUVEAU') then
    begin
    SetControlVisible('TATB_DATE',False);
    SetControlVisible ('ATBXDATE',False);
    SetControlText('ATBXDATE_',DateToStr (FInDeMois (PlusDate (V_PGI.DateEntree, (-1), 'M'))));
    SetControlText('TATB_DATE_','Fin au');
    end;
  end
Else if (TFQRS1(Ecran).NatureEtat = 'AFP') then begin
          // Etat préparatoire de facturation
  Ecran.caption :='Etat Préparatoire de Factures';
  UpdateCaption(Ecran);
  SetControlProperty('REGROUPEMENT','Plus','CUM" OR CO_LIBRE="DET') ;
  SetControlText('REGROUPEMENT','MOI');
  SetControlVisible ('DATEFAC',True);
  SetControlVisible ('TDATEFAC',True);
  SetControlVisible ('DETAIL',True);
  SetControlVisible ('PAGEAFF',True);
  SetControlVisible ('ALIMMOIS',False);
  SetControlChecked('ALIMMOIS',FALSE);
  SetControlVisible ('FLISTE',False);   // pas de liste d'exportation
  SetControlVisible('CUTOFFMois',False);
  SetControlChecked('CUTOFFMois',False);
  SetControlVisible('ALIMCUTOFF',False);
  if GetParamSoc('SO_AFFORMATEXER')<>'AUC' then SetControlVisible ('CUMEXER',True);
  if GetParamSoc('so_afAppCutOff') = true then SetControlChecked ('AlimCutoff',True)
   else SetControlChecked('ALIMCUTOFF',False);
  if GetParamSoc('so_afAppAvecBM') = true then SetControlChecked ('AlimBoni',True)
   else SetControlChecked('ALIMBONI',False);
  end
Else if (TFQRS1(Ecran).NatureEtat = 'AFR') then begin
          // cas Etat spécif ALGOE DLFY  Cut off par mois ...
  Ecran.caption :='Etat des CA réalisés dont Cut Off';
  UpdateCaption(Ecran);
  if Not (ctxScot in V_PGI.PGIContexte) then
      begin
      SetControlVisible ('TDATEDEBEX',True);
      SetControlVisible ('DATEDEBEX',true);
      QEX:=OpenSql('Select max(ex_datedebut) as wdex from exercice where ex_etatcpta = "OUV"',True);
      if not QEX.EOF then SetControlText('DateDebEx',DateToStr(QEX.Findfield('wdex').asDateTime))
                     else SetControlText('DateDebEx',DateTostr(GetParamSoc('SO_AFDATEDEBCAB')));
      Ferme(QEX);
      end
  else SetControlText('DATEDEBEX',DateTostr(GetParamSoc('SO_AFDATEDEBCAB')));
  SetControlEnabled ('FLISTE',True);
  SetControlVisible ('XX_VARIABLE1',False);
  SetControlVisible ('TXX_VARIABLE1',False);
  SetControlVisible ('XX_VARIABLE3',False);
  SetControlVisible ('TXX_VARIABLE3',False);
  SetControlEnabled ('XX_VARIABLE7',False);
  SetControlVisible ('ALIMVISA',False);
  SetControlVisible ('ALIMAN',False);
  SetControlChecked('ALIMAN',TRUE);
  SetControlVisible ('ALIMMOIS',False);
  SetControlChecked('ALIMMOIS',False);
  SetControlVisible ('ALIMBONI',False);
  SetControlChecked('ALIMBONI',False);
  SetControlVisible('ALIMREAL',false);
  SetControlChecked('ALIMREAL',false);
  SetControlVisible('ALIMPREVU',False);
  SetControlChecked('ALIMPREVU',False);
  SetControlVisible('ALIMCUTOFF',False);
  SetControlChecked('ALIMCUTOFF',TRUE);
  SetControlVisible('CUTOFFMOIS',False);
  SetControlChecked('CUTOFFMOIS',TRUE);
  SetControlEnabled ('REGROUPEMENT',False);
  SetControlText ('REGROUPEMENT','MOI');
  // mcd 13/09/02 SetControlEnabled ('NIVEAUREGROUPETB',False);
  SetControlText ('NIVEAUREGROUPETB','3');
//  SetControlVisible ('TREGROUPEMENT',False);
  end
else if TFQRS1(Ecran).NatureEtat='AAN' then begin
       // cat etats activité avec AN
  Ecran.caption :='Etats avec prise en compte AN';
  UpdateCaption(Ecran);
  SetControlVisible ('ALIMAN',True);
  SetControlChecked('ALIMAN',TRUE);
  end;
If GetParamSoc('SO_AFPROPOSACT')=False then
   begin // mcd 26/07/02 cache statut affaire si porposition intredite en activite
   SetControlVisible ('ATBXSTATUTAFFAIRE',False);
   SetControlVisible ('TATB_STATUTAFFAIRE',False);
   end;

{$IFDEF EAGLCLIENT}     // mcd 02/04/03 l'impression en eagl sur un état fait à partir de TOB ne marche pas...
  setControlVisible('BIMPRIMER', false);
{$ENDIF}
ReloadComboEtat;                 // mcd 13/01/03 pour OK prise en compte nature
SetControlVisible ('BPROG',False);  // passe par table de travil, planif interdite
End;

procedure TOF_TBAffaire.NomsChampsAffaireString(var sAff0, sAff1, sAff2, sAff3, sAff4 : string);
begin
sAff0:= '';
sAff1:='ATBXAFFAIRE1';
sAff2:='ATBXAFFAIRE2';
sAff3:='ATBXAFFAIRE3';
sAff4:='';
end;

procedure TOF_TBAffaire.NomsChampsAffaire(var Aff, Aff0,Aff1, Aff2, Aff3,Aff4, Aff_, Aff0_,Aff1_, Aff2_, Aff3_,Aff4_, Tiers, Tiers_:THEdit);
begin
Aff1:=THEdit(GetControl('ATBXAFFAIRE1'));
Aff2:=THEdit(GetControl('ATBXAFFAIRE2'));
Aff3:=THEdit(GetControl('ATBXAFFAIRE3'));
Aff4:=THEdit(GetControl('ATBXAVENANT'));

Aff1_:=THEdit(GetControl('ATBXAFFAIRE1_'));
Aff2_:=THEdit(GetControl('ATBXAFFAIRE2_'));
Aff3_:=THEdit(GetControl('ATBXAFFAIRE3_'));
Aff4_:=THEdit(GetControl('ATBXAVENANT_'));
Tiers:=THEdit(GetControl('ATBXTIERS'));
Tiers_:=THEdit(GetControl('ATBXTIERS_'));
end;

procedure TOF_TBAffaire.OnLoad;
Var Combo   : THValComboBox;
begin
inherited;
// mcd 03/06/02 mis dans la  fct Onload plutôt que OnUpdate, car OnUpdate fait après le recupWhereSql
// ne fonctionnait que lors de la 2eme sélection identique
Combo := THValComboBox(GetControl('NIVEAUREGROUPETB'));
If Ecran.name='AFAPPREC_ETAT' then
   begin  //mcd 24/04/02 mis dans la tof de ce qui était fait dans le scirpt sinon filtre déjà enregistré pas OK
    SetControlText('XX_RUPTURE10','');
    If (Combo.value = '3') then SetControlText('XX_RUPTURE11','ATB_TYPEARTICLE')
    else If (Combo.value = '4') then SetControlText('XX_RUPTURE11','ATB_CODEARTICLE')
    else If (Combo.value = '6') then begin
          SetControlText('XX_RUPTURE10','ATB_CODEARTICLE');
          SetControlText('XX_RUPTURE11','ATB_RESSOURCE');
          end
    else If (Combo.value = '5') then begin
          SetControlText('XX_RUPTURE10','ATB_TYPEARTICLE');
          SetControlText('XX_RUPTURE11','ATB_RESSOURCE');
          end
    else If (Combo.value = '7') then SetControlText('XX_RUPTURE11','ATB_RESSOURCE')
    else If (Combo.value = '8') then SetControlText('XX_RUPTURE11','ATB_TYPERESSOURCE')
    else If (Combo.value = '10') then SetControlText('XX_RUPTURE11','ATB_FAMILLENIV1')
    else If (Combo.value = '11') then SetControlText('XX_RUPTURE11','ATB_FAMILLENIV2')
    else If (Combo.value = '12') then SetControlText('XX_RUPTURE11','ATB_FAMILLENIV3')
    else If (Combo.value = '21') then SetControlText('XX_RUPTURE11','ATB_LIBREART1')
    else If (Combo.value = '22') then SetControlText('XX_RUPTURE11','ATB_LIBREART2')
    else If (Combo.value = '23') then SetControlText('XX_RUPTURE11','ATB_LIBREART3')
    else SetControlText('XX_RUPTURE11','ATB_DATE,ATB_CODEARTICLE');
    If TfQrs1(ecran).natureEtat ='AFP' then
      begin
        //mcd 23/05/03 dans le cas particulier état prep facture, les rupture 10 et 11 ne sont pas utiliser,
        // par contre le tri par date doit être prioritaire
        // manip faite dans le source, à modifier quand une fiche spécif pour cet état sera faite
      SetControlText('XX_ORDERBY',GetControlText ('XX_RUPTURE11'));
      SetControlText('XX_RUPTURE11',GetControlText ('XX_RUPTURE10'));
      SetControlText('XX_RUPTURE10','ATB_PERIODE,ATB_SEMAINE');
      end;
    end;
end;

procedure TOF_TBAffaire.OnUpdate;
Var Combo,COmbo1    : THValComboBox;
    CCheck   : TCheckBox;
    stChamps,stSQL,NatPieceAchat,NatPieceAchatEngage : string;
    NiveauTB : T_NiveauTB;
    TypeDateTB: T_TypeDateTB;
    MultiCombo : THMultiValComboBox;
    AvecSsAff : Boolean;
    ParamTB : TParamTBAffaire;
Begin
//AvecSsAff := False;
// Traitement du niveau de regroupement (Cumul, par mois, détail ...)
Combo := THValComboBox(GetControl('REGROUPEMENT'));
If (Ecran.name)='AFBALANCE' then begin
     // mcd 19/07/02 pour paser en cumul si pas de récap demandée
   if (TCheckBox(GetControl('CENTRFAM')).checked = False) and
     (TCheckBox(GetControl('CENTRSTATART')).checked = False) and
     (TCheckBox(GetControl('CENTRSTATRES')).checked = False) and
     (TCheckBox(GetControl('RESSOURCE')).checked = False) and
     (TCheckBox(GetControl('ARTICLE')).checked = False) and
     (TCheckBox(GetControl('TYPEART')).checked = False) and
     (TCheckBox(GetControl('RESSART')).checked = False) then
      begin
      Combo.value:='CUM';
      SetControlText ('NIVEAUREGROUPETB','3'); // pour cumul par type article seulement (certain état le veulent)
      end
     else   SetControlText ('NIVEAUREGROUPETB','6');  // pour cumul par art et ressource ==> récap demandée
   end;
TypeDateTB := GetRegroupeTB  (Combo.value);
if (Combo.value = 'DET') then SetControlText('XX_VARIABLE4','ATB_DATE')    else
if (Combo.value = 'SEM') then SetControlText('XX_VARIABLE4','ATB_SEMAINE') else
if (Combo.value = 'MOI') then SetControlText('XX_VARIABLE4','ATB_PERIODE') else
if (Combo.value = 'CUM') then SetControlText('XX_VARIABLE4','');

  //mcd 01/07/03 pour ne pas imprimer AUUCNE si pas de choix de rupture
Combo := THValComboBox(GetControl('XX_RUPTURE1'));
if Uppercase(combo.Text) ='AUCUNE' 
   then  begin
   Combo.style:=CsDropDown; // mise en saisie possible pour changer le texte
   Combo.text :='' ;
   end;
Combo := THValComboBox(GetControl('XX_RUPTURE2'));
if Uppercase(combo.Text) ='AUCUNE'
   then  begin
   Combo.style:=CsDropDown; // mise en saisie possible pour changer le texte
   Combo.text :='' ;
   end;

// Traitement du type de regroupement (article, ressource ...)
Combo := THValComboBox(GetControl('NIVEAUREGROUPETB'));
NiveauTB :=GetNiveauTB(Combo.value);
if (Combo.value = '1') then SetControlText('XX_VARIABLE5','ATB_TIERS')       else
if (Combo.value = '2') then SetControlText('XX_VARIABLE5','ATB_AFFAIRE')     else
if (Combo.value = '3') then SetControlText('XX_VARIABLE5','ATB_TYPEARTICLE') else
if (Combo.value = '4') then SetControlText('XX_VARIABLE5','ATB_CODEARTICLE') else
if (Combo.value = '5') then SetControlText('XX_VARIABLE5','ATB_TYPEARTICLE') else
if (Combo.value = '6') then SetControlText('XX_VARIABLE5','ATB_CODEARTICLE') else
if (Combo.value = '7') then SetControlText('XX_VARIABLE5','ATB_RESSOURCE')   else
if (Combo.value = '8') then SetControlText('XX_VARIABLE5','ATB_TYPERESSOURCE')else
if (Combo.value = '9') then SetControlText('XX_VARIABLE5','ATB_FONCTION') else
if (Combo.value = '10')then SetControlText('XX_VARIABLE5','ATB_FAMILLENIV1') else
if (Combo.value = '11')then SetControlText('XX_VARIABLE5','ATB_FAMILLENIV2') else
if (Combo.value = '12')then SetControlText('XX_VARIABLE5','ATB_FAMILLENIV3') else
if (Combo.value = '21')then SetControlText('XX_VARIABLE5','ATB_LIBREART1') else
if (Combo.value = '22')then SetControlText('XX_VARIABLE5','ATB_LIBREART2') else
if (Combo.value = '23')then SetControlText('XX_VARIABLE5','ATB_LIBREART3');

ParamTB := TParamTBAffaire.Create(ttbAffaire,TypeDateTB,NiveauTB,GetControlText('XX_VARIABLE3'),False);
// Reprise des paramètre de lancement de l'état
ParamTB.ChargeCriteres (Ecran);

CCheck := TCheckBox(GetControl('ALIMPREVU'));
if CCheck <> Nil then ParamTB.AlimPrevu :=CCheck.Checked ;
CCheck := TCheckBox(GetControl('ALIMFACTURE'));
if CCheck <> Nil then  ParamTB.AlimFacture :=CCheck.Checked ;
CCheck := TCheckBox(GetControl('ALIMREAL'));
if CCheck <> Nil then  ParamTB.AlimReal :=CCheck.Checked ;
ParamTB.AlimAchat := ParamTB.AlimReal;
CCheck := TCheckBox(GetControl('ALIMBONI'));
if CCheck <> Nil then  begin
   ParamTB.AlimBoni :=CCheck.Checked ;
   if (CCheck.checked=True) then begin
     Combo1:= THValComboBox(GetControl('XX_VARIABLE10'));
     if Combo1 <>Nil then begin
              Combo := THValComboBox(GetControl('XX_VARIABLE1'));
              if (Combo.value = 'ATB_TOTPRCHARGE') then Combo1.value := 'ATB_BONIPR'
                                  else   Combo1.value := 'ATB_BONIVENTE';
             end;
     END;
   end;
Combo1:= THValComboBox(GetControl('XX_VARIABLE1'));
if Combo1 <>Nil then begin
        if (Combo1.value = 'ATB_PRRESS') or (Combo1.value = 'ATB_PVRESS') or
           (Combo1.value = 'ATB_PVART') or (Combo1.value = 'ATB_PRART') then
              ParamTB.AlimValoRess_Art :=true ;
        end;
Combo1:= THValComboBox(GetControl('XX_VARIABLE7'));
if Combo1 <>Nil then begin
        if (Combo1.value = 'ATB_PRRESS') or (Combo1.value = 'ATB_PVRESS') or
           (Combo1.value = 'ATB_PVART') or (Combo1.value = 'ATB_PRART') then
              ParamTB.AlimValoRess_Art :=true ;
        end;
CCheck := TCheckBox(GetControl('ALIMAN'));
if CCheck <> Nil then ParamTB.AlimAN :=CCheck.Checked ;
CCheck := TCheckBox(GetControl('AFFSURPERIODE'));
if CCheck <> Nil then ParamTB.AffSurPeriode :=CCheck.Checked ;
CCheck := TCheckBox(GetControl('ALIMVISA'));
if CCheck <> Nil then ParamTB.AlimVisa :=CCheck.Checked ;
CCheck := TCheckBox(GetControl('ALIMMOIS'));
if CCheck <> Nil then ParamTB.AlimMois :=CCheck.Checked ;
CCheck := TCheckBox(GetControl('ALIMCUTOFF'));
if CCheck <> Nil then  begin
   ParamTB.AlimCutOff :=CCheck.Checked ;
   if (CCheck.checked=True) then begin
     Combo1:= THValComboBox(GetControl('XX_VARIABLE10'));
     if Combo1 <>Nil then begin
              Combo := THValComboBox(GetControl('XX_VARIABLE1'));
              if (Combo.value = 'ATB_TOTPRCHARGE') then Combo1.value := 'ATB_CUTOFF'
                                  else   Combo1.value := 'ATB_CUTOFF';
             end;
     end;
   end;
CCheck := TCheckBox(GetControl('CUTOFFMOIS'));
if CCheck <> Nil then     ParamTB.CutOffMois :=CCheck.Checked
   else ParamTB.CutOffMois :=False ;
CCheck := TCheckBox(GetControl('AVECSSAFF'));
if CCheck <> nil then AvecSsAff :=CCheck.Checked else AvecSsAff := False;
CCheck := TCheckBox(GetControl('CENTRFAM'));
if CCheck <> nil then ParamTB.CentrFam :=CCheck.Checked ;
CCheck := TCheckBox(GetControl('CENTRSTATRES'));
if CCheck <> nil then ParamTB.CentrStatRes :=CCheck.Checked ;
CCheck := TCheckBox(GetControl('CENTRSTATART'));
if CCheck <> nil then ParamTB.CentrStatArt :=CCheck.Checked ;
  // mcd 02/04/03
  // Gestion du Facturable / Non facturable
CCheck := TCheckBox(GetControl('FACTURABLE'));if CCheck <> nil then ParamTB.GereFacturable := CCheck.Checked else ParamTB.GereFacturable :=False;

if (TFQRS1(Ecran).NatureEtat = 'APF') then
  begin  // Appréciation
      // mcd 20/08/03 pour gérer cas apprécaiiton en solde et pas AN
  if GetParamSoc('So_AFAPPANOUVEAU') then
    begin
    Paramtb.alimSolde:=True;
    end;
  end;

stChamps := ''; stSQL := '';

// Natures de pièces d'achat reprises
MultiCombo := THMultiValComboBox(GetControl('NATUREPIECEGACHAT'));
if MultiCombo <> Nil then NatPieceAchat := MultiCombo.Text else NatPieceAchat:='';
If ParamTb.alimAchat =False then NatPieceAchat:='';       // mcd 06/08/02 si pas achat on prend tout
// Natures de pièces d'achatpour engage
MultiCombo := THMultiValComboBox(GetControl('NATUREPIECEGACHATENG'));
if MultiCombo <> Nil then NatPieceAchatEngage := MultiCombo.Text else NatPieceAchatEngage:='';
If ParamTb.alimAchat =False then NatPieceAchatEngage:='';       // mcd 06/08/02 si pas achat on prend tout
if (ctxScot in V_PGI.PGIContexte) then NatPieceAchatEngage:=''; // mcd 27/12/02 pas d'engagé pour GI
ParamTB.AlimTableauBordAffaire(stchamps,stSQL,NatPieceAchat,NatPieceAchatEngage,False,False,False,False,False,AvecSSAff);
ParamTB.Free;
End;

procedure TOF_TBAffaire.ImpactRegroupementEdit (Sender: TObject);
Var Combo : THValComboBox;
BEGIN
Combo := THValComboBox(GetControl('REGROUPEMENT'));
if (Combo.value = 'DET') then
    BEGIN
    SetControlEnabled ('NIVEAUREGROUPETB', False);
    SetControlText ('NIVEAUREGROUPETB', '');
    END
else SetControlEnabled ('NIVEAUREGROUPETB', True);
END;


procedure TOF_TBAffaire.TableauObjectsInvisibles(Crit:string; var iNbChamps:integer; var tbChamps:PString);
begin
if (Crit='RESS') then
    begin
    iNbChamps := NbTbChInv;
    tbChamps := @TbChampsInvisibles;
    end;
end;


//******************************************************************************
//**************** TOB d'affichage du tableau de bord **************************
//******************************************************************************
procedure TOF_TBAffViewer.OnArgument(stArgument : String );
Var CC    : TCheckBox;
    Combo : THValComboBox;
    Zone : THMultiValComboBox;
    Menu : TPopUpMenu;
    Critere, Champ, Valeur,Part0,Part1,Part2,Part3,Avenant ,req: string;
    X     : integer;
Begin
Inherited;
TypeAppelTB := '';
bMulti := False;
zbut := '';
// recup des arguments si TB appelé depuis une affaire ou un client
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <> '' ) do
    BEGIN
    if Critere <> '' then
        BEGIN
        X:=pos(':',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end ;
        if Champ = 'TYPETB'      then TypeAppelTB := Valeur
        else if Champ = 'AFFAIRE' then
          begin
          SetControlText('ATBXAFFAIRE',Valeur) ;
            //mcd 10/06/02 ajout champs enabled si passé en paramètre
          SetControlEnabled ('ATBXAFFAIRE1',False);
          SetControlEnabled ('ATBXAFFAIRE2',False);
          SetControlEnabled ('ATBXAFFAIRE3',False);
          SetControlEnabled ('ATBXAVENANT',False);
          SetControlEnabled ('ATBXTIERS',False);
          setControlEnabled ('BSELECTAFF1',False);
          setControlEnabled ('BEFFACEAFF1',False);
          BchangeTiers :=False;
          end
        else if Champ = 'TIERS' then
          begin
          SetControlText('ATBXTIERS',Valeur);
          SetControlEnabled ('ATBXTIERS',False);
          BchangeTiers :=False;
          end
        else if Champ = 'MULTI' then
          begin
          if Valeur='X' then bMulti := True;
          end
        else if Champ = 'BUT' then zbut := 'APP'   // appreciation
        else if champ = 'NOFILTRE' then  TFstat(Ecran).FiltreDisabled:=true ;  //gm 17/07/02
        END;
    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;
// Suppression des champs mois de cloture en Gestion d'Affaire
if Not (ctxScot in V_PGI.PGIContexte) then
   begin
   SetControlVisible('ATBXMOISCLOTURE',False); SetControlVisible('ATBXMOISCLOTURE_',False);
   SetControlVisible('TATB_MOISCLOTURE',False); SetControlVisible('TATB_MOISCLOTURE_',False);
   if (THEdit(GetControl('ATBXAFFDATEDEBEXER'))<>nil) then SetControlVisible('ATBXAFFDATEDEBEXER',False);
   if (THEdit(GetControl('ATBXAFFDATEDEBEXER_'))<>nil) then SetControlVisible('ATBXAFFDATEDEBEXER_',False);
   if (THLabel(GetControl('TAFF_DATEDEBEXER'))<>nil) then SetControlVisible('TAFF_DATEDEBEXER',False);
   if (THLabel(GetControl('TAFF_DATEDEBEXER_'))<>nil) then SetControlVisible('TAFF_DATEDEBEXER_',False);
   if (THEdit(GetControl('ATBXAFFDATEFINEXER'))<>nil) then SetControlVisible('ATBXAFFDATEFINEXER',False);
   if (THEdit(GetControl('ATBXAFFDATEFINEXER_'))<>nil) then SetControlVisible('ATBXAFFDATEFINEXER_',False);
   if (THLabel(GetControl('TAFF_DATEFINEXER'))<>nil) then SetControlVisible('TAFF_DATEFINEXER',False);
   if (THLabel(GetControl('TAFF_DATEFINEXER_'))<>nil) then SetControlVisible('TAFF_DATEFINEXER_',False);
   end
else begin
   if (THEdit(GetControl('ATBXAFFDATELIMITE'))<>nil) then SetControlVisible('ATBXAFFDATELIMITE',False);
   if (THEdit(GetControl('ATBXAFFDATELIMITE_'))<>nil) then SetControlVisible('ATBXAFFDATELIMITE_',False);
   if (THLabel(GetControl('TAFF_DATELIMITE'))<>nil) then SetControlVisible('TAFF_DATELIMITE',False);
   if (THLabel(GetControl('TAFF_DATELIMITE_'))<>nil) then SetControlVisible('TAFF_DATELIMITE_',False);
  end;// appel automatique du TB depuis l'affaire
if (TypeAppelTB = 'AFF') and (GetControlText('ATBXAFFAIRE') <> '') then
   BEGIN
   {$IFDEF BTP}
   BTPCodeAffaireDecoupe(GetControlText('ATBXAFFAIRE'),Part0,Part1,Part2,Part3,Avenant,taModif,false);   
   {$ELSE}
   CodeAffaireDecoupe(GetControlText('ATBXAFFAIRE'),Part0,Part1,Part2,Part3,Avenant,taModif,false);
   {$ENDIF}
   SetControlText('ATBXAFFAIRE1',Part1); SetControlText('ATBXAFFAIRE2',Part2);
   SetControlText('ATBXAFFAIRE3',Part3); SetControlText('ATBXAVENANT',avenant);
   SetControlText('ATBXAFFAIRE0',Part0); 
   END;
// appel automatique du TB depuis le tiers
if (TypeAppelTB = 'T') and (GetControlText('ATBXTIERS') <> '') then
   BEGIN
   SetControlchecked ('TBTIERS',true); ImpactTBTiers(Nil);
   END;
if GetParamSoc('so_afVisaActivite') = False then begin
   SetControlVisible('alimvisa',False);
   SetControlChecked ('AlimVisa',False);
   end;
       // mcd 25/03/02
if GetParamSoc('so_afAppAvecBM') = False then begin
   if TCheckBox(GetControl('ALIMBONI')) <>NIl then begin
    SetControlVisible('alimBoni',False);
    SetControlChecked ('AlimBoni',False);
    end;
  end;
if GetParamSoc('so_afAppCutOff') = False then begin
   if TCheckBox(GetControl('ALIMCutOFF')) <>NIl then begin
    SetControlVisible('alimCutoff',False);
    SetControlChecked ('AlimCutoff',False);
    end;
  end;

  // mcd 06/08/02 cache achat si pas géré plus mise nature gérée par scot
Zone :=ThMultiValCOmboBox(GetControl('NATUREPIECEGACHAT'));
If zone <>Nil then Zone.plus := Zone.plus + AfPlusNatureAchat(true);
Zone :=ThMultiValCOmboBox(GetControl('NATUREPIECEGACHATENG'));
If zone <>Nil then Zone.plus := Zone.plus + AfPlusNatureAchat(true);
if VH_GC.GAAchatSeria =False then
   begin
   SetControlVisible('NATUREPIECEGACHATENG',False);
   SetControlVisible('TNATUREPIECEGACHATENG',False);
   SetControlVisible('NATUREPIECEGACHAT',False);
   SetControlVisible('TNATUREPIECEGACHAT',False);
   SetControlVisible('ACHAT',False);
   SetControlChecked('ACHAT',False);
   end
else if CtxScot in V_PGI.PGIContexte then begin
        // pas gérer en GI, on ne gère que les factures d'achats....
     SetControlVisible('NATUREPIECEGACHATENG',False);
     SetControlVisible('TNATUREPIECEGACHATENG',False);
     Zone:=ThMultiValCOmboBox(Getcontrol('NATUREPIECEGACHAT'));
     if zone <>Nil then Zone.text:='FF;';     //pas de cde pour SCOT
     end;

// Gestion des évènements du TB
TV:=TTobViewer(GetControl('TV'));
TV.OnDblClick := OnDblClickTV;

CC:=TCheckBox(GetControl('TBTIERS')); if CC <>  Nil then CC.OnClick := ImpactTBTiers;
CC:=TCheckBox(GetControl('ECART1'));  if CC <>  Nil then CC.OnClick := ImpactEcart;
ImpactEcart(CC);
CC:=TCheckBox(GetControl('ECART2'));  if CC <>  Nil then CC.OnClick := ImpactEcart;
ImpactEcart(CC);
CC:=TCheckBox(GetControl('ECART3'));  if CC <>  Nil then CC.OnClick := ImpactEcart;
ImpactEcart(CC);
CC:=TCheckBox(GetControl('ECART4'));  if CC <>  Nil then CC.OnClick := ImpactEcart;
ImpactEcart(CC);
CC:=TCheckBox(GetControl('ECART5'));  if CC <>  Nil then CC.OnClick := ImpactEcart;
ImpactEcart(CC);
CC:=TCheckBox(GetControl('TOTALCHAMPS')); if CC <>  Nil then CC.OnClick := ImpactTotalChamps;
ImpactTotalChamps(CC);
CC:=TCheckBox(GetControl('ACHAT')); if CC <>  Nil then CC.OnClick := ImpactAchat;
ImpactAchat(CC);
req:='';
If ctxscot in V_PGI.PgiCOntexte then Req:= req+ 'And CO_CODE NOT LIKE "5%"'; // pas d'engagé en GI
if not(GetParamSoc('SO_AFAPPAVECBM')) then req:=req+ 'AND CO_CODE NOT LIKE "3%"'; // pas de gestion de boni/mali
Menu:=TPopUpMenu(GetControl('POPUPMENUECART1')); if Menu <>  Nil then ChargePopUp(Menu,MenuPopOnClick,'AFR',req);
Menu:=TPopUpMenu(GetControl('POPUPMENUECART2')); if Menu <>  Nil then ChargePopUp(Menu,MenuPopOnClick,'AFR',req);
Menu:=TPopUpMenu(GetControl('POPUPMENUECART3')); if Menu <>  Nil then ChargePopUp(Menu,MenuPopOnClick,'AFR',req);
Menu:=TPopUpMenu(GetControl('POPUPMENUECART4')); if Menu <>  Nil then ChargePopUp(Menu,MenuPopOnClick,'AFR',req);
Menu:=TPopUpMenu(GetControl('POPUPMENUECART5')); if Menu <>  Nil then ChargePopUp(Menu,MenuPopOnClick,'AFR',req);


// if ctxScot in V_PGI.PGIContexte then BEGIN SetControlVisible ('ACHAT', False); SetControlProperty ('ACHAT','Checked',false); END;
// Initialisation
SetControlText ('UNITE',VH_GC.AFMesureActivite);
// SetControlText ('NIVEAUREGROUPETB', '4'); // par article
SetControlText ('NIVEAUREGROUPETB', '6'); // par article  ressource (gm 19/09/02)
if bMulti then SetcontrolText ('REGROUPEMENT','CUM')  // en cumul (champs non visible)
          else if  GetControlText('ATBXAFFAIRE') <>''  //mcd 22/05/03 ajout test pour passer en det si appel depuis TB global
                 then SetcontrolText ('REGROUPEMENT','DET')
                 else SetcontrolText ('REGROUPEMENT','MOI'); // par mois
SetControltext ('MONTANTACTIVITE',GetValoActDefaut(True));
Combo := THValComboBox(GetControl('REGROUPEMENT'));
Combo.OnClick := ImpactRegroupementView;
setControlVisible('AVECSSAFF',GereSousAffaire);
if (VH_GC.AFDATEANALYSETB <>iDate1900) then
  begin
  SetControlText('ATBXDATE',DateTostr(VH_GC.AFDATEANALYSETB) );
  SetControltext('ATBXDATE_',DatetoStr(FinDeMois(PlusDate(VH_GC.AFDATEANALYSETB,11,'M'))));
  end;
if (ctxscot in V_PGI.PGIcontexte) then   //mcd 17/06/03 pour aligner avec impression
   begin
   SetControlText('ATBXDATE',DateTostr(GetParamSoc('SO_AFDATEDEBCAB') ));
   SetControltext('ATBXDATE_',DatetoStr(GetParamSoc('SO_AFDATEFINCAB')));
   end;
if (TypeAppelTB = 'AFF') or (TypeAppelTB = 'T') then
  begin // mcd 16/12/02 si appel depuis client ou mission, on prend sur une péridoe date globale
  SetControlText('ATBXDATE',DateTostr(Idate1900));
  SetControltext('ATBXDATE_',DateTostr(Idate2099));
  end;
(*if ((TypeAppelTB = 'AFF') and (GetControlText('AFFAIRE') <> '')) or
   ((TypeAppelTB = 'T')   and (GetControlText('TIERS')   <> '')) then
   BEGIN
   Bt := TToolBarbutton97(GetControl('Bcherche'));
   if Bt <> nil then Bt.OnClick(Bt);
   END; pb car avant reprise de la présentation sauvegardée ...*)
SetControlText('ATBXSTATUTAFFAIRE','AFF' );
If GetParamSoc('SO_AFPROPOSACT')=False then
   begin
   SetControlVisible ('ATBXSTATUTAFFAIRE',False);
   SetControlVisible ('TATB_STATUTAFFAIRE',False);
   end;

{$IFDEF EAGLCLIENT}  // mcd le bouton bimprimer ne marche ps en eAgl sur des tob. donc il faut cacher le bouton
  setControlVisible('BIMPRIMER', false);
{$ENDIF}

End;


//****** Récupération du paramétrage et lancement du tob viewer ****************
procedure TOF_TBAffViewer.OnLoad;
Var CodeAffaire,CodeTiers,ListeChampsSel,Part0,Part1,Part2,Part3,Part4,stSQL,NatPieceAchat,NatPieceAchatEngage: string;
    CC   : TCheckBox;
    MultiCombo : THMultiValComboBox;
    TypeTB    : T_TypeTB;
    NiveauTiers, AvecDetailAff,NivAffRef,ChargementSelectif,ChampsRes,
    ChampsArt,AvecSsAff : Boolean;

BEGIN
  inherited;
  CodeAffaire := ''; CodeTiers := ''; stSQL := '';
  //AvecSsAff:=false;
  if ParamTB <> Nil then BEGIN ParamTB.Free; ParamTB := nil; END;
  //if TFStat(Ecran).LaTOB <> Nil then BEGIN TFStat(Ecran).LaTOB := Nil; END;
  CC := TCheckBox(GetControl('TBTIERS'));
  if CC <> nil then NiveauTiers := CC.Checked else NiveauTiers := False;
  CC := TCheckBox(GetControl('DETAILPARAFFAIRE'));
  if CC <> nil then AvecDetailAff :=CC.Checked else AvecDetailAff := False;
  CC := TCheckBox(GetControl('NIVANALYSEAFFAIRE'));
  if CC <> nil then NivAffRef :=CC.Checked else NivAffRef := False;
  CC := TCheckBox(GetControl('AVECSSAFF'));
  if CC <> nil then AvecSsAff :=CC.Checked else AvecSsAff := False;

  TypeTB := GetTypeTB (NiveauTiers, AvecDetailAff, bMulti);
  ParamTB := TParamTBAffaire.Create(TypeTB,GetRegroupeTB (GetControlText('REGROUPEMENT')),
             GetNiveauTB (GetControlText('NIVEAUREGROUPETB')),GetControlText('UNITE'),NivAffRef);

  // Quels éléments doivent être pris en compte
  CC := TCheckBox(GetControl('PREVU')); if CC <> Nil then ParamTB.AlimPrevu := CC.Checked;
  CC := TCheckBox(GetControl('FACT')); if CC <> Nil then ParamTB.AlimFacture:= CC.Checked;
  CC := TCheckBox(GetControl('REAL')); if CC <> Nil then ParamTB.AlimReal := CC.Checked;
  CC := TCheckBox(GetControl('ACHAT')); if CC <> Nil then ParamTB.AlimAchat:= CC.Checked;
  CC := TCheckBox(GetControl('ALIMVISA')); if CC <> Nil then ParamTB.AlimVisa:= CC.Checked;
  CC := TCheckBox(GetControl('ALIMBONI')); if CC <> Nil then ParamTB.AlimBoni:= CC.Checked;
  CC := TCheckBox(GetControl('ALIMCUTOFF')); if CC <> Nil then ParamTB.AlimCutOff:= CC.Checked;
  CC := TCheckBox(GetControl('CUTOFFMOIS')); if CC <> Nil then ParamTB.CutOffMois:= CC.Checked;
  CC := TCheckBox(GetControl('ALIMVALORESS_ART')); if CC <> Nil then ParamTB.AlimValoRess_Art:= CC.Checked;
  if Not(bMulti) then
    BEGIN
    // recup code client ou affaire
    if NiveauTiers then
       BEGIN // TB par client
       CodeTiers := GetControlText('ATBXTIERS');
       ParamTB.MajMonotiers (CodeTiers,GetcontrolText('ATBXDATE'),GetcontrolText('ATBXDATE_'));
       if not ExisteTiers (CodeTiers) then
            BEGIN
            PGIBoxAF ('Tiers non valide','Tableau de bord'); Exit;
            END;
       ParamTB.DetailParAffaire := AvecDetailAff;
       END
    else
       BEGIN
       ParamTB.DetailParAffaire := False;
       //mcd 30/09/03 pour ok depuis proposition Part0 :='A';
       Part0 := GetControlText('ATBXAFFAIRE0');
       Part1 := GetControlText('ATBXAFFAIRE1');
       Part2 := GetControlText('ATBXAFFAIRE2');
       Part3 := GetControlText('ATBXAFFAIRE3');
       Part4 := GetControlText('ATBXAVENANT');
       CodeAffaire := CodeAffaireRegroupe(Part0, Part1  ,Part2 ,Part3,Part4, taModif, false,false,False);
       SetControlText('ATBXAFFAIRE',CodeAffaire);
       if not ExisteAffaire (CodeAffaire,'') then
            BEGIN
            if TestCleAffaire(THCritMaskEdit(GetControl('ATBXAFFAIRE')),THCritMaskEdit(GetControl('ATBXAFFAIRE1')),
               THCritMaskEdit(GetControl('ATBXAFFAIRE2')),THCritMaskEdit(GetControl('ATBXAFFAIRE3')),
               THCritMaskEdit(GetControl('ATBXAVENANT')), THCritMaskEdit(GetControl('ATBXTIERS')),
               Part0,False,False,False) <> 1 then
                 BEGIN
                 PGIBoxAF ('Affaire non valide','Tableau de bord');
                 SetControlText('ATBXAFFAIRE',''); Exit;
                 END;
            CodeAffaire := GetControlText('ATBXAFFAIRE');
            END;
       // Recherche si affaire de référence pour reprise des sous affaires
       if GereSousAffaire then
        begin
        if GetChampsAffaire( CodeAffaire, 'AFF_ISAFFAIREREF')='X' then
           begin
           if (PGIAskAf ('Voulez-vous visualiser les sous-affaires','Gestion des sous-affaires')= mrYes) then
              begin AvecSsAff :=true; ParamTB.DetailParAffaire:=true; end;
           end;
        end;
       ParamTB.MajMonoAffaire (CodeAffaire,GetcontrolText('ATBXDATE'),GetcontrolText('ATBXDATE_'),AvecSsAff);
       END;
    ParamTB.ChargeCriteres (Ecran);// mcd 17/09/02 ajouter pour prendre en compte les critères éventuels de la fiche
    END
  else  // En Multiclient
    BEGIN
    ParamTB.DetailParAffaire := Not(NiveauTiers);
    ParamTB.ChargeCriteres (Ecran);
    END;
  // Valorisation de l'activité reprise
  if GetControltext('MONTANTACTIVITE')= 'PR' then ParamTB.MtActivite := ActPR else
  if GetControltext('MONTANTACTIVITE')= 'PV' then ParamTB.MtActivite := ActPV else
  if GetControltext('MONTANTACTIVITE')= 'PRV'then ParamTB.MtActivite := ActPRPV;

  // Gestion du Facturable / Non facturable
  CC := TCheckBox(GetControl('FACTURABLE')); ParamTB.GereFacturable := CC.Checked;
  // Gestion du chargement sélectif
  CC := TCheckBox(GetControl('TOTALCHAMPS')); ChargementSelectif := Not(CC.Checked);
  // Gestion des compléments articles + ressources
  CC := TCheckBox(GetControl('CHAMPSRESSOURCE')); ChampsRes := CC.Checked;
  CC := TCheckBox(GetControl('CHAMPSARTICLE'));   ChampsArt := CC.Checked;
  // Natures de pièces d'achat reprises
  MultiCombo := THMultiValComboBox(GetControl('NATUREPIECEGACHAT'));
  if MultiCombo <> Nil then NatPieceAchat := MultiCombo.Text else NatPieceAchat:='';
  if ParamTb.alimAchat=False then NatPieceAchat:=''; // mcd 06/08/02 si pas achat on prend tout
  // Natures de pièces d'achat reprises  pour l'engagé
  MultiCombo := THMultiValComboBox(GetControl('NATUREPIECEGACHATENG'));
  if MultiCombo <> Nil then NatPieceAchatEngage := MultiCombo.Text else NatPieceAchatEngage:='';
  if ParamTb.alimAchat=False then NatPieceAchatEngage:='';
  if (ctxScot in V_PGI.PGIContexte) then NatPieceAchatEngage:=''; // mcd 27/12/02 pas d'engagé pour GI
  // gestion des écarts
  AlimEcartTB;
  // lancement du TB en consultation par une tob

  TobTB := ParamTB.AlimTableauBordAffaire(ListeChampsSel,stSQL,NatPieceAchat,NatPieceAchatEngage,Not(bMulti),bMulti,ChargementSelectif,ChampsRes,ChampsArt,AvecSsAff);
  // alim du tobViewer
  if bmulti then  // en multi on peut retourner beaucoup d'enregistrements => passage par une requête.
     BEGIN
     TFStat(Ecran).FSQL.Lines[0] := stSQL;
     END
  else
     BEGIN
     if TOBTB <> Nil then
         BEGIN
         TFStat(Ecran).LaTOB :=  TOBTB;
         TFStat(Ecran).ColNames := ListeChampsSel;
         END
     else PGIBoxAF ('Aucune ligne affichée','Tableau de bord');
     END;
  // maj des cumuls
  SetControlText ('TOTPREVU', Floattostr(Arrondi(ParamTB.SyntheseTB.TotPrevuPV,V_PGI.OkDecV)));

  if (ParamTB.MtActivite = ActPR) or (ParamTB.MtActivite = ActPRPV) then
    SetControlText ('TOTREAL', Floattostr(Arrondi(ParamTB.SyntheseTB.TotRealPR,V_PGI.OkDecV)));
  if (ParamTB.MtActivite = ActPV) or (ParamTB.MtActivite = ActPRPV) then
  SetControlText ('TOTREALV', Floattostr(Arrondi(ParamTB.SyntheseTB.TotRealPV,V_PGI.OkDecV)));
  if (ParamTB.MtActivite = ActPV) then
  Begin
    SetControlVisible('TOTREAL',false);
    SetControlVisible('TTOTREAL',false);
    SetControlVisible('TOTREALV',true);
    SetControlVisible('TTOTREALV',true);
  End;
  if (ParamTB.MtActivite = ActPR) then
  Begin
    SetControlVisible('TOTREALV',false);
    SetControlVisible('TTOTREALV',false);
    SetControlVisible('TOTREAL',true);
    SetControlVisible('TTOTREAL',true);
  End;
  if (ParamTB.MtActivite = ActPRPV) then
  Begin
    SetControlVisible('TOTREALV',true);
    SetControlVisible('TTOTREALV',true);
    SetControlVisible('TOTREAL',true);
    SetControlVisible('TTOTREAL',true);
  End;

  SetControlText ('TOTFACT', Floattostr(Arrondi(ParamTB.SyntheseTB.TotFact,V_PGI.OkDecV)));
  TButton(Getcontrol('BAGRANDIR')).OnClick(Self);
END;

procedure TOF_TBAffViewer.OnClose;
BEGIN
if ParamTB <> Nil then
  begin
    ParamTB.Free;
    ParamTB := nil;
  end;
inherited;
//if LaTOB <> Nil then LaTOB.Free;
//if TOBTB <> Nil then TOBTB.Free;
END;

//******************************************************************************
//********************* Héritage AFBASECODEAFFAIRE *****************************
//******************************************************************************
procedure TOF_TBAffViewer.NomsChampsAffaire(var Aff, Aff0,Aff1, Aff2, Aff3,Aff4, Aff_, Aff0_,Aff1_, Aff2_, Aff3_,Aff4_, Tiers, Tiers_:THEdit);
begin
Aff1:=THEdit(GetControl('ATBXAFFAIRE1'));
Aff2:=THEdit(GetControl('ATBXAFFAIRE2'));
Aff3:=THEdit(GetControl('ATBXAFFAIRE3'));
Aff4:=THEdit(GetControl('ATBXAVENANT'));
Tiers:=THEdit(GetControl('ATBXTIERS'));

Aff1_:=THEdit(GetControl('ATBXAFFAIRE1_'));
Aff2_:=THEdit(GetControl('ATBXAFFAIRE2_'));
Aff3_:=THEdit(GetControl('ATBXAFFAIRE3_'));
Aff4_:=THEdit(GetControl('ATBXAVENANT_'));
Tiers_:=THEdit(GetControl('ATBXTIERS_'));
end;


//******************************************************************************
//************************** Evénements écrans  ********************************
//******************************************************************************
procedure TOF_TBAffViewer.ImpactTBTiers (Sender: TObject);
Var bVisible : Boolean;
    CC   : TCheckBox;
BEGIN
CC := TCheckBox(GetControl('TBTIERS'));
if CC <>  Nil then bVisible := Not(CC.Checked) else bVisible := true;
// réinitialisation des zones affaires
if Not(bVisible) then
   BEGIN
   ChargeCleAffaire(Nil,THEdit(GetControl('ATBXAFFAIRE1')),THedit(GetControl('ATBXAFFAIRE2')),THEdit(GetControl('ATBXAFFAIRE3')),THEdit(GetControl('ATBXAVENANT')),Nil,taModif,'',False);
   if bMulti then
      ChargeCleAffaire(Nil,THEdit(GetControl('ATBXAFFAIRE1_')),THedit(GetControl('ATBXAFFAIRE2_')),THEdit(GetControl('ATBXAFFAIRE3_')),THEdit(GetControl('ATBXAVENANT_')),Nil,taModif,'',False);
   END;
SetControlVisible ('ATBXAFFAIRE1',bVisible); SetControlVisible ('ATBXAFFAIRE2',bVisible);
SetControlVisible ('ATBXAFFAIRE3',bVisible); SetControlVisible ('ATBXAVENANT',bVisible);
SetControlVisible ('TATB_AFFAIRE',bVisible);
SetControlVisible ('BSELECTAFF1',bVisible);SetControlVisible ('BEFFACEAFF1',bVisible);
if bMulti then
   BEGIN
   SetControlVisible ('ATBXAFFAIRE1_',bVisible); SetControlVisible ('ATBXAFFAIRE2_',bVisible);
   SetControlVisible ('ATBXAFFAIRE3_',bVisible); SetControlVisible ('ATBXAVENANT_',bVisible);
   SetControlVisible ('TATB_AFFAIRE_',bVisible);
   SetControlVisible ('BSELECTAFF2',bVisible);SetControlVisible ('BEFFACEAFF2',bVisible);
   END
else
   SetControlVisible ('DETAILPARAFFAIRE',Not(bVisible));   // possibilité de détail avec ou sans affaire
if (bVisible) then
   BEGIN
   ChargeCleAffaire(Nil,THEdit(GetControl('ATBXAFFAIRE1')),THedit(GetControl('ATBXAFFAIRE2')),THEdit(GetControl('ATBXAFFAIRE3')),THEdit(GetControl('ATBXAVENANT')),Nil,taModif,'',False);
   if bMulti then
      ChargeCleAffaire(Nil,THEdit(GetControl('ATBXAFFAIRE1_')),THedit(GetControl('ATBXAFFAIRE2_')),THEdit(GetControl('ATBXAFFAIRE3_')),THEdit(GetControl('ATBXAVENANT_')),Nil,taModif,'',False);
   END;
END;
// **** Affichage élèments sur achats ***
procedure TOF_TBAffViewer.ImpactAchat (Sender: TObject);
Var bVisible : Boolean;
BEGIN
bVisible := TcheckBox(Sender).Checked;
SetControlVisible('NATUREPIECEGACHAT',bVisible); SetControlVisible('TNATUREPIECEGACHAT',bVisible);
if not(ctxscot in V_PGI.PGIContexte)
  then begin
  SetControlVisible('NATUREPIECEGACHATENG',bVisible);
  SetControlVisible('TNATUREPIECEGACHATENG',bVisible);
  end;
END;
// **** Affichage des écarts ou non ***
procedure TOF_TBAffViewer.ImpactEcart (Sender: TObject);
Var stEcart : string;
    bVisible : Boolean;
BEGIN
//mcd 15/10/02 if TCheckBox(Sender).tag = 1 then stEcart := '1' else stEcart := '2';
stecart:=  IntTostr(TCheckBox(Sender).tag );
bVisible := TcheckBox(Sender).Checked;
SetControlEnabled ('TFORMULEECART'+ stEcart,bVisible);
SetControlEnabled ('FORMULEECART'+ stEcart,bVisible);
END;
// **** Affichage des compléments si reprise totale des champs ***
procedure TOF_TBAffViewer.ImpactTotalChamps (Sender: TObject);
Var bVisible : Boolean;
BEGIN
bVisible := TcheckBox(Sender).Checked;
SetControlVisible ('CHAMPSRESSOURCE',bVisible);
SetControlVisible ('CHAMPSARTICLE' , bVisible);
END;
//*** Gestion auto des regroupements / paramétrage ***
procedure TOF_TBAffViewer.ImpactRegroupementView (Sender: TObject);
Var Combo : THValComboBox;
BEGIN
Combo := THValComboBox(GetControl('REGROUPEMENT'));
if (Combo.value = 'DET') then
    BEGIN
    SetControlEnabled ('NIVEAUREGROUPETB', False);
    SetControlText ('NIVEAUREGROUPETB', '');
    END
else SetControlEnabled ('NIVEAUREGROUPETB', True);
END;

//******************************************************************************
//********************* Zooms / double clic  ***********************************
//******************************************************************************
procedure TOF_TBAffViewer.OnDblClickTV(Sender: TObject);
var
Ressource, Affaire, Tiers, Select, titre:string;
dDate,Fdate:TDateTime;
iPeriode, iSemaine:integer;
TypeSaisie : T_Sai;
BEGIN
TypeSaisie := tsaClient; // pour initialiser ....
if (bMulti) and (zbut = 'APP')  then // appel appréciation pour ne rien changer dans ce cas
    Begin
    RecupAffaireLigneTB(Affaire,Tiers);
    if (Affaire = '') and (Tiers = '') then exit;
    AppelAppreciation(Affaire);
    exit;
    End  ;
  //mcd 07/04/2003 tout revu pour permettre plus de choix  sur le clic d'une ligne
if (bMulti) then Titre :='Consultation globale'
 else Titre := traduitGA ('Consultation client/affaire');
Ressource:=''; Affaire:=''; Tiers := ''; dDate:=0; iPeriode:=0; iSemaine:=0;
Voir :=VRien;
if THValComboBox(GetControl('REGROUPEMENT')).value ='DET' then
 begin  // cas détail, on regarde sur quel type de ligne on a cliquer
  // zone sur activite
 If (TV.ColIndex('ATB_QTE')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_QTE'),TV.CurrentRow] <>0) then Voir:=Vactivite;
 If (TV.ColIndex('ATB_QTEUNITEREF')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_QTEUNITEREF'),TV.CurrentRow] <>0) then Voir:=Vactivite;
 If (TV.ColIndex('ATB_TOTPR')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_TOTPR'),TV.CurrentRow] <>0) then Voir:=Vactivite;
 If (TV.ColIndex('ATB_TOTPRCHARGE')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_TOTPRCHARGE'),TV.CurrentRow] <>0) then Voir:=Vactivite;
 If (TV.ColIndex('ATB_TOTPRCHINDI')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_TOTPRCHINDI'),TV.CurrentRow] <>0) then Voir:=Vactivite;
 If (TV.ColIndex('ATB_TOTVENTEACT')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_TOTVENTEACT'),TV.CurrentRow] <>0) then Voir:=Vactivite;
 If (TV.ColIndex('ATB_PRRESS')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_PRRESS'),TV.CurrentRow] <>0) then Voir:=Vactivite;
 If (TV.ColIndex('ATB_PVRESS')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_PVRESS'),TV.CurrentRow] <>0) then Voir:=Vactivite;
 If (TV.ColIndex('ATB_PRART')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_PRART'),TV.CurrentRow] <>0) then Voir:=Vactivite;
 If (TV.ColIndex('ATB_PVART')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_PVART'),TV.CurrentRow] <>0) then Voir:=Vactivite;
  // zone sur facture
 If (TV.ColIndex('ATB_FACPV')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_FACPV'),TV.CurrentRow] <>0) then Voir:=VFActure;
 If (TV.ColIndex('ATB_FACUNITE')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_FACUNITE'),TV.CurrentRow] <>0) then Voir:=VFActure;
 If (TV.ColIndex('ATB_FACQTE')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_FACQTE'),TV.CurrentRow] <>0) then Voir:=VFActure;
 If (TV.ColIndex('ATB_FACQTEUNITEREF')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_FACQTEUNITEREF'),TV.CurrentRow] <>0) then Voir:=VFActure;
  // zone sur Engage
 If (TV.ColIndex('ATB_ENGAGEPR')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_ENGAGEPR'),TV.CurrentRow] <>0) then Voir:=VEngage;
 If (TV.ColIndex('ATB_ENGAGEPV')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_ENGAGEPV'),TV.CurrentRow] <>0) then Voir:=VEngage;
 If (TV.ColIndex('ATB_ENGAGEQTE')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_ENGAGEQTE'),TV.CurrentRow] <>0) then Voir:=VEngage;
   // zone sur Cutoff
 If (TV.ColIndex('ATB_FAE')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_FAE'),TV.CurrentRow] <>0) then Voir:=VCutOff;
 If (TV.ColIndex('ATB_AAE')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_AAE'),TV.CurrentRow] <>0) then Voir:=VCutOff;
 If (TV.ColIndex('ATB_PCA')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_PCA'),TV.CurrentRow] <>0) then Voir:=VCutOff;
 If (TV.ColIndex('ATB_FAR')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_FAR'),TV.CurrentRow] <>0) then Voir:=VCutOff;
 If (TV.ColIndex('ATB_AAR')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_AAR'),TV.CurrentRow] <>0) then Voir:=VCutOff;
 If (TV.ColIndex('ATB_CCA')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_CCA'),TV.CurrentRow] <>0) then Voir:=VCutOff;
  // zone sur Boni
 If (TV.ColIndex('ATB_BONIQTE')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_BONIQTE'),TV.CurrentRow] <>0) then Voir:=VBoni;
 If (TV.ColIndex('ATB_BONIQTEUNITERE')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_BONIQTEUNITERE'),TV.CurrentRow] <>0) then Voir:=VBoni;
 If (TV.ColIndex('ATB_BONIVENTE')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_BONIVENTE'),TV.CurrentRow] <>0) then Voir:=VBoni;
 If (TV.ColIndex('ATB_BONIPR')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_BONIPR'),TV.CurrentRow] <>0) then Voir:=VBoni;
  //zone sur prevu   Bien laisser teste sur qte en 1ER. doit être remplacer par option prix si existe
 If (TV.ColIndex('ATB_PREVUQTE')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_PREVUQTE'),TV.CurrentRow] <>0) then
  begin
  If (GetParamSoc('SO_AFTYPEPREVUPV')='PLA') then voir :=Vpla
  else   If (GetParamSoc('SO_AFTYPEPREVUPV')='BUD') then voir :=VBudg
    else Voir :=Vaffaire;
  end;
 If (TV.ColIndex('ATB_PREQTEUNITEREF')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_PREQTEUNITEREF'),TV.CurrentRow] <>0) then
  begin
  If (GetParamSoc('SO_AFTYPEPREVUPV')='PLA') then voir :=Vpla
  else   If (GetParamSoc('SO_AFTYPEPREVUPV')='BUD') then voir :=VBudg
    else Voir :=Vaffaire;
  end;
 If (TV.ColIndex('ATB_PREVUPR')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_PREVUPR'),TV.CurrentRow] <>0) then
  begin
  If (GetParamSoc('SO_AFTYPEPREVUPR')='PLA') then voir :=Vpla
  else   If (GetParamSoc('SO_AFTYPEPREVUPR')='BUD') then voir :=VBudg
    else Voir :=Vaffaire;
  end;
 If (TV.ColIndex('ATB_PREVUPV')>-1)  and (TV.AsDouble[TV.ColIndex('ATB_PREVUPV'),TV.CurrentRow] <>0) then
  begin
  If (GetParamSoc('SO_AFTYPEPREVUPV')='PLA') then voir :=Vpla
  else   If (GetParamSoc('SO_AFTYPEPREVUPV')='BUD') then voir :=VBudg
    else Voir :=Vaffaire;
  end;
 end
else begin // multi, appel de la fiche qui va saisir le type de consultation
  Select := AGLLanceFiche('AFF','AFTBMENU','','','');
  if (Select='1') then  Voir :=VActivite
  else if (Select='2') then Voir := VFacture
  else if (Select='3') then Voir := VBoni
  else if (Select='4') then Voir := VCutoff
  else if (Select='5') then Voir := VEngage
  else if (Select='6') then Voir := VPla
  else if (Select='7') then Voir := VBudg
  else if (Select='8') then Voir := VArt
  else if (Select='A') and (Bmulti) then Voir := VTB
  else if (Select='B') then Voir := VApp
  else if (Select='9') then Voir := VAffaire  ;
 end;
if TV.CurrentRow < 0 then Exit; // pas de sélection sur une ligne
If Voir <> VRien then
 begin
 if bMulti then   // TB multi client  : On lance le TB détaillé
    BEGIN
    RecupAffaireLigneTB(Affaire,Tiers);
    end
 else begin    // si pas multi, code affaire et clent dans les onglets de sélections
    affaire :=GetControlText ('ATBXAFFAIRE');
    Tiers :=GetControlText ('ATBXTIERS');
    if affaire ='' then RecupAffaireLigneTB(Affaire,Tiers); // cas ou lcient renseugner dans mul, mais pas affaire
   end ;
   // les dates sont souvent utilisée, donc on les stocke une fois pour tous les cas
 if (TV.ColIndex('ATB_DATE')>-1) then DDate := TV.AsDateTime[TV.ColIndex('ATB_DATE'),TV.CurrentRow];
 Fdate := Ddate;
 if THValComboBox(GetControl('REGROUPEMENT')).value <>'DET' then Fdate := Findemois(DDate); //mcd 21/10/03 pour voir tout le mois
 if (TV.ColIndex('ATB_PERIODE')>-1) then iPeriode := TV.AsInteger[TV.ColIndex('ATB_PERIODE'),TV.CurrentRow];
 if (TV.ColIndex('ATB_SEMAINE')>-1) then isemaine := TV.AsInteger[TV.ColIndex('ATB_SEMAINE'),TV.CurrentRow];
 if ((dDate=0) or (dDate=iDate1900) or (dDate=iDate2099)) then
    begin
    if (iPeriode<>0) then begin
        dDate:=GetDateDebutPeriode(iPeriode);
        Fdate := Findemois(DDate);
        end;
    if (iPeriode<>0) and (iSemaine<>0) then begin
      dDate:=GetDateDebutSemaine(iSemaine, iPeriode);
      Fdate := Findemois(DDate);
      end;
    end;
 if ((dDate=0) or (dDate=iDate1900)) then begin
     DDate := StrtoDate(GetControltext ('ATBXDATE'));
     Fdate := StrtoDate(GetControltext ('ATBXDATE_'));
     end;
 If Voir =VTB then
   begin
    if (Affaire = '') and (Tiers = '') then
        begin
        PgiBoxAF ('Pas de client ou affaire sur cette ligne',Titre);
        exit;
        end;
    if (Affaire <> '') then
          AGLLanceFiche('AFF','AFTBVIEWER','','','TYPETB:AFF;AFFAIRE:'+Affaire+';TIERS:'+Tiers)
      else
          AGLLanceFiche('AFF','AFTBVIEWER','','','TYPETB:T;TIERS:'+Tiers);
    End
 else If Voir =VActivite then
   BEGIN
   if (TV.ColIndex('ATB_RESSOURCE')>-1) then Ressource := TV.AsString[TV.ColIndex('ATB_RESSOURCE'),TV.CurrentRow];
   if (Ressource<>'') then TypeSaisie := tsaRess
     else if (Affaire<>'') then TypeSaisie := tsaClient;
 //  if ((dDate=0) or (dDate=iDate1900) or (dDate=iDate2099) or ((Ressource='') and (Affaire=''))) then
   if  (Ressource='') and (Affaire='') then
     Begin
     PgiInfoAf('Pas de ressource, date ou affaire sur cette ligne',Titre);
     exit;
     end;
   if THValComboBox(GetControl('REGROUPEMENT')).value <>'CUM' then
       AFCreerActiviteModale(TypeSaisie, tacGlobal, 'REA', Ressource, Affaire, Tiers, dDate  )
   else AFLanceFiche_Consult_Activite ('AFF='+affaire+ ';DATEDEB='+ DateToStr(DdAte)+ ';DATEFIN='+ DateToStr(FDate));
   End
{$ifdef GIGI}
 else If Voir =VApp then      // appréciation
   BEGIN
    if (Affaire = '') and (Tiers = '') then
        begin
        PgiBoxAF ('Pas de client ou affaire sur cette ligne',Titre);
        exit;
        end;
    AFLanceFiche_Mul_Consult_Apprec ('TIERS:'+tiers+';ZAFF:'+Affaire + ';DATEDEB:'+ DateToStr(DdAte)+ ';DATEFIN:'+ DateToStr(FDate));
   End
 else If Voir =VBoni then
   begin   // fct de consultaiton boni/mali
    if (Affaire = '') and (Tiers = '') then
        begin
        PgiBoxAF ('Pas de client ou affaire sur cette ligne',Titre);
        exit;
        end;
   AFLanceFiche_ModifBoni_Mul('TIERS='+tiers+';AFFAIRE='+ Affaire+ ';DATE='+datetostr(Ddate)+';DATEFIN='+datetostr(Fdate)+';ACTION=CONSULTATION');
   end
{$endif}
 else If Voir =VArt then
   begin
   if (TV.ColIndex('ATB_CODEARTICLE')>-1) then
      begin
      Tiers := TV.AsString[TV.ColIndex('ATB_CODEARTICLE'),TV.CurrentRow];
      if tiers <>'' then V_PGI.DispatchTT( 7, taModif,tiers, '', 'MONOFICHE')
       else PgiInfoAf ('Code article non renseigné',Titre);
      end
   else PgiInfoAf ('Code article non renseigné',Titre);
   end
 else If Voir =VAffaire then
   begin
    if (Affaire = '') then
        begin
        PgiBoxAF ('Pas d''affaire sur cette ligne',Titre);
        exit;
        end;
   V_PGI.DispatchTT( 5, taModif, Affaire, '', 'MONOFICHE') ;
   end
 else If Voir =VBudg then
   begin
   if ((dDate=0) or (dDate=iDate1900) or (dDate=iDate2099) or ((Ressource='') and (Affaire=''))) then
     begin
     PgiInfoAf('Pas de ressource, date ou affaire sur cette ligne',Titre);
     exit;
     end;
   tiers := 'ZAFF='+affaire;
   tiers := tiers+ ';ZDEB='+DateToStr(dDate);
   tiers := tiers + ';ZFIN='+DateToStr(PlusDate(Ddate,1,'A'));
   AglLanceFiche('AFF','AFBUDGET','','','ACTION=MODIFICATION;'+tiers );
   end
 else If Voir =VFacture then
   begin   // accès pièce
    if (Affaire = '') and (Tiers = '') then
        begin
        PgiBoxAF ('Pas de client ou affaire sur cette ligne',Titre);
        exit;
        end;
   AFLanceFiche_Mul_RegroupLigne('GL_TIERS='+Tiers+'; GL_AFFAIRE='+Affaire,'NOAFFAIRE;NATUREAUXI:CLI;TABLE:GL;DATEDEB:'+datetostr(Ddate)+';DATEFIN:'+datetostr(Fdate)+';AFFAIRE:'+affaire);
   end
 else If Voir =VCutOff then
   begin
    if (Affaire = '') and (Tiers = '') then
        begin
        PgiBoxAF ('Pas de client ou affaire sur cette ligne',Titre);
        exit;
        end;
   AFLanceFiche_ModifCutOffMul ('TYPE:CVE;SAISIE:DETAIL;AFFAIRE:'+Affaire+';DATE:'+datetostr(FinDeMois(Ddate))+';TIERS:'+tiers)
   end
 else If Voir =VPla then
   begin
    AFLanceFicheAFPlanning_Mul('', 'PLANNING1' + ';NUMPLANNING:153201;AFFAIRE:'+Affaire+';TIERS:'+tiers );
   end
 else If Voir =VEngage then
   begin // pas de consultation de l'engagé au 04/2003
   end ;
 end;
END;

// appel de l'apprecaition depuis le Tableau de bord
Procedure   TOF_TBAffViewer.AppelAppreciation(zaff : string);
{$IFDEF GIGI}
var
//affaire,st, ,titre
     zdate,zdatapp,zori : string;
     ret : string;
{$endif}
BEGIN
{$IFDEF GIGI}
      // mcd 12/06/02	Zdate :=AGLLanceFiche('AFF','AFSAISDATE','','','ZORI:A;') ;
  Zdate:=AFLanceFiche_SaisieDate( 'ZORI:A');
  if (zdate = '0') then
   begin
//       PGIInfoAF('Traitement abandonné', titre); // PL le 07/08/03 : y'avait comme un truc bizarre...
       PGIInfoAF('Traitement abandonné', Ecran.caption);
       exit;
     end
   else
   Begin
   		zdatapp := zdate;
   End;

//  St:= 'Confirmez vous l''appréciation de cette affaire au ' + zdate;

//	If (PGIAskAF(st,titre)<> mrYes) then exit;

  if (GetParamSoc('SO_AFAPPPOINT')=true)  then zori:='COM' else zori :='SIM';
  if (zori = 'SIM') then
    Ret := AFLanceFiche_AppreciationNiv0('ZORI:'+zori+';ZAFF:'+zaff+';ZDATAPP:'+zdatapp+';')
  else
    Ret := AFLanceFiche_AppreciationCompl('ZORI:'+zori+';ZAFF:'+zaff+';ZDATAPP:'+zdatapp+';');


{$endif}

END;

// ***** Bt Zoom Affaire ****
procedure TOF_TBAffViewer.LanceAffaireTB;
Var Codeaffaire, Tiers : string;
  CC : TcheckBox;
  BmultiForce:boolean;
    //Part0, Part1, Part2, Part3, Part4 : string;
BEGIN
BmultiForce :=False;
CC:=TCheckBox(GetControl('TBTIERS'));
if CC <>  Nil then BmultiForce :=CC.checked;
if (bMulti) or (BmultiForce) then
   RecupAffaireligneTB(codeAffaire,Tiers)
else
   begin
//   Part0 :='A'; Part1 := GetControlText('ATBXAFFAIRE1'); Part2 := GetControlText('ATBXAFFAIRE2');
//   Part3 := GetControlText('ATBXAFFAIRE3'); Part4 := GetControlText('ATBXAVENANT');
//   CodeAffaire := CodeAffaireRegroupe(Part0, Part1  ,Part2 ,Part3,Part4, taModif, false,false,false);

   // PL le 18/01/02 car ne reconnait pas la mission si le nombre de partie n'est pas maxi
   TFStat(Ecran).ChercheClick;
   CodeAffaire :=  GetControlText('ATBXAFFAIRE');
   // Fin PL le 18/01/02
   end;

if not ExisteAffaire (CodeAffaire,'') then
   BEGIN PGIBoxAF ('Affaire non valide','Tableau de bord'); Exit; END
else V_PGI.DispatchTT( 5, taModif, CodeAffaire, '', 'MONOFICHE') ;
END;

Procedure TOF_TBAffViewer.RecupAffaireLigneTB(Var Affaire, tiers : string);
Var Affaire0,Affaire1,Affaire2,Affaire3,Avenant : string;
begin
Affaire :=''; Tiers :=''; Affaire0 := 'A'; Affaire1 := ''; Affaire2 := ''; Affaire3 := ''; Avenant:='00';
if TV.CurrentRow < 0 then Exit;

if (TV.ColIndex('ATB_AFFAIRE1')>-1) then Affaire1:=TV.AsString[TV.ColIndex('ATB_AFFAIRE1'),TV.CurrentRow];
if (TV.ColIndex('ATB_AFFAIRE2')>-1) then Affaire2:=TV.AsString[TV.ColIndex('ATB_AFFAIRE2'),TV.CurrentRow];
if (TV.ColIndex('ATB_AFFAIRE3')>-1) then Affaire3:=TV.AsString[TV.ColIndex('ATB_AFFAIRE3'),TV.CurrentRow];
if (TV.ColIndex('ATB_AVENANT')>-1) then Avenant:=TV.AsString[TV.ColIndex('ATB_AVENANT'),TV.CurrentRow];
if (TV.ColIndex('ATB_TIERS')>-1) then Tiers := TV.AsString[TV.ColIndex('ATB_TIERS'),TV.CurrentRow];
if Affaire1 <> '' then
  BEGIN
  Affaire:=CodeAffaireRegroupe(Affaire0, Affaire1  ,Affaire2 ,Affaire3,Avenant, taModif, false,false,False);
  if not ExisteAffaire (Affaire,'') then
      BEGIN
      if TeststCleAffaire(Affaire,Affaire0,Affaire1,Affaire2,Affaire3,Avenant,Tiers,False,False,False,True) <> 1 then
          Affaire := '';
      END;
  END;
end;

// Gestion des écarts
procedure TOF_TBAffViewer.RecupFormuleEcart(valeur ,zone : string);
Var Formule : THEdit;
BEGIN
Formule := THEdit(GetControl (zone));
if Formule = Nil then Exit;
If (Valeur = 'ATB_PREQTEUNITERE') or (Valeur='ATB_FACQTEUNITERE') then valeur:=valeur+'F'; //zone sur 18C, ne tient pas dans abrégé ... 
Formule.SelText:='['+valeur+']';
END;

procedure TOF_TBAffViewer.AlimEcartTB;
Var CC   : TCheckBox;
BEGIN
CC := TCheckBox(GetControl('ECART1'));
if CC.Checked then ParamTB.MajEcart (1, GetControlText('FORMULEECART1'));
CC := TCheckBox(GetControl('ECART2'));
if CC.Checked then ParamTB.MajEcart (2, GetControlText('FORMULEECART2'));
CC := TCheckBox(GetControl('ECART3'));
if CC.Checked then ParamTB.MajEcart (3, GetControlText('FORMULEECART3'));
CC := TCheckBox(GetControl('ECART4'));
if CC.Checked then ParamTB.MajEcart (4, GetControlText('FORMULEECART4'));
CC := TCheckBox(GetControl('ECART5'));
if CC.Checked then ParamTB.MajEcart (5, GetControlText('FORMULEECART5'));
END;




///////////// FIN TOF TB VIEWER /////////////////////////////

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 03/10/2000
Modifié le ... : 03/10/2000
Description .. : paramétrage de la valorisation par défaut des états type
Suite ........ : tableau de bord / paramétrage du dossier
Mots clefs ... : ACTIVITE;VALORISATION ACTIVITE
*****************************************************************}
Function GetValoActDefaut (LesDeuxPossible : Boolean) : string;
BEGIN
Result := 'PR';
if (VH_GC.AFTypevaloAct = '001') and (LesDeuxPossible) then result := 'PRV' else
if (VH_GC.AFTypevaloAct = '002') or (VH_GC.AFTypevaloAct = '004') then result := 'PR' else
if (VH_GC.AFTypevaloAct = '003') or (VH_GC.AFTypevaloAct = '005') then result := 'PV';
END;

// *********  Fonctions publiées pour le script de l'AGL ************************
// appel de l'affaire associée au TB
procedure AGLLanceAffaireTB( parms: array of variant; nb: integer );
var  F : TForm;
     TOFenc : TOF;
begin
TOFenc := Nil;
F:=TForm(Longint(Parms[0]));
if (F is TFStat) then TOFenc := TFStat(F).Latof;
if (TOFenc is TOF_TBAffViewer) then TOF_TBAffViewer(TOFenc).LanceAffaireTB;
end;

// valorisation par defaut des Tb et editions en fonction du paramétrage du dossier
Function AGLGetvaloActDefaut( parms: array of variant; nb: integer ): variant;
begin
GetValoActDefaut (False);
end;



Procedure TOF_TBAffViewer.MenuPopOnCLick(Sender:Tobject);
begin
RecupFormuleEcart(TmenuItem(Sender).name,'FORMULEECART'+Copy(TPopupmenu(Tmenuitem(sender).owner).name,15,1));
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 13/06/2003
Modifié le ... :   /  /    
Description .. : Fct qui permet d'alimenter la table AfTableauBord
Suite ........ : en fct d'un source
Mots clefs ... : TB;
*****************************************************************}
Function  AfBuildAfTableauBord (Zone : T_PARAMTB):boolean;
Var  ParamTB    :TParamTBAffaire;
 SQL, ListeChamp : string;
 vargroupeConf, critere : String;
begin
     Result := FALSE;
     ParamTb := Nil;
     try
        // Paramètres de sélection des enregistrements du tableau de bord
        ParamTB := TParamTBAffaire.Create (ttbAffaire, Zone.Regroup, Zone.NiveauAna,Zone.unite, FALSE);
        if (GetParamsoc('SO_AFTYPECONF')='AGR') then
            begin
              vargroupeConf := GIGereCritGroupeConf;
              if vargroupeConf <> '' then
                begin
                critere := ReadTokenPipe(vargroupeConf,';ComboPlus;');
                Zone.GroupConf := critere;
                end;
            end;
        ParamTB.CutOffMois       := Zone.CutOffMOis;
        ParamTB.AffSurPeriode    := Zone.AffSurPeriode;
        ParamTB.AlimPrevu        := Zone.AlimPRevu;
        ParamTB.AlimSolde        := Zone.AlimSolde;
        ParamTB.AlimFacture      := Zone.AlimFacture;
        ParamTB.AlimReal         := Zone.AlimReal;
        ParamTB.AlimAchat        := Zone.AlimAchat;
        ParamTB.AlimValoRess_Art := False;
        ParamTB.AlimAN           := Zone.AlimAn;
        ParamTB.AlimVisa         := Zone.AlimVisa;
        ParamTB.AlimMois         := Zone.AlimMois;
        ParamTB.AlimCutOff       := Zone.AlimCutOff;
        ParamTB.AlimBoni         := Zone.AlimBoni;
        ParamTB.GereFacturable   := False;
        ParamTB.CentrFam         := FALSE;
        ParamTB.CentrStatRes     := FALSE;
        ParamTB.CentrStatArt     := FALSE;
        ParamTB.MtActivite       := Zone.Valoris;
        ParamTB.MajSelect (Zone);
        ParamTB.DetailParAffaire := False;
        ListeChamp :=''; SQL :='';
        ParamTB.AlimTableauBordAffaire ( ListeChamp,SQL, Zone.PieceAchat, Zone.PieceAchatEng, FALSE, FALSE, FALSE, FALSE, FALSE, Zone.AvecSsAff);
        Result := TRUE;
     except
           Result := FALSE;
     end;
 ParamTB.Free;
 ParamTB := nil;
end;


Procedure AFLanceFiche_Balance(Argument:String);
begin
AGLLanceFiche ('AFF','AFBALANCE','','',Argument);
end;
Procedure AFLanceFiche_GrandLivre(Argument:String);
begin
AGLLanceFiche ('AFF','AFGRANDLIVRE','','',Argument);
end;
Procedure AFLanceFiche_TableauBord;
begin
AGLLanceFiche ('AFF','AFTABLEAUBORD','','','');
end;
Procedure AFLanceFiche_TBAffaire(Argument:String);
begin
AGLLanceFiche ('AFF','AFTBVIEWER','','',argument);
end;
Procedure AFLanceFiche_TBMultiAff(Argument:String);
begin
AGLLanceFiche ('AFF','AFTBVIEWERMULTI','','',Argument);
end;
Procedure AFLanceFiche_EtatAppreciation(Argument:String);
begin
AGLLanceFiche ('AFF','AFAPPREC_ETAT','','',Argument);
end;
Procedure AFLanceFiche_EtatContrProd(Argument:String);
begin
AGLLanceFiche ('AFF','AFETAT_CTRL_PROD','','',Argument);
end;


(*Procedure ComboPlusAchat (Var Pluscombo: string);
Begin
     // ATTENTION, si modif voir aussi UtofAfBasePiece_Mul où il y a les mêmes choses
PlusCombo:=PlusCombo+' AND (';
  // ATtention si modif, voir aussi dans UtofAfTableauBord où ces test sont répété
PlusCombo:=PlusCombo+' (GPP_NATUREPIECEG="FF")';

// If  ctxTempo  in V_PGI.PGIContexte  then  begin
If  not(ctxScot  in V_PGI.PGIContexte)  then  begin
  PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="BLF")';
  PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="CF")';
  PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="BLC")';//mcd 05/05/03 fiche 10130
  END;
PlusCombo:=PlusCombo+' )';
end;  *)

procedure TOF_AFTBMENU.OnArgument(stArgument : String );
Begin
Inherited;
If Not (Bmulti) then SetControlVisible ('TB',False);
If Not GetParamsoc('So_AfAppAvecBM') then SetControlVisible ('BON',False); // pas de consultation boni/mali
SetControlVisible ('ENG',False); // pas de consultation engagé au 04/2003
if not(GetParamSoc('SO_AFAPPCUTOFF')) then SetControlVisible ('CUT',False);
if VH_GC.GAPlanningSeria =False then SetControlVisible ('PLA',False);
if Not (ctxscot in V_PGI.PGIContexte) then SetControlVisible ('APP',False);
{$ifdef DIFUSGI}
SetControlVisible ('BUD',False); // pas de budget en GI pour l'instant
SetControlVisible ('PLA',False); // pas de planning pour l'instant
{$ENDIF}

end;

procedure AGLImpGl( parms: array of variant; nb: integer );
var
Affaire, Tiers:string;
begin
Affaire := Parms[0]; Tiers := Parms[1];
AFLanceFiche_grandLivre('NOFILTRE:NOFILTRE;TYPETB:AFF;AFFAIRE:'+Affaire+';TIERS:'+Tiers);
End;


Initialization
registerclasses([Tof_TBAffaire,TOF_TBAffViewer,TOF_AFTBMENU]);
RegisterAglProc( 'LanceAffaireTB', True ,1, AGLLanceAffaireTB);
RegisterAglProc( 'ImpGl', False ,2, AGLImpGl);
RegisterAglfunc( 'GetValoActDefaut', False ,0, AGLGetvaloActDefaut);
end.
