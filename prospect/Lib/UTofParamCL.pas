unit UTofParamCL;

interface
                    
uses HSysMenu, HTB97,HEnt1, hmsgbox, Grids, ExtCtrls, EntRT,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}FE_Main,
{$ENDIF}
     UtilGC,
     Hctrls, Variants,
     StdCtrls, Forms, Controls,
     UTOF, Classes,Graphics,SysUtils,comctrls,Vierge,M3FP,
     UTOB,AGLInit,Spin,ParamSoc,TiersUtil,UtilRT,
     UtilConfid,UTOFPARAMONGLETS,HRichOle,UtilPGI;

Const
   { pour pouvoir gérer les particularité propre à chaque fiche
    fiche Client/prospect à part }
   CodeFichierProspect : String = '0';
   CodeFichierAction : String = '1';
   CodeFichierOperation : String = '2';
   CodeFichierFournisseur : String = '3';
   CodeFichierArticle : String = '4';
   CodeFichierParc : String = '5';
   CodeFichierContact : String = '6';
   CodeFichierIntervention : String = '7';
   CodeFichierLigne : String = '8';
   CodeFichierElemParc : String = '9';
   CodeFichierArticleParc : String = 'B';
   CodeFichierPiece : String = 'D';
   CodeFichierProjet : String = 'Q';
   CodeFichierProposition : String = 'V';
   TabletteProspect : String = 'RTLIBCHAMPSLIBRES';
   TabletteAutres : String = 'RTLIBCHAMPS';
   TabletteTypeSSContact : String = 'RTTYPECL_OPE';
   PrefixeProspect : String = 'RPR' ;
   NomTabletteCombo : String = 'RTRPRLIBTABLE';
   NomTabletteComboMul : String = 'RTRPRLIBTABMUL';
   ParamSocLargeurFiche : String = 'SO_RTLARGEURFICHE' ;
   ParamSocHauteurFiche : String = 'SO_RTHAUTEURFICHE' ;
   ParamSocFichier : String = 'SO_RTGESTINFOS00' ;
   RequeteRessource : String ='Select ARS_LIBELLE from RESSOURCE where ARS_RESSOURCE="';
   RequeteOperation : String ='Select ROP_LIBELLE from OPERATIONS where ROP_OPERATION="';
   RequeteProjet : String = 'Select RPJ_LIBELLE from PROJETS where RPJ_PROJET="';
   NbLignesMini : Integer = 2;
   BLargeur : Integer = 28;
   BHauteur : Integer = 27;
   NbElemListe : Integer = 26;

   NbOngletMax: integer =20;
   PosSeparateur: integer =500;  // pour les lignes et espaces, la numérotation du combo démarre à 100
   PosLigne: integer =0;
   PosEspace: integer =1;
   HauteurLigneSeule : integer = 6;
   HauteurPanel : integer = 30;
   NbLigneMulti : integer = 3;
   HauteurLigneListe : integer = 15;
   LargeurTexte : integer = 100;
   LargeurContact : integer = 40;
   ChampsNonTraitesRDA : String = '_DESC;_CLEDATA;_DATECREATION;_CREATEUR;_UTILISATEUR;_DATEMODIF';
   ChampsProprietes : String = 'INTITULE;OBLIGATOIRE;NBLIGNES;CRITERESEL;TYPECHAMP;CHOIXFILTRE;NOMCHAMP;CHPASSOCIE';
type
     TOF_ParamChamp = Class (TOF)
     Private
         MajParamChamp : Boolean;
         SavTob:Tob;
         CodeFichier,TypeChamp4 : String;
         procedure TypeChampClick(Sender: TObject);
         procedure ChargeListeChpOrigine;
         procedure AssocieClick(Sender: TObject);
         procedure CopieChampsClick(Sender: TObject);
         procedure ChoixClick(Sender: TObject);
     public
         procedure OnUpdate ; override ;
         procedure OnArgument(Arguments : string) ; override ;
         procedure OnLoad ; override ;
         procedure OnClose ; override ;
     END;

     TOF_ParamMulti = Class (TOF)
     public
         MC: TListBox;
         bValide : boolean;
         bValider : TToolBarButton97 ;
         ValChampOri : string;
         procedure OnArgument(Arguments : string) ; override ;
         procedure OnClose ; override ;
     Private
         procedure bValiderClick(Sender: TObject);
         procedure ListeExit(Sender: TObject);
         procedure ListeClick(Sender: TObject);
     END;

  TOF_ParamCL = Class (TOF)
  private    { Déclarations privées }
    LBChamps: THGrid;
    TobInfos,TobLibres,TobDesc,TobType,TobAction,TobProposition,TobTiers,TobEcran,SavTobEcran,TobArticle,TobLigne,TobPiece : TOB;
    TobArticleParc,TobParc,TobParcNome,TobPropriete : Tob ;
    CurLig            : LongInt;
    ParamChamps,WidthEcran,HeightEcran,PremierOnglet : Integer;
    NumSeparateur : Integer;
    CodeFichier,TabletteLibelle,Prefixe,TabletteCombo,TabletteComboMul : String;
    ParamSocLargeur,ParamSocHauteur : String;

    Nbpost,MaxCombo,MaxMultiChoix,MaxTexte,MaxValeur,MaxBool,MaxDate : Integer ;

    OrigineFichier : boolean; // true lorsque l'on vient d'une fiche de données
    ExisteInfos : boolean; // true lorsque l'enreg infos complémentaire existe
    OrigineChampOblig : boolean; // true lorsque la fiche est appelée en création et qu'existe un champ obligatoire
    OrigineParam : boolean; // true lorsqu'en paramétrage fiche, on teste l'affichage de la fiche
    CodeInfos,DernierPanel : String;
    FinBoucle : Boolean;
    VerrouModif,Creation : Boolean; // Confidentialité
    ModifLot : boolean;
    StSQL : string;
    // 1 : Nom du Champ, 2 : Type du Champ, 3 : Valeur paramétrée pour la condition d'affichage
    //     de l'onglet, 4: Intitulé Onglet
    // 10 = NbOngletMax
    InfosOnglets: array[1..20,1..4] of string;
    TypeAction : String ; // appel fiche depuis le paramétrage des types d'action
    rac_TypeAction : String ;
    PrefixeAutres : String;
    ErreurAppel,BougePropriete,BougeOnglet,BougeChamp : Boolean;
    MajParamChamp : Boolean;
    PremierMul: Integer;
    CouleurPanel : TColor;
    bLignes:boolean; // trt particulier pour les lignes ou l'on traite en TOB 
    procedure ChargeListe ;
    procedure ChargeParam ;
    procedure CreerOnglet (NoOnglet : integer);
    procedure CreerPanel  (OngletActif:integer; Curlig : integer; TC : string; NewPos : Boolean; Libelle : string; NonGrise : Boolean);
    procedure AjouterChamps;
    procedure EnleverChamps;
    procedure MonterChamps;
    procedure DescendreChamps;
    procedure AjouterOnglet;
    procedure OngletAGauche;
    procedure OngletADroite;
    procedure EnleverOnglet;
    procedure FicheSaisie;
    procedure FicheParam;
    procedure Panels_OnEnter(Sender: TObject);
    procedure NewPosition;
    procedure MajParametrage;
    procedure MajInfosCompl ;
    procedure InitForme ;
    procedure Propriete ;
    procedure InfoOblig ( LibChamps:string ; Onglet:integer ; Panel:integer );
    procedure bOngletClick(Sender: TObject);
    //procedure cFichierClick(Sender: TObject);
    procedure TestAffichageOnglets;
    procedure btestficheClick(Sender: TObject);
    procedure bDeleteClick(Sender: TObject);
    procedure bValDefClick(Sender: TObject);
// cd 221100
    procedure SetArgumentsCL(StSQL : string);
    procedure SelectionMulti(Sender: TObject);

    procedure TexteClick(Sender: TObject);
    procedure TestExiste(Sender: TObject);
    procedure BoutonsParamInvisibles;
    procedure ChampExit (Sender: TObject);
    procedure TaillePanelMulti(P : TPanel; TL : TListBox; iPosit : integer);
    Procedure CreerChampTob (NomChamp : String);
    procedure ValidModifs;
    Procedure TestInfoOnglet(i : integer; TobInfos : Tob; Var PremierOnglet : integer; NomChamp : String;TS: TTabSheet);    
    function ExisteOngletChampOblig : boolean ;
  public
    { Déclarations publiques }
         //procedure OnLoad ; override ;
         procedure OnClose ; override ;
         procedure OnUpdate ; override ;
         procedure OnLoad ; override ;

         procedure OnArgument (Arguments : string) ; override ;
         procedure updateTablette (stLibelle, stAbrege, stType, stCode : string);
  end;


procedure ChargeListBox (TabletteComboMul : string; MC:TListBox; Complet : Boolean ; CleChoixCode : String; Valchamp : String; ProspectExiste : boolean);
procedure ChargeChampMulti (C:TListBox ; var Valeur: String; LibComplet : boolean; Tous : boolean);
Function RTLanceFiche_ParamCl(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
Function RTCompareTobs(Tob1,tob2 : tob) : boolean;
Procedure AddChpsProprietes(TobP : Tob);

implementation
uses 
   CbpMCD
   ,CbpEnumerator;

Function RTLanceFiche_ParamCl(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
  VH_RT.TobChampsPro.Load;
  VH_RT.TobTypesAction.Load;

  AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure Tof_ParamChamp.OnArgument(Arguments : string) ;
var Critere,Critere2 : string;
    TypeChamp: THValComboBox;
    Choix: THEdit;
begin
inherited ;
  Critere:=uppercase(Trim(ReadTokenSt(Arguments)));
  if ( Critere <> 'MULTI' ) and ( Critere <> 'BLOB' ) then
  begin
     SetControlVisible('NBLIGNES',FALSE) ;
     SetControlVisible('LTAILLE',FALSE) ;
  end;

  if Critere = 'BOOLEAN' then
   SetControlVisible('OBLIGATOIRE',FALSE) ;

  if Critere='NON' then
   BEGIN
   SetControlVisible('OBLIGATOIRE',FALSE) ;
   SetControlVisible('CRITERESEL',FALSE) ;
   SetControlVisible('SUITECRITERESEL',FALSE) ;
   END;

  if copy(critere,1,4) <> 'EDIT' then
   begin
   SetControlVisible('TYPECHAMP',FALSE) ;
   SetControlVisible('LTYPECHAMP',FALSE) ;
   SetControlVisible('CHOIXFILTRE',FALSE) ;
   SetControlVisible('LCHOIXFILTRE',FALSE) ;
   end
  else
   begin
   TypeChamp:= THValComboBox(GetControl('TYPECHAMP'));
   TypeChamp.OnClick := TypeChampClick;
   Choix:= THEdit(GetControl('CHOIXFILTRE'));
   Choix.OnElipsisClick := ChoixClick;
   end;


  Critere2:=ReadTokenSt(Arguments);
  if (Critere = 'BLOB') then
    begin
    SetControlVisible('CRITERESEL',FALSE) ;
    SetControlVisible('SUITECRITERESEL',FALSE) ;
    SetControlProperty ('NBLIGNES','MaxValue',20);
    end;
  if ( Critere2 = CodeFichierOperation ) or ( Critere2 = CodeFichierArticle ) or
     ( Critere2 = CodeFichierArticleParc ) or ( Critere2 = CodeFichierElemParc ) or
     ( Critere2 = CodeFichierProjet) then
    SetControlProperty ('TYPECHAMP','DataType',TabletteTypeSSContact);

  TypeChamp4:=Critere;
  CodeFichier:=Critere2;
  if (CodeFichier = CodeFichierLigne) and (GetParamSocSecur('SO_RTGESTINFOS004',False) = true ) then
     begin
     SetControlVisible('LCHPASSOCIE',true) ;
     SetControlVisible('CHPASSOCIE',true) ;
     ChargeListeChpOrigine;
     THValComboBox(GetControl('CHPASSOCIE')).OnClick:=AssocieClick;
     if (TypeChamp4 = 'COMBO') or (TypeChamp4 = 'MULTI') then
       begin
       SetControlVisible('COPIECHAMPS',true) ;
       THValComboBox(GetControl('COPIECHAMPS')).OnClick:=CopieChampsClick;
       end
     end;
 {$ifdef GIGI}  //mcd 04/07/2005
  Critere2:=ReadTokenSt(Arguments); //3eme argument pour la GI
  if critere2='DP' then
    begin
    SetControlVisible('OBLIGATOIRE',FALSE) ;
    SetControlChecked('CRITERESEL',true) ;
    end;
  if critere2='CAT' then
    begin  //mcd 19/07/2006
    SetControlVisible('TypeChamp',FALSE) ;
    SetControlVisible('LTypeChamp',FALSE) ;
    end;
 {$ENDIF GIGI}
end;

procedure Tof_ParamChamp.OnUpdate ;
begin
inherited ;
MajParamChamp:=True;
TheTob:=Latob;
end;

procedure Tof_ParamChamp.OnClose ;
var NowTob : Tob;
begin
inherited ;
NowTob:=Latob;
NowTob.GetEcran(TFVierge(ecran),Nil);
if (MajParamChamp=False) and ( RTCompareTobs(SavTob,NowTob)=False ) then
    begin
    Case PGIAskCancel('Voulez-vous enregistrer les modifications ?',TFVierge(ecran).Caption) of
      mrNo :
         TheTob:=Nil;
      mrCancel :
          begin
          LastError:=1;
          exit;
          end;
      mrYes : TheTob:=Latob;
    end ;
    end;
if Assigned(SavTob) then
    FreeAndNil(SavTob);
end;

procedure Tof_ParamChamp.OnLoad ;
var TypeChamp: THValComboBox;
begin
inherited ;
{ mng : les champs de la fiche sont automatiquement alimentés via TheTob }
TypeChampClick(Ecran);
SavTob:=tob.create('Sauv Propriétés',Nil,-1);
AddChpsProprietes(SavTob);
// SavTob contient un champ de moins, ce doit être la première tob de RTCompareTobs
SavTob.DelChampSup('NOMCHAMP',False);
SavTob.GetEcran(TFVierge(ecran),Nil);
  TypeChamp:= THValComboBox(GetControl('TYPECHAMP'));
  TypeChamp.OnClick := TypeChampClick;
  if ( CodeFichier = CodeFichierFournisseur ) then
    begin
    TypeChamp.Plus:=' AND CO_CODE <>"P  "';
    SetControlText ('NATUREAUXI','FOU');
    end;
  { en multisoc on enleve ressource, operations et projets }
  if EstBaseMultiSoc then
     TypeChamp.Plus:=' AND CO_CODE <>"R  " AND CO_CODE <>"O  " AND CO_CODE <>"P  "'
{$IFDEF GRCLIGHT}
  else
    if not GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False) then
       TypeChamp.Plus:=' AND CO_CODE <>"O  " AND CO_CODE <>"P  "'
{$ENDIF GRCLIGHT}
  ;
end;

procedure Tof_ParamChamp.TypeChampClick(Sender: TObject);
begin
if (GetControlText('TYPECHAMP')='') or (GetControlText('TYPECHAMP')='C') then
    begin
    SetControlVisible('CHOIXFILTRE',FALSE) ;
    SetControlVisible('LCHOIXFILTRE',FALSE) ;
    end
else
    begin
    SetControlVisible('CHOIXFILTRE',True) ;
    SetControlVisible('LCHOIXFILTRE',True) ;
    end
end;

procedure Tof_ParamChamp.ChoixClick(Sender: TObject);
var TypeChamp: THValComboBox;
    ChoixFiltre : String;
begin
TypeChamp:= THValComboBox(GetControl('TYPECHAMP'));
if TypeChamp.Value = 'T' then
   begin
     if GetControlText ('NATUREAUXI') = 'FOU' then
        ChoixFiltre:=AGLLanceFiche ('RT', 'RFTIERS_TL','' , '', 'PARAM')
     else
        ChoixFiltre:=AGLLanceFiche ('RT', 'RTTIERS_TL','' , '', 'PARAM');
   end;
if TypeChamp.Value = 'R' then
    ChoixFiltre:=AGLLanceFiche ('RT', 'RTRESSOURCE_TL','' , '', 'PARAM');

{$IFDEF GRCLIGHT}
    if GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False) then
      begin
{$ENDIF GRCLIGHT}
      if TypeChamp.Value = 'O' then
          if CodeFichier = CodeFichierFournisseur then
            ChoixFiltre:=AGLLanceFiche ('RT', 'RFOPERATIONS_TL','' , '', 'PARAM')
          else
            ChoixFiltre:=AGLLanceFiche ('RT', 'RTOPERATIONS_TL','' , '', 'PARAM');
      if TypeChamp.Value = 'P' then
          ChoixFiltre:=AGLLanceFiche ('RT', 'RTPROJETS_TL','' , '', 'PARAM');
{$IFDEF GRCLIGHT}
      end;
{$ENDIF GRCLIGHT}
SetControlText ('CHOIXFILTRE',ChoixFiltre);
end;

procedure Tof_ParamChamp.ChargeListeChpOrigine;
var C:THValComboBox;
    iTable,iChamp : integer;
    ListeChamp,ListeLibelle : HTStringList;
    stTypeChamp : String;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;

begin
MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();
	//
  stTypeChamp:=TypeChamp4;
  if stTypeChamp = 'MULTI' then stTypeChamp:='VARCHAR(80)'
     else if stTypeChamp = 'EDIT' then stTypeChamp:='VARCHAR(35)';
  C:=THValComboBox(getcontrol('CHPASSOCIE')) ;
  ListeChamp := HTStringList.Create;
  ListeLibelle := HTStringList.Create;
  ListeChamp.add (''); ListeLibelle.add (TraduireMemoire('<<Aucun>>'));

  table := Mcd.getTable(mcd.PrefixetoTable('RD4'));
  FieldList := Table.Fields;
  FieldList.Reset();
  While FieldList.MoveNext do
     if ( (FieldList.Current as IFieldCOM).libelle[1] <> '.') and
        ( (copy((FieldList.Current as IFieldCOM).name,5,3) = 'RD4') or
          (copy((FieldList.Current as IFieldCOM).name,5,8) = 'BLOCNOTE') ) and
        ( (FieldList.Current as IFieldCOM).Tipe=stTypeChamp) then
        begin
        ListeChamp.add((FieldList.Current as IFieldCOM).name);
        ListeLibelle.add ((FieldList.Current as IFieldCOM).libelle);
        end;
  C.Items.Assign(ListeLibelle);
  C.Values.Assign(ListeChamp);
  if LaTob.GetString('CHPASSOCIE') <> '' then
     C.Value:=LaTob.GetString('CHPASSOCIE');
  ListeChamp.free;ListeLibelle.Free;
end;

procedure Tof_ParamChamp.AssocieClick(Sender: TObject);
var C:THValComboBox;
begin
  C:=THValComboBox(getcontrol('CHPASSOCIE')) ;
  if C.ItemIndex <> 0 then
     THEdit(GetControl('INTITULE')).Text := C.Items[C.ItemIndex];
end;

procedure Tof_ParamChamp.CopieChampsClick(Sender: TObject);
var C:THValComboBox;
    stCode4,stCode8 : String;
    TobLib004,TobLib008 : tob;
    i,j : integer;
    Q : TQuery;
    N:THLabel;
begin
  C:=THValComboBox(getcontrol('CHPASSOCIE')) ;
  if C.Value = '' then exit;
  N:=THLabel(getcontrol('NOMCHAMP')) ;
  if N.Caption = '' then exit;
  i:=Pos(')',N.Caption);
  if i > 2 then
    begin
    if (TypeChamp4 = 'MULTI') then
      stCode8:='R8'+Chr(ValeurI(copy(N.Caption,length(N.Caption)-2,1))+65)
    else
      stCode8:='R8'+copy(N.Caption,length(N.Caption)-2,1);
    end;

  if PGIAsk('Voulez-vous vraiment récupérer les tablettes Articles ?','')=mrNo then exit;

  if (TypeChamp4 = 'MULTI') then
    stCode4:='R4'+Chr(ValeurI(copy(C.Value,length(C.Value),1))+65)
  else
    stCode4:='R4'+copy(C.Value,length(C.Value),1);

  TobLib004:=TOB.create ('CHOIXCOD',NIL,-1);
  Q:=OpenSQL('Select * from CHOIXCOD where CC_TYPE="'+stCode4+'"',True);
  TobLib004.LoadDetailDB ('CHOIXCOD', '', '', Q, False);
  if TobLib004.Detail.Count = 0 then
    begin
    Ferme(Q);
    if Assigned(TobLib004) then
      FreeAndNil(TobLib004);
    exit;
    end;
  Ferme(Q);
  TobLib008:=TOB.create ('CHOIXCOD',NIL,-1);
  for j :=0 to TobLib004.Detail.Count-1  do
    begin
      for i :=1 to TobLib004.Detail[j].NbChamps do
         if TobLib004.Detail[j].GetNomChamp(i) <> 'CC_TYPE' then
          TobLib008.PutValue(TobLib004.Detail[j].GetNomChamp(i),TobLib004.Detail[j].GetValue(TobLib004.Detail[j].GetNomChamp(i)));
    TobLib008.PutValue('CC_TYPE','R8'+copy(stCode8,3,1));
    TobLib008.InsertOrUpdateDB(False);
    end;
    if Assigned(TobLib004) then
      FreeAndNil(TobLib004);
    if Assigned(TobLib008) then
      FreeAndNil(TobLib008);
  AvertirTable ('RT008LIBTABMUL'+copy(N.Caption,length(N.Caption)-2,1));
end;

procedure Tof_ParamMulti.OnArgument(Arguments : string) ;
var Critere,NomCode,ChampMul,ValMul,NomTablette : string;
    ProspectExiste : boolean;
    x : Integer;
begin
inherited ;
bValide := false;
bValider := TToolBarButton97(GetControl('bValider'));
bValider.OnClick := bValiderClick;
ProspectExiste:=False;
Repeat
  Critere:=ReadTokenPipe(Arguments,'|') ;
  if Critere<>'' then
      begin
      x:=pos('=',Critere);
      if x<>0 then
         begin
         ChampMul:=copy(Critere,1,x-1);
         ValMul:=copy(Critere,x+1,length(Critere));
         if ChampMul='VALCHAMP' then ValChampOri := ValMul
         else if ChampMul='NOMCODE' then NomCode := ValMul
         else if ChampMul='TABLETTEMULTI' then NomTablette := ValMul
         else if ChampMul='PROSPECTEXISTE' then
              if ValMul = '0' then ProspectExiste:=False else ProspectExiste:=True
         ;
         end;
      end;
until (Critere = '') ;

MC:= TListBox(GetControl('LISTEMULTI'));
ChargeListBox (NomTablette,MC,True,NomCode,ValChampOri,ProspectExiste);
MC.OnExit:=ListeExit;
MC.OnClick:=ListeClick;
end;


procedure Tof_ParamMulti.OnClose;
Var Valeur : string;
begin
inherited ;
if bValide then
   begin
   ChargeChampMulti (TListBox(MC),Valeur,True,False);
   if (ecran <> nil) then TFVierge(ecran).retour:=Valeur;
   end
   else
      if (ecran <> nil) then TFVierge(ecran).retour:=ValChampOri;
end;

procedure Tof_ParamMulti.bValiderClick(Sender: TObject);
begin
bValide := true;
end;

procedure Tof_ParamCL.bOngletClick(Sender: TObject);
var TS: TTabSheet;
    i,j : integer;
    ret,critere:string;
begin
i:=TPageControl(getControl( 'PAGEONGLETS')).ActivePage.PageIndex+1;
TS:= TTabSheet(Ecran.FindComponent('ONGLETLIBRE'+IntToStr(i)));
if TS=Nil then exit;
Critere:='';
for j:=1 to 4 do
  Critere:=Critere+'|'+InfosOnglets[i,j];
Critere:=Critere+'|'+CodeFichier;
ret:=RTLanceFiche_PARAMONGLETS('RT','RTPARAMONGLETS','','',IntToStr(i)+Critere) ;

if ret<>'' then
   begin
   BougeOnglet:=true;
   for j:=1 to 4 do
      begin
      Critere:=ReadTokenPipe(ret,'|') ;
      InfosOnglets[i,j]:=Critere;
      end;
   end;
TS.Caption:= InfosOnglets[i,4];
end;

procedure Tof_ParamCL.btestficheClick(Sender: TObject);
var TobChampsProFille : tob;
begin
  if CodeFichier = '' then exit;
  TobChampsProFille:=VH_RT.TobChampsPro.FindFirst(['CO_CODE'], [CodeFichier], TRUE);
  if not Assigned(TobChampsProFille) or (TobChampsProFille.Detail.Count=0) then
    begin
    PGIInfo('Le paramétrage de cette saisie n''a pas été effectué');
    exit;
    end;

AGLLanceFiche('RT','RTPARAMCL','','','PARAMFICHE'+CodeFichier) ;
end;

procedure Tof_ParamCL.bDeleteClick(Sender: TObject);
begin
if PGIAsk('Voulez-vous vraiment supprimer ce paramétrage ?','')=mrNo then exit;
if CodeFichier=CodeFichierProspect then
    ExecuteSQL('DELETE FROM CHAMPSPRO') 
else
    ExecuteSQL('DELETE FROM RTINFOSDESC where RDE_DESC="'+CodeFichier+'"') ;
VH_RT.TobChampsPro.Load (true);
end;

procedure Tof_ParamCL.bValDefClick(Sender: TObject);
var TobDef : Tob;
    i,j : integer;
    FieldNameTo,FieldNameFrom,NomCode : String;
    MC: TListBox;

begin
TobDef:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[TypeAction,'---',0],TRUE);

//TobType:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[TypeAction,ReadToKenSt(CodeInfos),StrToInt(ReadToKenSt(CodeInfos))],TRUE) ;

if (TobDef <> Nil) and (TobType <> Nil) then
begin
   for i :=1 to TobDef.NbChamps do
   begin
      FieldNameFrom := TobDef.GetNomChamp(i);
      if (copy(FieldNameFrom,1,7) <> 'RPA_RPR') and
         (copy(FieldNameFrom,1,12) <> 'RPA_BLOCNOTE') then continue;
      FieldNameTo  := 'RPR_' + copy(FieldNameFrom,5,length(FieldNameFrom)) ;
      TobInfos.PutValue(FieldNameTo, TobDef.GetValue(FieldNameFrom));
      // recherche du champs et MAJ valeurs
      for j:= 0 to ( LBChamps.RowCount - 1 ) do
      begin
        if (LBChamps.Cells[3,j] = FieldNameTo) and (LBChamps.RowHeights[j]=0) then
        begin
          if LBChamps.Cells[1,j]<>'MULTI' then
            SetControlText(LBChamps.Cells[0 ,j],TobDef.GetValue(FieldNameFrom))
          else
          begin
           MC:= TListBox(GetControl(LBChamps.Cells[0 ,j]));
           MC.Clear;
           NomCode:='RM'+copy(LBChamps.Cells[0 ,j],3,1);
           ChargeListBox (TabletteComboMul, MC,( not OrigineFichier) ,NomCode,TobDef.GetValue(FieldNameFrom),true);
          end;
          break;
        end;
      end;
   end;
end;

end;

procedure Tof_ParamCL.TexteClick(Sender: TObject);
var Ctrl : TControl;
    Fiche : TFVierge ;
    Retour,StArg : String;
    Libelle : TLabel;
begin
Fiche := TFVierge(ecran);
Ctrl:=TControl(Fiche.FindComponent(LBChamps.Cells[0,Curlig]));
if Ctrl = Nil then exit;
Libelle:=TLabel(Fiche.FindComponent('LT'+IntToStr(Curlig)));

StArg:='';
if LBChamps.Cells[7,Curlig] = 'T' then
    begin
    if TEdit(Ctrl).Text <> '' then
       StArg:= 'T_TIERS='+ TEdit(Ctrl).Text;
    if GetControlText('NATUREAUXI') = 'FOU' then
       Retour:=AGLLanceFiche ('RT', 'RFTIERS_TL',StArg , '', 'PROSPECT;FILTRE='+LBChamps.Cells[8,Curlig])
    else
       Retour:=AGLLanceFiche ('RT', 'RTTIERS_TL',StArg , '', 'PROSPECT;FILTRE='+LBChamps.Cells[8,Curlig]);
    end
else

if LBChamps.Cells[7,Curlig] = 'R' then
    begin
    if TEdit(Ctrl).Text <> '' then
       StArg:= 'ARS_RESSOURCE='+ TEdit(Ctrl).Text;
    Retour:=AGLLanceFiche ('RT', 'RTRESSOURCE_TL',StArg , '', 'PROSPECT;FILTRE='+LBChamps.Cells[8,Curlig])
    end
else
if LBChamps.Cells[7,Curlig] = 'C' then
    begin
    if (TEdit(Ctrl).Text <> '') and (Libelle <> Nil) then
       StArg:= 'FROMSAISIE='+Libelle.Caption+';';

    if CodeFichier = CodeFichierProspect then
       Retour:=AGLLanceFiche('YY','YYCONTACT','T;'+CodeInfos,'',StArg+'PROSPECT')
    else
       if CodeFichier = CodeFichierAction then
          Retour:=AGLLanceFiche('YY','YYCONTACT','T;'+copy(CodeInfos,1,(pos(';',CodeInfos)-1)),'',StArg+'PROSPECT')
       else
         if (CodeFichier = CodeFichierParc) and (Assigned(TobParc)) then
           begin
           if TobParc.GetString('WPC_TIERS') <> '' then
             Retour:=AGLLanceFiche('YY','YYCONTACT','T;'+TiersAuxiliaire(TobParc.GetString('WPC_TIERS')),'',StArg+'PROSPECT');
           end
         else
         if (CodeFichier = CodeFichierProposition) and (Assigned(TobProposition)) then
           begin
           if TobProposition.GetString('RPE_TIERS') <> '' then
             Retour:=AGLLanceFiche('YY','YYCONTACT','T;'+TiersAuxiliaire(TobProposition.GetString('RPE_TIERS')),'',StArg+'PROSPECT');
           end
         else
           if CodeFichier = CodeFichierPiece then
             Retour:=AGLLanceFiche('YY','YYCONTACT','T;'+TiersAuxiliaire(TobPiece.GetString('GP_TIERS')),'',StArg+'PROSPECT')
           else
             Retour:=AGLLanceFiche('YY','YYCONTACT',copy(CodeInfos,1,(pos(';',copy(CodeInfos,3,length(CodeInfos)))+2)),'',StArg+'PROSPECT');
    end
{$IFDEF GRCLIGHT}
    else
      if GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False) then
        begin
{$ELSE}
    else
{$ENDIF GRCLIGHT}
        if LBChamps.Cells[7,Curlig] = 'P' then
            begin
            if TEdit(Ctrl).Text <> '' then
               StArg:= 'RPJ_PROJET='+ TEdit(Ctrl).Text;
            Retour:=AGLLanceFiche ('RT', 'RTPROJETS_TL',StArg , '', 'PROSPECT;FILTRE='+LBChamps.Cells[8,Curlig])
            end
        else
        if LBChamps.Cells[7,Curlig] = 'O' then
            begin
            if TEdit(Ctrl).Text <> '' then
               StArg:= 'ROP_OPERATION='+ TEdit(Ctrl).Text;
            if CodeFichier = CodeFichierFournisseur then
              Retour:=AGLLanceFiche ('RT', 'RFOPERATIONS_TL',StArg , '', 'PROSPECT;FILTRE='+LBChamps.Cells[8,Curlig])
            else
              Retour:=AGLLanceFiche ('RT', 'RTOPERATIONS_TL',StArg , '', 'PROSPECT;FILTRE='+LBChamps.Cells[8,Curlig]);
            end
{$IFDEF GRCLIGHT}
        end
{$ENDIF GRCLIGHT}
;
if Retour <> '' then
   begin
   TEdit(Ctrl).Text:=ReadTokenSt(Retour);
   if Libelle <> Nil then
       TLabel(Libelle).Caption:=ReadTokenSt(Retour);
   end;
end;

procedure Tof_ParamCL.TestExiste(Sender: TObject);
var Ctrl : TControl;
    Fiche : TFVierge ;
    Requete : String;
    Libelle : TLabel;
    Q : TQuery;
begin
Fiche := TFVierge(ecran);
Ctrl:=TControl(Fiche.FindComponent(LBChamps.Cells[0,Curlig]));
if Ctrl = Nil then exit;
Libelle:=TLabel(Fiche.FindComponent('LT'+IntToStr(Curlig)));
if THEdit(Ctrl).text = '' then
   begin
   if Libelle <> Nil then
      TLabel(Libelle).Caption:='';
   exit;
   end;

if LBChamps.Cells[7,CurLig] = 'T' then
   begin
   if GetControlText ('NATUREAUXI')='FOU' then
     Requete:='Select T_LIBELLE from TIERS where T_TIERS="'+THEdit(Ctrl).text+
       '" and (T_NATUREAUXI = "FOU" )'
   else
{$ifdef GIGI}
     Requete:='Select T_LIBELLE from TIERS where T_TIERS="'+THEdit(Ctrl).text+
     '" and (T_NATUREAUXI = "CLI" or T_NATUREAUXI="PRO"  or T_NATUREAUXI="NCP" or T_NATUREAUXI="CON")';
{$else}
     Requete:='Select T_LIBELLE from TIERS where T_TIERS="'+THEdit(Ctrl).text+
     '" and (T_NATUREAUXI = "CLI" or T_NATUREAUXI="PRO" or T_NATUREAUXI="CON")';
{$endif}
   end
else
   if LBChamps.Cells[7,CurLig] = 'R' then
      Requete:=RequeteRessource+THEdit(Ctrl).text+'"'
   else
     if LBChamps.Cells[7,CurLig] = 'O' then
        Requete:=RequeteOperation+THEdit(Ctrl).text+'"'
     else
       if LBChamps.Cells[7,CurLig] = 'P' then
          Requete:=RequeteProjet+THEdit(Ctrl).text+'"'
       else
        if LBChamps.Cells[7,CurLig] = 'C' then
            begin
            if CodeFichier = CodeFichierProspect then
               Requete:='Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'+CodeInfos+'" AND C_NUMEROCONTACT='+THEdit(Ctrl).text
            else
               if CodeFichier = CodeFichierAction then
                  Requete:='Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'+copy(CodeInfos,1,(Pos(';',CodeInfos)-1))+'" AND C_NUMEROCONTACT='+THEdit(Ctrl).text
               else
               if CodeFichier = CodeFichierPiece then
                 Requete:='Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'+TiersAuxiliaire(TobPiece.GetString('GP_TIERS'),false)+'" AND C_NUMEROCONTACT='+THEdit(Ctrl).text
               else
               if CodeFichier = CodeFichierProposition then
                 Requete:='Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'+TiersAuxiliaire(TobProposition.GetString('RPE_TIERS'),false)+'" AND C_NUMEROCONTACT='+THEdit(Ctrl).text
               else
                 if (CodeFichier = CodeFichierParc) and (Assigned(TobParc)) then
                   begin
                   if TobParc.GetString('WPC_TIERS') <> '' then
                     Requete:='Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'+TiersAuxiliaire(TobParc.GetString('WPC_TIERS'),false)+'" AND C_NUMEROCONTACT='+THEdit(Ctrl).text;
                   end
                 else
                   Requete:='Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'+copy(CodeInfos,3,(pos(';',copy(CodeInfos,3,length(CodeInfos))))-1)+'" AND C_NUMEROCONTACT='+THEdit(Ctrl).text;
            end;
if Requete <> '' then
    begin
    Q:=OpenSQL(Requete,True);
    if Q.Eof then
       begin
       PGIBox(Format('Le code %s n''existe pas',[THEdit(Ctrl).Text]),'Code Inexistant');
       THEdit(Ctrl).SetFocus;
       end
    else
       begin
       if Libelle <> Nil then
           TLabel(Libelle).Caption:=Q.Fields[0].AsString;
       end;
    Ferme (Q);
    end;
end;

procedure Tof_ParamCL.TestAffichageOnglets;
var i,j,k,nbpanel : integer;
    TS: TTabSheet;
    Critere, NomChamp, CritOnglet,St,LibChamps : String;
    trouve : boolean;
    ListeActive: TScrollBox;
    P : TPanel;
    C : TControl;
begin
PremierOnglet:=0;
for i:=1 to TPageControl(getcontrol('PAGEONGLETS')).PageCount do
begin
  TS:= TTabSheet(Ecran.FindComponent('ONGLETLIBRE'+IntToStr(i)));
  if TS=Nil then exit;
  TS.TabVisible:=True;
  NomChamp:=InfosOnglets[i,1] ;
  if copy(InfosOnglets[i,1],1,3) = PrefixeAutres then
      //NomChamp:='RPR'+ copy(InfosOnglets[i,1],4,length(InfosOnglets[i,1])) ;
      NomChamp:='RPR_RPR'+ copy(InfosOnglets[i,1],8,length(InfosOnglets[i,1])) ;
  if InfosOnglets[i,2] = '' then
     if PremierOnglet=0 then PremierOnglet:=i;
  if TypeAction = '' then
  begin
    if {(ExisteInfos) and} (copy(InfosOnglets[i,1],1,3) = 'RPR') then
      TestInfoOnglet(i,TobInfos,PremierOnglet,NomChamp,TS);

    if (TobLibres <> Nil) and (copy(InfosOnglets[i,1],1,1) = 'Y') then
        begin
        if (InfosOnglets[i,2] = 'COMBO') or (InfosOnglets[i,2] = 'VARCHAR(6)') then
           begin
           trouve:=false;
           CritOnglet:= InfosOnglets[i,3];
           Repeat
            Critere:=ReadTokenSt(CritOnglet) ;
            if Critere<>'' then
               begin
               if TobLibres.GetValue(NomChamp) = Critere then
                  begin
                  trouve:=true;
                  if PremierOnglet=0 then PremierOnglet:=i;
                  break;
                  end;
               end;
           until (Critere = '') ;
           if not trouve then
                  // on ne supprime pas l'Onglet mais on le rend invisible
                  TS.TabVisible:=False;
           end;
        end;
    if (TobTiers <> Nil) and (copy(InfosOnglets[i,1],1,1) = 'T') then
      TestInfoOnglet(i,TobTiers,PremierOnglet,NomChamp,TS)
    else
      if (TobAction <> Nil) and (copy(InfosOnglets[i,1],1,3) = 'RAC') then
         TestInfoOnglet(i,TobAction,PremierOnglet,NomChamp,TS)
      else
        if (TobArticle <> Nil) and (copy(InfosOnglets[i,1],1,2) = 'GA') then
            TestInfoOnglet(i,TobArticle,PremierOnglet,NomChamp,TS)
        else
          if (TobArticleParc <> Nil) and (copy(InfosOnglets[i,1],1,3) = 'WAP') then
              TestInfoOnglet(i,TobArticleParc,PremierOnglet,NomChamp,TS)
          else
            if (TobParc <> Nil) and (copy(InfosOnglets[i,1],1,3) = 'WPC') then
                TestInfoOnglet(i,TobParc,PremierOnglet,NomChamp,TS)
            else
              if (TobParcNome <> Nil) and (copy(InfosOnglets[i,1],1,3) = 'WPN') then
                  TestInfoOnglet(i,TobParcNome,PremierOnglet,NomChamp,TS)
                else
                  if (TobLigne <> Nil) and (copy(InfosOnglets[i,1],1,2) = 'GL') then
                    TestInfoOnglet(i,TobLigne,PremierOnglet,NomChamp,TS)
                  else
                    if (TobPiece <> Nil) and (copy(InfosOnglets[i,1],1,2) = 'GP') then
                      TestInfoOnglet(i,TobPiece,PremierOnglet,NomChamp,TS)
                    else
                      if (TobProposition <> Nil) and (copy(InfosOnglets[i,1],1,3) = 'RPE') then
                        TestInfoOnglet(i,TobProposition,PremierOnglet,NomChamp,TS);

  ;

    // Valeur par défaut enlevées si Onglet rendu invisible
    if (OrigineFichier) and (not ExisteInfos) and ( TS.TabVisible = False ) then
//==================
       begin
       ListeActive:= TScrollBox(GetControl('SCB'+IntToStr(i)));
       For j := 0 to ListeActive.ControlCount - 1 do
          begin
          if ListeActive.Controls[j] is TPanel then
             begin
             P := TPanel(ListeActive.Controls[j]);
             St:= P.Name;
               if Pos('PCHAMPS_',St)<=0 then exit;
             CurLig:=StrToInt(Copy(St,pos('_',St)+1,Length(St)-pos('_',St)));
             LibChamps:=LbChamps.Cells[2,Curlig];
             if Curlig < PosSeparateur then
                begin
                For k:=0 to P.ControlCount-1 do
                    begin
                    C:=P.Controls[k];
                    if ( C is TLabel ) or ( C.Enabled = False ) or ( VerrouModif ) then
                    {RAF} else
                       begin
                       if C is TListBox then
                          begin
                          TListBox(C).Clear;
                          TaillePanelMulti(P,TListBox(C),Curlig);
                          end
                       else
                           if (C is THEdit) and (THEdit(C).OpeType=otDate) then
                              SetControlText (LBChamps.Cells[0,Curlig],DateToStr(iDate1900))
                           else
                              SetControlText (LBChamps.Cells[0,Curlig],'');
                       end;
                    end;
                end;
             end;
          end;
       end;
//=========
    end
  else
    if (TobType <> Nil) and (copy(InfosOnglets[i,1],1,3) = 'RAC') then
    begin
      NomChamp:='RPA_'+ copy(InfosOnglets[i,1],5,length(InfosOnglets[i,1])) ;
      trouve:=false;
      CritOnglet:= InfosOnglets[i,3];
      Repeat
        Critere:=ReadTokenSt(CritOnglet) ;
        if Critere<>'' then
        begin
          if TobType.GetValue(NomChamp) = Critere then
            begin
            trouve:=true;
            if PremierOnglet=0 then PremierOnglet:=i;
            break;
            end;
         end;
      until (Critere = '') ;
      if not trouve then
            TS.TabVisible:=False;
      end;
  { inversement des taborder }
  nbpanel:=0;
  ListeActive:= TScrollBox(GetControl('SCB'+IntToStr(i)));
  if assigned(ListeActive) then
    For j := ListeActive.ControlCount - 1 downto 0 do
      begin
      if ListeActive.Controls[j] is TPanel then
        begin
        P := TPanel(ListeActive.Controls[j]);
        P.TabOrder:=nbpanel;
        Inc(nbpanel);
        end;
    end;
end;
end;

Procedure Tof_ParamCL.TestInfoOnglet(i : integer; TobInfos : Tob; Var PremierOnglet : integer; NomChamp : String; TS: TTabSheet);
var trouve : boolean;
    CritOnglet,Critere : String;
begin
  if InfosOnglets[i,2] = 'BOOLEAN' then
     begin
     if TobInfos.GetValue(NomChamp) <> InfosOnglets[i,3] then
        // on ne supprime pas l'Onglet mais on le rend invisible
        TS.TabVisible:=False
     else
        if PremierOnglet=0 then PremierOnglet:=i;
     end;
  if InfosOnglets[i,2] = 'COMBO' then
     begin
     trouve:=false;
     CritOnglet:= InfosOnglets[i,3];
     Repeat
      Critere:=ReadTokenSt(CritOnglet) ;
      if Critere<>'' then
         begin
         if TobInfos.GetValue(NomChamp) = Critere then
            begin
            trouve:=true;
            if PremierOnglet=0 then PremierOnglet:=i;
            break;
            end;
         end;
     until (Critere = '') ;
     if not trouve then
            // on ne supprime pas l'Onglet mais on le rend invisible
            TS.TabVisible:=False;
     end;
end;

procedure Tof_ParamCL.OnArgument(Arguments : string) ;
var Critere,ChampMul,ValMul,CleInfos : string;
    x,i,j,iTable,iChamp : integer;
    bOnglet,btestfiche,bDelete,bValDef : TToolBarButton97 ;
    //cFichier : THValComboBox ;
    TobCl : Tob;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;
begin
inherited ;
MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();

bLignes:=false;
CouleurPanel:=clAqua;
MajParamChamp :=False;
ErreurAppel:=False;
BougePropriete:=false;
BougeOnglet:=False;
{récup des arguments }
  PremierOnglet:=1;
  TypeAction:=''; rac_TypeAction:='';
  bOnglet := TToolBarButton97(GetControl('BPARAMONGLET'));
  if bOnglet <> Nil then
    bOnglet.OnClick := bOngletClick;
  CodeFichier:=CodeFichierProspect;

  btestfiche := TToolBarButton97(GetControl('BTESTFICHE'));
  if btestfiche <> Nil then
    btestfiche.OnClick := btestficheClick;

  bDelete := TToolBarButton97(GetControl('BDelete'));
  if bDelete <> Nil then
    bDelete.OnClick := bDeleteClick;

  bValDef := TToolBarButton97(GetControl('BVALDEF'));
  if bValDef <> Nil then
      bValDef.OnClick := bValDefClick;

  FinBoucle := False;
  OrigineFichier:=False;   ExisteInfos:=false; OrigineChampOblig:=false;
  OrigineParam:=False;
  VerrouModif := False;
  Creation := false;

// Ajout CD 221100
i := pos('MODIFLOT',Arguments);
ModifLot := i<>0;
if ModifLot then
  begin
  OrigineFichier:=true;
  Repeat
    { mng 03/01/05 pourquoi transformer en majuscule ? FQ no 010-12090
    Critere:=uppercase(Trim(ReadTokenPipe(Arguments,'^'))) ;}
    Critere:=Trim(ReadTokenPipe(Arguments,'^')) ;
    if Critere<>'' then
      begin
        x:=pos('=',Critere);
        if x<>0 then
        begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           if ChampMul='CODEINFOS' then
              CodeInfos :=ValMul;
           if ChampMul='MODIFLOT' then
              CodeFichier :=ValMul;
           if ChampMul='SQL' then
              StSQL :=ValMul;
           if ChampMul='ACTION' then
              VerrouModif :=true;
        end;
      end;
  until  Critere='';
  end
else
  begin
  Repeat
      Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
      if Critere<>'' then
      begin
          x:=pos('=',Critere);
          if x<>0 then
          begin
             ChampMul:=copy(Critere,1,x-1);
             ValMul:=copy(Critere,x+1,length(Critere));
             if ChampMul='NATUREAUXI' then
                SetControlText ('NATUREAUXI',ValMul);
             if ChampMul='FICHEINFOS' then
                begin
                Ecran.Caption:=TraduireMemoire('Informations complémentaires');
                CleInfos := ValMul;
                OrigineFichier:=true;
                end;

             if ChampMul='FICHEPARAM' then
             begin
                CodeFichier:=ValMul;
                j:=pos('_', ValMul);
                if (j <> 0) and (copy(ValMul,length(ValMul)-3,4) <> '_FIC')
                   and (copy(ValMul,length(ValMul)-3,4) <> '_FIL')
                   and (copy(ValMul,length(ValMul)-3,4) <> '_FSL')then
                   CodeFichier:=copy(ValMul,1,j-1);
                { mng 15/12/04 : cas particulier des infos compl lignes : récup par Latob}
                if ValMul='LIGNESPIECES' then
                   bLignes:=true;
             end;
             if ChampMul='TOBPIECE' then TobPiece:= Tob(ValeurI(ValMul));
             if (ChampMul='ACTION') then
                 if (ValMul='CONSULTATION') then
                     VerrouModif :=true
                 else if (ValMul='CREATION') then
                     Creation :=true;
             if ChampMul='TYPEACTION' then
                TypeAction := ValMul;
             if ChampMul='RAC_TYPEACTION' then
                rac_TypeAction := ValMul;
          end
         else
           begin
           x:=pos('PARAMFICHE',Critere);
           if x<>0 then
              begin
              OrigineParam:=true;
              CodeFichier:=copy(Critere,11,11);
              SetControlEnabled('HelpBtn',FALSE) ;
              Ecran.Caption:=TraduireMemoire('Informations complémentaires');
              end;
           if pos('EXISTOBLIG',Critere) > 0 then
              OrigineChampOblig:=true;
           end;
      end;
  until  Critere='';

  CodeInfos:=StringReplace(CleInfos,'|',';',[rfReplaceAll]);
  end;



if OrigineParam or OrigineFichier then
   BoutonsParamInvisibles;

if (Not OrigineParam) and (CodeFichier<> '') then
    begin

    VH_RT.TobChampsPro.Load;
    
    TobCl:=VH_RT.TobParamCl.FindFirst(['CO_TYPE','CO_ABREGE'],['RPC',CodeFichier],TRUE) ;
    if TobCl = Nil then
        begin
        PGIBox(Format('Le code %s n''existe pas',[CodeFichier]),'Code inexistant');
        ErreurAppel:=true;
        BoutonsParamInvisibles;
        SetControlVisible('BValider',FALSE) ;
        SetControlVisible('HelpBtn',FALSE) ;
        Exit;
        end;
    CodeFichier:=TobCl.GetValue('CO_CODE');
    ecran.caption:=ecran.caption+' '+LowerCase(TobCl.GetValue('CO_LIBELLE'));
    UpdateCaption(Ecran);
    end
else
    if (CodeFichier= '') then
       begin
       PGIBox('Le paramètre code fichier est vide','Code fichier');
       BoutonsParamInvisibles;
       SetControlVisible('BValider',FALSE) ;
       SetControlVisible('HelpBtn',FALSE) ;
       ErreurAppel:=true;
       exit;
       end;

  if (ModifLot) and (CodeFichier='3') then SetControlText ('NATUREAUXI','FOU');         
{ calcul du nombre maxi de combos, de champs et du nombre de lignes dans le grid }
  MaxCombo:=0;
  MaxMultiChoix:=0;
  MaxTexte:=0;
  MaxValeur:=0;
  MaxBool:=0;
  MaxDate:=0;
  Nbpost:=0;
  if CodeFichier = CodeFichierProspect then Prefixe:=PrefixeProspect
     else Prefixe:='RD'+CodeFichier;

  table := Mcd.getTable(mcd.PrefixetoTable(prefixe));
  FieldList := Table.Fields;
  FieldList.Reset();
  While FieldList.MoveNext do
  begin
    if copy((FieldList.Current as IFieldCOM).name,8,8) = 'LIBTABLE' then inc(MaxCombo);
    if copy((FieldList.Current as IFieldCOM).name,8,8) = 'LIBTEXTE' then inc(MaxTexte);
    if copy((FieldList.Current as IFieldCOM).name,8,6) = 'LIBMUL' then inc(MaxMultiChoix);
    if copy((FieldList.Current as IFieldCOM).name,8,6) = 'LIBVAL' then inc(MaxValeur);
    if copy((FieldList.Current as IFieldCOM).name,8,7) = 'LIBBOOL' then inc(MaxBool);
    if copy((FieldList.Current as IFieldCOM).name,8,7) = 'LIBDATE' then inc(MaxDate);
  end;
  { nb combos + 5 types de champ * nb champs + bloc-notes + ligne Espace + ligne fine }
  NbPost:=MaxCombo+MaxMultiChoix+MaxTexte+MaxValeur+MaxBool+MaxDate+3;
if CodeFichier = CodeFichierProspect then
    begin
    TabletteLibelle:=TabletteProspect;
    Prefixe:=PrefixeProspect;
    TabletteCombo:=NomTabletteCombo;
    TabletteComboMul:=NomTabletteComboMul;
    ParamSocLargeur:=ParamSocLargeurFiche ;
    ParamSocHauteur:=ParamSocHauteurFiche ;
    end
else
    begin
    PrefixeAutres:= '00'+CodeFichier;
    TabletteLibelle:=TabletteAutres+'00'+CodeFichier;
    Prefixe:=PrefixeProspect;
    TabletteCombo:=copy(NomTabletteCombo,1,2)+'00'+Codefichier+copy(NomTabletteCombo,6,13) ;
    TabletteComboMul:=copy(NomTabletteComboMul,1,2)+'00'+Codefichier+copy(NomTabletteComboMul,6,14) ;
    ParamSocLargeur:=ParamSocLargeurFiche+'00'+Codefichier ;
    ParamSocHauteur:=ParamSocHauteurFiche+'00'+Codefichier ;
    end;
If (OrigineFichier) and (CodeFichier=CodeFichierAction) and
   (TypeAction<>'') and (CodeInfos<>'X;0') then
   bValDef.Visible:=True;

LBChamps :=THGrid(GetControl ('LBChamps'));
ParamChamps:=0;
//TFVierge(Ecran).Left:=200;
end;

procedure Tof_ParamCL.OnLoad;
var i,x,IVal2 : integer;
    Q : TQuery;
    StLibelle,FieldNameFrom,FieldNameTo,StVal1,stAuxi : string;
    TCompTmp : TComponent;
    TL : TListBox;
    B: TToolbarButton97;
begin
inherited ;
if ErreurAppel then exit;
TobInfos:=Nil;
TobLibres:=Nil; TobTiers:=Nil;
TobDesc:=Nil;
TobAction:=Nil;

if (OrigineFichier) and ( TypeAction = '' )then
    begin
    //TobInfos:=TOB.create ('RTINFOSDESC',NIL,-1);
    // je suis obligée de travailler dans la tob prospect c'est la plus grande !
    TobInfos:=TOB.create ('PROSPECTS',NIL,-1);
    if (CodeFichier <> CodeFichierProspect) and (CodeFichier <> CodeFichierLigne)
    and (CodeFichier <> CodeFichierPiece) and (CodeFichier <> CodeFichierFournisseur) then
        begin
        TobDesc:=TOB.create ('RTINFOS00'+CodeFichier,NIL,-1);
        Q:=OpenSQL('Select * from RTINFOS00'+CodeFichier+' where RD'+CodeFichier+'_CLEDATA="'+CodeInfos+'"',True);
        if not TobDesc.selectDB ('',Q,False) then
           begin
           TobDesc.putvalue('RD'+CodeFichier+'_CLEDATA',CodeInfos);
           for i:=0 to (MaxDate-1) do  TobDesc.putvalue('RD'+CodeFichier+'_RD'+CodeFichier+'LIBDATE'+intToStr(i),iDate1900);
           // dans le cas des actions, on prend les valeurs par défaut paramétrées
           // au niveau du type d'action, si elles existent !!
           if (CodeInfos='X;0') or (TypeAction = '') then
               TobType:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[TypeAction,'---',0],TRUE)
             else
             begin
               StVal1:=ReadToKenSt(CodeInfos);
               IVal2:=StrToInt(ReadToKenSt(CodeInfos));
               TobType:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[TypeAction,StVal1,IVal2],TRUE) ;
             end;
             if TobType <> Nil then
                 begin
                 ExisteInfos:=True;
                 for i :=1 to TobType.NbChamps do
                    begin
                    FieldNameFrom := TobType.GetNomChamp(i);
                    if (copy(FieldNameFrom,1,7) <> 'RPA_RPR') and
                       (copy(FieldNameFrom,1,12) <> 'RPA_BLOCNOTE') then continue;
                    FieldNameTo  := 'RPR_' + copy(FieldNameFrom,5,length(FieldNameFrom)) ;
                    TobInfos.PutValue(FieldNameTo, TobType.GetValue(FieldNameFrom));
                 end;
             end;
           end
        else
           begin
           ExisteInfos:=True;
           for i :=1 to TobDesc.NbChamps do
                begin
                FieldNameFrom := TobDesc.GetNomChamp(i);
                x:= Pos(copy(FieldNameFrom,4,length(FieldNameFrom)),ChampsNonTraitesRDA);
                if x <> 0 then continue;
                if copy(FieldNameFrom,5,12) = 'BLOCNOTE' then
                    FieldNameTo:='RPR_BLOCNOTE'
                else
                    FieldNameTo  := 'RPR_RPR' + copy(FieldNameFrom,8,length(FieldNameFrom)) ;
                TobInfos.PutValue(FieldNameTo, TobDesc.GetValue(FieldNameFrom));
                end;
           end;
        Ferme (Q);
        if (CodeFichier = CodeFichierArticle) then
          begin
          TobArticle:=TOB.create ('ARTICLE',NIL,-1);
          TobArticle.selectDB ('"'+CodeInfos+'"',Nil,False);
          end
        else
          if (CodeFichier = CodeFichierArticleParc) then
            begin
            TobArticleParc:=TOB.create ('ARTICLEPARC',NIL,-1);
            TobArticleParc.selectDB ('"'+CodeInfos+'"',Nil,False);
            end
          else
            if (CodeFichier = CodeFichierParc) then
              begin
              TobParc:=TOB.create ('WPARC',NIL,-1);
              TobParc.selectDB ('"'+CodeInfos+'"',Nil,False);
              end
            else
              if (CodeFichier = CodeFichierElementParc) then
                begin
                TobParcNome:=TOB.create ('WPARCNOME',NIL,-1);
                TobParcNome.selectDB ('"'+CodeInfos+'"',Nil,False);
                end
            else
              if (CodeFichier = CodeFichierProposition) then
                begin
                TobProposition:=TOB.create ('PERSPECTIVES',NIL,-1);
                TobProposition.selectDB ('"'+CodeInfos+'"',Nil,False);
                end;

        end
    else   { cas des infos compl. tiers }
      if (CodeFichier = CodeFichierProspect) or (CodeFichier = CodeFichierFournisseur) then
          begin
          if not TobInfos.selectDB ('"'+CodeInfos+'"',Nil,False) then
             begin
             TobInfos.putvalue('RPR_AUXILIAIRE',CodeInfos);
             for i:=0 to (MaxDate-1) do  TobInfos.putvalue('RPR_RPRLIBDATE'+intToStr(i),iDate1900);
             end
          else
             ExisteInfos:=True;

          TobLibres:=TOB.create ('TIERSCOMPL',NIL,-1);
          TobLibres.selectDB ('"'+CodeInfos+'"',Nil,False);
          TobTiers:=TOB.create ('TIERS',NIL,-1);
          TobTiers.selectDB ('"'+CodeInfos+'"',Nil,False);
          end
          { cas des infos compl. lignes, valeur dans LaTob }
      else
          if (CodeFichier = CodeFichierLigne) then
          begin
          TobLigne:=TOB.create ('LIGNE',NIL,-1);
          { spécial pouir la ligne, on a les infos dans latob
            TobLigne.selectDB ('"'+CodeInfos+'"',Nil,False);}
          for i := 0 to TobLigne.nbchamps-1 do
            if (TobLigne.GetNomChamp(i)<>'') and (Latob.FieldExists(TobLigne.GetNomChamp(i))) then
               TobLigne.PutValue(TobLigne.GetNomChamp(i),LaTob.GetValue(TobLigne.GetNomChamp(i)));
          TobDesc:=TOB.create ('RTINFOS00'+CodeFichier,NIL,-1);
          if ( pos(';',LaTob.GetString('CLEDATA_RD8')) <> 0 ) or
             ( LaTob.GetString('VALID_INFOSCOMPL') = 'X' ) then
            ExisteInfos:=True;
          for i := 1000 to (1000+(LaTob.ChampsSup.count-1)) do
            begin
            FieldNameFrom := LaTob.GetNomChamp(i);
            if copy(FieldNameFrom,1,3) <> 'RD8' then continue;
            if copy(FieldNameFrom,5,12) = 'BLOCNOTE' then
                FieldNameTo:='RPR_BLOCNOTE'
            else
                FieldNameTo  := 'RPR_RPR' + copy(FieldNameFrom,8,length(FieldNameFrom)) ;
            TobInfos.PutValue(FieldNameTo, LaTob.GetValue(FieldNameFrom));
            end;
          end
        else
            if (CodeFichier = CodeFichierPiece) then
            begin
            TobDesc:=TOB.create ('RTINFOS00'+CodeFichier,NIL,-1);
            TobDesc.Dupliquer(LaTob, False, True);
            if ( TobDesc.GetValue('RDD_CLEDATA') <> '' )  or
             ( LaTob.GetString('VALID_INFOSCOMPL') = 'X' ) then
              ExisteInfos:=True;
            for i :=1 to TobDesc.NbChamps do
                begin
                FieldNameFrom := TobDesc.GetNomChamp(i);
                x:= Pos(copy(FieldNameFrom,4,length(FieldNameFrom)),ChampsNonTraitesRDA);
                if x <> 0 then continue;
                if copy(FieldNameFrom,5,12) = 'BLOCNOTE' then
                    FieldNameTo:='RPR_BLOCNOTE'
                else
                    FieldNameTo  := 'RPR_RPR' + copy(FieldNameFrom,8,length(FieldNameFrom)) ;
                TobInfos.PutValue(FieldNameTo, TobDesc.GetValue(FieldNameFrom));
                end;
            end;
    end;

if TypeAction <> '' then
    begin
    TobInfos:=TOB.create ('PROSPECTS',NIL,-1);
    if (CodeInfos='X;0') or (TypeAction = '') then
         TobType:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[TypeAction,'---',0],TRUE)
    else
    begin
       StVal1:=ReadToKenSt(CodeInfos);
       IVal2:=StrToInt(ReadToKenSt(CodeInfos));
       TobType:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[TypeAction,StVal1,IVal2],TRUE) ;
    end;
    if TobType = Nil then
       for i:=0 to (MaxDate-1) do  TobInfos.putvalue('RPR_RPRLIBDATE'+intToStr(i),iDate1900)
    else
       begin
       ExisteInfos:=True;
       for i :=1 to TobType.NbChamps do
          begin
          FieldNameFrom := TobType.GetNomChamp(i);
          if (copy(FieldNameFrom,1,7) <> 'RPA_RPR') and
              (copy(FieldNameFrom,1,12) <> 'RPA_BLOCNOTE') then continue;
          FieldNameTo  := 'RPR_' + copy(FieldNameFrom,5,length(FieldNameFrom)) ;
          TobInfos.PutValue(FieldNameTo, TobType.GetValue(FieldNameFrom));
          end;
       end;
    end;
InitForme ;

if ModifLot then
   begin
   if CodeFichier = CodeFichierProspect then
     begin
     Q := OpenSQL('SELECT T_LIBELLE,T_NATUREAUXI FROM TIERS WHERE T_AUXILIAIRE="'+CodeInfos+'"', True);
     if not Q.Eof then
     begin
     StLibelle:=Q.FindField('T_LIBELLE').AsString;
     SetControlText ('NATUREAUXI',Q.FindField('T_NATUREAUXI').AsString);
     end;
     Ferme(Q) ;
     Ecran.Caption:=TiersAuxiliaire(CodeInfos,true)+' : '+StLibelle;
     end;
   if CodeFichier = CodeFichierContact then
     begin
     StLibelle:=CodeInfos;
     stAuxi:=ReadToKenSt(StLibelle);
     stAuxi:=ReadToKenSt(StLibelle);
     Q := OpenSQL('SELECT C_NOM,C_PRENOM FROM CONTACT WHERE C_TYPECONTACT="T" AND C_AUXILIAIRE="'+stAuxi+'" AND C_NUMEROCONTACT='+StLibelle, True);
     if not Q.Eof then
        StLibelle:=Q.FindField('C_PRENOM').AsString+' '+Q.FindField('C_NOM').AsString;
     Ferme(Q) ;
     Ecran.Caption:=StLibelle;
     end;
   UpdateCaption(Ecran);
   SetArgumentsCL(StSQL);
   end ;

{$IFNDEF CCS3}
if (OrigineFichier) then
   if (CodeFichier = CodeFichierProspect) or (CodeFichier =  CodeFichierFournisseur) then
     AppliquerConfidentialite(Ecran,GetControlText ('NATUREAUXI'))
   else
     if (CodeFichier = CodeFichierArticle) and (Assigned(TobArticle)) then
       AppliquerConfidentialite(Ecran,TobArticle.getString('GA_TYPEARTICLE')+TobArticle.getString('GA_TYPENOMENC'))
     else
       if (CodeFichier = CodeFichierContact) or ( CodeFichier = CodeFichierArticleParc ) or
          ( CodeFichier = CodeFichierParc ) or ( CodeFichier = CodeFichierElemParc ) or
          ( CodeFichier = CodeFichierProjet) or (CodeFichier = CodeFichierOperation) or
          (CodeFichier = CodeFichierAction) or ( CodeFichier = CodeFichierIntervention )
          or ( CodeFichier = CodeFichierProposition ) then
         AppliquerConfidentialite(Ecran,CodeFichier);

// boucle sur les multi-choix pour traiter le bouton si une restriction existe
for i := 0 to Ecran.ComponentCount - 1 do
    begin
    TCompTmp := Ecran.Components[i];
    FieldNameTo := TCompTmp.Name;
    if copy(FieldNameTo,1,3) = 'BML' then
        begin
        TL:= TListBox(Ecran.FindComponent(copy(FieldNameTo,2,3)));
        B:= TToolbarButton97(GetControl(FieldNameTo));
        if (TL <> Nil) and (B <> Nil) then
            begin
            if TL.Visible = False then B.Visible:=False;
            if TL.Enabled = False then B.Enabled:=False;
            end;
        end;
    end;
{$ENDIF}

end;


procedure Tof_ParamCL.OnClose;
var i : integer;
    Bouge : Boolean;
    Valeur,Mess : String;
    TL: TListBox;
begin
inherited ;
if ErreurAppel then exit;
Bouge:=False;
// modif propriétés champ ou Param Onglet ou Déplacement d'un champ
if ( not ExisteInfos) and ( BougePropriete or BougeOnglet or BougeChamp) then
   Bouge:=True;

// modif d'un libelle de champ dans la liste
if ( not ExisteInfos) and (not Bouge) and (not MajParamChamp) then
   begin
   for i:= 0 to ( NbPost - 1 ) do
       if LBChamps.Cells[9,i]<> LBChamps.Cells[2,i] then
          begin
          Bouge:=True;
          break;
          end;
   end;

// modif d'une valeur d'un multi-choix, non stocké dans tob
if (not Bouge) and (not MajParamChamp) then
   begin
   for i:= PremierMul to ( PremierMul + MaxMultiChoix - 1 ) do
        begin
        if LBChamps.RowHeights[i]<>0 then continue;
        TL:= TListBox(GetControl(LBChamps.Cells[0,i]));
        if ExisteInfos then
           ChargeChampMulti (TListBox(TL),Valeur,False,True)
        else
           ChargeChampMulti (TListBox(TL),Valeur,False,False);
        if LBChamps.Cells[8,i]<> Valeur then
          begin
          Bouge:=True;
          break;
          end;
        end;
   end;
// modif des valeurs des chanps stockées en tob
if (not Bouge) and (not MajParamChamp) then
   begin
   TobEcran.GetEcran(TFVierge(ecran),Nil);
   if RTCompareTobs(TobEcran,SavTobEcran)=False then Bouge:=True;
   end;

  Mess:='Voulez-vous enregistrer les modifications ?';
  if OrigineChampOblig then
    Mess:='Il existe des champs obligatoires. '+Mess;

  if ( (Bouge) and not (OrigineParam) ) then
    Case PGIAskCancel(Mess,TFVierge(ecran).Caption) of
      mrCancel :
          begin
          LastError:=1;
          exit;
          end;
      mrYes : ValidModifs;
      mrNo :
          begin
            { si un champ est obligatoire, la fiche s'ouvre automatiquement
              on pouvait faire annuler sans que l'enreg soit créé }
            if (OrigineFichier) and ( CodeFichier = CodeFichierProspect ) and ( not ExisteInfos ) then
               TobInfos.InsertOrUpdateDB(False);
          end;
    end
  else
    if (OrigineChampOblig) and (not MajParamChamp) and (ExisteOngletChampOblig) then
      begin
      Mess:='Il existe des champs obligatoires. Voulez-vous revenir en saisie ?';
      Case PGIAskCancel(Mess,TFVierge(ecran).Caption) of
        mrYes,
        mrCancel :
            begin
            LastError:=1;
            exit;
            end;
        mrNo :
            begin
              { si un champ est obligatoire, la fiche s'ouvre automatiquement
                on pouvait faire annuler sans que l'enreg soit créé }
              if (OrigineFichier) and ( CodeFichier = CodeFichierProspect ) and ( not ExisteInfos ) then
                 TobInfos.InsertOrUpdateDB(False);
            end;
      end;
      end;
if OrigineFichier and not FinBoucle then
   begin
   if Assigned(TobLibres) then
     FreeAndNil(TobLibres);
   if Assigned(TobTiers) then
     FreeAndNil(TobTiers);
   if Assigned(TobInfos) then
     FreeAndNil(TobInfos);
   if CodeFichier <> CodeFichierProspect then
     if Assigned(TobDesc) then
       FreeAndNil(TobDesc);
   if Assigned(TobAction) then
     FreeAndNil(TobAction);
   if Assigned(TobArticle) then
     FreeAndNil(TobArticle);
   if Assigned(TobArticleParc) then
     FreeAndNil(TobArticleParc);
   if Assigned(TobParc) then
     FreeAndNil(TobParc);
   if Assigned(TobParcNome) then
     FreeAndNil(TobParcNome);
   if Assigned(TobLigne) then
     FreeAndNil(TobLigne);
   if Assigned(TobProposition) then
     FreeAndNil(TobProposition);
   end;

if FinBoucle then
   begin
   LastError := 1 ;
   FinBoucle := False;
   end
else
   begin
   if Assigned(SavTobEcran) then
     FreeAndNil(SavTobEcran);
   if Assigned(TobEcran) then
     FreeAndNil(TobEcran);
   end;

If OrigineParam then
   begin
   SetParamsoc(ParamSocLargeur,TFVierge(ecran).width) ;
   SetParamsoc(ParamSocHauteur,TFVierge(ecran).height) ;
   end;
end;

procedure Tof_ParamCL.OnUpdate;
begin
inherited ;
ValidModifs;
end;

procedure Tof_ParamCL.ValidModifs;
begin
MajParamChamp:=True;
BougePropriete:=False;
BougeOnglet:=False;
BougeChamp:=False;

if not OrigineFichier and not OrigineParam then
  begin
  Transactions(MajParametrage,1);
  AvertirTable (TabletteLibelle) ;
  VH_RT.TobChampsPro.ClearDetail;
  VH_RT.TobChampsPro.Load (true);
  end
else
  if (OrigineFichier) and (not VerrouModif) then
    MajInfosCompl ;
end;

procedure Tof_ParamCL.MajInfosCompl ;
var st,LibChamps,Requete : string;
    P : TPanel;
    ListeActive : TScrollBox;
    OngletActif,i,j,x : Integer;
    TS: TTabSheet;
    C : TControl;
    Champ,Valeur,FieldNameFrom,FieldNameTo : String;
begin
     Valeur:='';
     FinBoucle:=False;
     // Balayage des Onglets
     for OngletActif := 1 to TPageControl(getcontrol('PAGEONGLETS')).PageCount do
     begin
        TS:= TTabSheet(Ecran.FindComponent('ONGLETLIBRE'+IntToStr(OngletActif)));
        if TS=Nil then break;
        if TS.TabVisible <> False then
        begin
           ListeActive:= TScrollBox(GetControl('SCB'+IntToStr(OngletActif)));
           For i := 0 to ListeActive.ControlCount - 1 do
           begin
              if ListeActive.Controls[i] is TPanel then
              begin
                 P := TPanel(ListeActive.Controls[i]);
                 St:= P.Name;
                 if Pos('PCHAMPS_',St)<=0 then exit;
                 CurLig:=StrToInt(Copy(St,pos('_',St)+1,Length(St)-pos('_',St)));
                 LibChamps:=LbChamps.Cells[2,Curlig];
                  if Curlig < PosSeparateur then
                  begin
                    For j:=0 to P.ControlCount-1 do
                    BEGIN
                        C:=P.Controls[j];
                        if ( C is TLabel ) or ( C.Enabled = False ) or ( VerrouModif ) then {RAF} else
                           BEGIN
                           Champ:=LBChamps.Cells[3,Curlig];
                           if C is TCheckBox then
                              BEGIN
                              if TCheckBox(C).state=cbChecked then Valeur:='X'
                              else if TCheckBox(C).state=cbUnChecked then Valeur:='-';
                              TobInfos.PutValue (Champ,Valeur);
                              END
                           else if C is THValComboBox then
                              begin
                              if ( LbChamps.Cells[5,Curlig] = 'X' ) and ( THValComboBox(C).Value = '' )
                                 and ( TypeAction = '' ) then
                                 begin
                                InfoOblig (LibChamps,OngletActif,i);
                                FinBoucle:=True;
                                break;
                                end;
                                TobInfos.PutValue (Champ,THValComboBox(C).Value );
                              end
                           else if C is TListBox then
                              begin
                              ChargeChampMulti (TListBox(C),Valeur,false,True);
                              if ( LbChamps.Cells[5,Curlig] = 'X' ) and ( Valeur = '' )
                                 and ( TypeAction = '' ) then
                                 begin
                                InfoOblig (LibChamps,OngletActif,i);
                                FinBoucle:=True;
                                break;
                                end;
                              TobInfos.PutValue (Champ,Valeur );
                              end
                           else if C is THEdit then
                              begin
                              if ( LbChamps.Cells[5,Curlig] = 'X' ) and (
                                  ( ( THEdit(C).Text = '' ) and ( THEdit(C).OpeType <> otDate) ) or
                                  ( ( THEdit(C).Text = DateToStr(iDate1900) ) and ( THEdit(C).OpeType = otDate) ) )
                                 and ( TypeAction = '' ) then
                                 begin
                                InfoOblig (LibChamps,OngletActif,i);
                                FinBoucle:=True;
                                break;
                                end;
                                if ( LbChamps.Cells[7,Curlig] <> '' ) and ( THEdit(C).Text <> '' )  then
                                     begin
                                     if LBChamps.Cells[7,CurLig] = 'T' then
                                        begin
                                           if GetControlText ('NATUREAUXI') = 'FOU' then
                                              Requete:='Select T_LIBELLE from TIERS where T_TIERS="'+THEdit(C).text+
                                              '" and (T_NATUREAUXI = "FOU")'
                                           else
                                         {$ifdef GIGI}
                                              Requete:='Select T_LIBELLE from TIERS where T_TIERS="'+THEdit(C).text+
                                              '" and (T_NATUREAUXI = "CLI" or T_NATUREAUXI="PRO"  or T_NATUREAUXI="NCP" or T_NATUREAUXI="CON")';
                                         {$else}
                                              Requete:='Select T_LIBELLE from TIERS where T_TIERS="'+THEdit(C).text+
                                              '" and (T_NATUREAUXI = "CLI" or T_NATUREAUXI="PRO" or T_NATUREAUXI="CON")';
                                         {$endif}
                                        end
                                     else
                                     if LBChamps.Cells[7,CurLig] = 'R' then
                                         Requete:=RequeteRessource+THEdit(C).text+'"'
                                     else
                                       if LBChamps.Cells[7,CurLig] = 'O' then
                                           Requete:=RequeteOperation+THEdit(C).text+'"'
                                       else
                                         if LBChamps.Cells[7,CurLig] = 'P' then
                                             Requete:=RequeteProjet+THEdit(C).text+'"'
                                         else
                                            if LBChamps.Cells[7,CurLig] = 'C' then
                                                begin
                                                if CodeFichier = CodeFichierProspect then
                                                   Requete:='Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'
                                                      +CodeInfos+'" AND C_NUMEROCONTACT='+THEdit(C).text
                                                else
                                                   if CodeFichier = CodeFichierAction then
                                                      Requete:='Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'
                                                         +copy(CodeInfos,1,(Pos(';',CodeInfos)-1))+'" AND C_NUMEROCONTACT='+THEdit(C).text
                                                   else
                                                     if (CodeFichier = CodeFichierParc) and (Assigned(TobParc)) then
                                                       begin
                                                       if TobParc.GetString('WPC_TIERS') <> '' then
                                                         Requete:='Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'
                                                            +TiersAuxiliaire(TobParc.GetString('WPC_TIERS'),false)+'" AND C_NUMEROCONTACT='+THEdit(C).text;
                                                       end
                                                     else
                                                     if (CodeFichier = CodeFichierProposition) and (Assigned(TobProposition)) then
                                                       begin
                                                       if TobProposition.GetString('RPE_TIERS') <> '' then
                                                         Requete:='Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'
                                                            +TiersAuxiliaire(TobProposition.GetString('RPE_TIERS'),false)+'" AND C_NUMEROCONTACT='+THEdit(C).text;
                                                       end
                                                     else
                                                       if (CodeFichier = CodeFichierPiece) then
                                                         Requete:='Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'
                                                            +TiersAuxiliaire(TobPiece.GetString('GP_TIERS'),false)+'" AND C_NUMEROCONTACT='+THEdit(C).text
                                                       else
                                                         Requete:='Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'
                                                           +copy(CodeInfos,3,(pos(';',copy(CodeInfos,3,length(CodeInfos))))-1)+'" AND C_NUMEROCONTACT='+THEdit(C).text;
                                                end;
                                         if ( Requete <> '') and ( not ExisteSQL(Requete) ) then
                                            begin
                                            PGIBox(Format('Le code %s n''existe pas',[THEdit(C).Text]),'Code Inexistant');
                                            FinBoucle:=True;
                                            break;
                                            end;
                                        end;
                                if THEdit(C).OpeType=otDate then
                                   TobInfos.PutValue(Champ,StrToDate(THEdit(C).Text))
                                else TobInfos.PutValue (Champ,THEdit(C).Text ) ;
                              end
                           else if C is THNumEdit then
                              begin
                              if ( LbChamps.Cells[5,Curlig] = 'X' ) and ( THNumEdit(C).Value = 0 )
                                 and ( TypeAction = '' ) then
                                 begin
                                InfoOblig (LibChamps,OngletActif,i);
                                FinBoucle:=True;
                                break;
                                end;
                              TobInfos.PutValue (Champ,THNumEdit(C).Value);
                              end
                           else if C is THRichEditOle then
                              begin
                              if ( LbChamps.Cells[5,Curlig] = 'X' ) and ( THRichEditOle(C).Lines.Text = '' )
                                 and ( TypeAction = '' ) then
                                 begin
                                InfoOblig (LibChamps,OngletActif,i);
                                FinBoucle:=True;
                                break;
                                end;
                              TobInfos.PutValue (Champ,THRichEditOle(C).LinesRTF.Text);
                              end;
                           END;
                    END;
                  end;
              end;
              if FinBoucle then break;
           end;
        end;
        if FinBoucle then break;
     end;
{     if not FinBoucle and MajPro then
        TobInfos.InsertOrUpdateDB(False);}
     // cas de l'appel de la fiche depuis le paramétrage des types d'action
     if TypeAction = '' then
         begin
         if not finBoucle then
             begin
             if (CodeFichier = CodeFichierProspect) then
                 begin
                 if (VH_RT.RTCreatInfos) or (RTControleModifTiers(TFVierge(Ecran), TobInfos, ExisteInfos)) then
                    TobInfos.InsertOrUpdateDB(False)
                 else
                    finBoucle:=True;
                 end
             else
               if (CodeFichier = CodeFichierFournisseur) then
                   begin
                   if (VH_RT.RFCreatInfos) or (RTControleModifFou(TFVierge(Ecran), TobInfos, ExisteInfos)) then
                      TobInfos.InsertOrUpdateDB(False)
                   else
                      finBoucle:=True;
                   end
               else
                 if (CodeFichier <> CodeFichierLigne) and (CodeFichier <> CodeFichierPiece) then
                   begin
                   for i :=1 to TobDesc.NbChamps do
                      begin
                      FieldNameTo := TobDesc.GetNomChamp(i);
                      x:= Pos(copy(FieldNameTo,4,length(FieldNameTo)),ChampsNonTraitesRDA);
                      if x <> 0 then continue;

                      //FieldNameFrom  := 'RPR_' + copy(FieldNameTo,5,length(FieldNameTo)) ;
                      if copy(FieldNameTo,5,12) = 'BLOCNOTE' then
                          FieldNameFrom:='RPR_BLOCNOTE'
                      else
                          FieldNameFrom  := 'RPR_RPR' + copy(FieldNameTo,8,length(FieldNameTo)) ;
                      TobDesc.PutValue(FieldNameTo, TobInfos.GetValue(FieldNameFrom));
                      end;
                   if (CodeFichier = CodeFichierFournisseur) then
                      if (not VH_RT.RFCreatInfos) and (not (RTControleModifFou(TFVierge(Ecran), TobDesc, ExisteInfos))) then
                       finBoucle:=True;
                   if not finBoucle then
                      TobDesc.InsertOrUpdateDB(False) ;
                   end
                 else
                   if (CodeFichier = CodeFichierLigne) then
                   begin
                   for i := 1000 to (1000+(LaTob.ChampsSup.count-1)) do
                      begin
                      FieldNameTo := LaTob.GetNomChamp(i);
                      if copy(FieldNameTo,1,3) <> 'RD8' then continue;
                      if copy(FieldNameTo,5,12) = 'BLOCNOTE' then
                          FieldNameFrom:='RPR_BLOCNOTE'
                      else
                          FieldNameFrom  := 'RPR_RPR' + copy(FieldNameTo,8,length(FieldNameTo)) ;
                      if ( LaTob.GetValue(FieldNameTo)<>TobInfos.GetValue(FieldNameFrom)) then
                        begin
                        LaTob.PutValue(FieldNameTo, TobInfos.GetValue(FieldNameFrom));
                        LaTob.SetString('MODIF_INFOSCOMPL', 'X');
                        LaTob.SetString('VALID_INFOSCOMPL', 'X');
                        end;
                      end;
                   TheTob:=LaTob;
                   end
                   else
                     if (CodeFichier = CodeFichierPiece) then
                       begin
                       for i :=1 to LaTob.NbChamps do
                          begin
                          FieldNameTo := TobDesc.GetNomChamp(i);
                          x:= Pos(copy(FieldNameTo,4,length(FieldNameTo)),ChampsNonTraitesRDA);
                          if x <> 0 then continue;
                          if copy(FieldNameTo,5,12) = 'BLOCNOTE' then
                              FieldNameFrom:='RPR_BLOCNOTE'
                          else
                              FieldNameFrom  := 'RPR_RPR' + copy(FieldNameTo,8,length(FieldNameTo)) ;
                          if ( LaTob.GetValue(FieldNameTo)<>TobInfos.GetValue(FieldNameFrom)) then
                            begin
                            LaTob.PutValue(FieldNameTo, TobInfos.GetValue(FieldNameFrom));
                            LaTob.SetString('MODIF_INFOSCOMPL', 'X');
                            LaTob.SetString('VALID_INFOSCOMPL', 'X');
                            end;
                          end;
                       TheTob:=LaTob;
                       end;
             end
         else finBoucle := true;
         end
     else
         if not finBoucle then
         begin
             for i :=1 to TobType.NbChamps do
                begin
                FieldNameTo := TobType.GetNomChamp(i);
                //x:= Pos(copy(FieldNameTo,5,length(FieldNameTo)),ChampsNonTraitesRPA);
                //if x <> 0 then continue;
                if (copy(FieldNameTo,1,7) <> 'RPA_RPR') and
                     (copy(FieldNameTo,1,12) <> 'RPA_BLOCNOTE') then continue;

                FieldNameFrom  := 'RPR_' + copy(FieldNameTo,5,length(FieldNameTo)) ;
                TobType.PutValue(FieldNameTo, TobInfos.GetValue(FieldNameFrom));
                end;
             TobType.InsertOrUpdateDB(False) ;
         end;
end;

procedure Tof_ParamCL.InfoOblig ( LibChamps:string ; Onglet:integer ; Panel:integer );
var st,Name : string;
begin
    Name:='ONGLETLIBRE'+IntToStr(Onglet);
    SetActiveTabSheet(Name) ;
    st:=Format('Le champ %s est obligatoire',[LibChamps]);
    PGIInfo(st,'');
end;

procedure Tof_ParamCL.MajParametrage;
var st,lib,lib2, CodeCombo,ValDefaut : string;
    P : TPanel;
    ListeActive : TScrollBox;
    OngletActif,i,NoPanel,NoOnglet,j : Integer;
    TS: TTabSheet;
    C:TControl;
begin
     // on detruit tous les parametres
     if CodeFichier=CodeFichierProspect then
        begin
        ExecuteSQL('DELETE FROM CHAMPSPRO') ;
        CodeCombo:='RLZ';
        end
     else
        begin
        ExecuteSQL('DELETE FROM RTINFOSDESC where RDE_DESC="'+CodeFichier+'"') ;
        CodeCombo:='R'+CodeFichier+'Z';
        end ;

     // Balayage des Onglets
     NoOnglet:=0;
     for OngletActif := 1 to TPageControl(getcontrol('PAGEONGLETS')).PageCount do
     begin
        TS:= TTabSheet(Ecran.FindComponent('ONGLETLIBRE'+IntToStr(OngletActif)));
        if TS=Nil then break;
        if TS.TabVisible <> False then
        begin
           ListeActive:= TScrollBox(GetControl('SCB'+IntToStr(OngletActif)));
           if ListeActive.ControlCount = 0 then Inc(NoOnglet)
           else
           For i := 0 to ListeActive.ControlCount - 1 do
           begin
              if i = 0 then
              begin
                 Inc(NoOnglet);
                 //NoPanel:=0;
              end;
              if ListeActive.Controls[i] is TPanel then
                begin
                //Inc(NoPanel);
                P := TPanel(ListeActive.Controls[i]);
                St:= P.Name;
                NoPanel:=P.Tag;
                if Pos('PCHAMPS_',St)<=0 then exit;
                CurLig:=StrToInt(Copy(St,pos('_',St)+1,Length(St)-pos('_',St)));
                if Curlig > PosSeparateur then
                begin
                   St:=P.Caption;
                // test hauteur du panel pour savoir si c'est un espace ou une ligne
                   if P.Height > 6 then
                      Curlig:=PosEspace
                   else
                      Curlig:=PosLigne;
                end
                else
                    begin
                    St:=IntToStr(P.Height);
                    ValDefaut:='';
                    // valeur par défaut sauf bloc-notes
                    if LBChamps.Cells[0,Curlig] <> 'BLO' then
                        begin
                        if copy(LBChamps.Cells[0,Curlig],1,1) = 'M' then
                            begin  // choix multiple
                            ChargeChampMulti (TListBox(GetControl(LBChamps.Cells[0,Curlig])),ValDefaut,false,false);
                            St:=St+';'+ValDefaut;
                            end
                        else
                            begin
                            ValDefaut:=GetControlText(LBChamps.Cells[0,Curlig]);
                            St:=St+';'+ValDefaut;
                            end;
                        end;
                    end;


                if CodeFichier=CodeFichierProspect then
                   ExecuteSQL('INSERT INTO CHAMPSPRO (RCL_ONGLET,RCL_PANEL,RCL_CODE'+
                      ',RCL_OBLIGATOIRE,RCL_LIBELLE,RCL_NOMCHAMP,RCL_TYPECHAMP'+
                      ',RCL_CRITERESEL,RCL_TYPETEXTE,RCL_FILTRE,RCL_CHPORIGINE) '+
                      ' VALUES '+
                      '("'+IntToStr(NoOnglet)+'","'+IntToStr(NoPanel)+'","'+LBChamps.Cells[0,Curlig]+
                      '","'+LBChamps.Cells[5,Curlig]+'","'+St+'","'+LBChamps.Cells[3,Curlig]+
                      '","'+LBChamps.Cells[1,Curlig]+'","'+LBChamps.Cells[6,Curlig]+
                      '","'+LBChamps.Cells[7,Curlig]+'","'+LBChamps.Cells[8,Curlig]+'","'+
                      LBChamps.Cells[10,Curlig]+'")')
                else
                   ExecuteSQL('INSERT INTO RTINFOSDESC (RDE_DESC,RDE_ONGLET,RDE_PANEL,RDE_CODE'+
                      ',RDE_OBLIGATOIRE,RDE_LIBELLE,RDE_NOMCHAMP,RDE_TYPECHAMP'+
                      ',RDE_CRITERESEL,RDE_TYPETEXTE,RDE_FILTRE,RDE_CHPORIGINE) '+
                      ' VALUES '+
                      '("'+CodeFichier+'","'+IntToStr(NoOnglet)+'","'+IntToStr(NoPanel)+'","'+LBChamps.Cells[0,Curlig]+
                      '","'+LBChamps.Cells[5,Curlig]+'","'+St+'","'+LBChamps.Cells[3,Curlig]+
                      '","'+LBChamps.Cells[1,Curlig]+'","'+LBChamps.Cells[6,Curlig]+
                      '","'+LBChamps.Cells[7,Curlig]+'","'+LBChamps.Cells[8,Curlig]
                      +'","'+LBChamps.Cells[10,Curlig]+'")') ;

                if Curlig < PosSeparateur then
                    begin
                    For j:=0 to P.ControlCount-1 do
                        begin
                          C:=P.Controls[j];
                          if ( C is TLabel ) or ( C is TCheckBox ) then
                          begin
                          if C is TLabel then
                                Lib:=THLabel(C).Caption
                          else
                             if C is TCheckBox then
                                Lib:=TCheckBox(C).Caption ;
                          Lib2:=copy(Lib,1,17);
                          updateTablette (CheckdblQuote(Lib),CheckdblQuote(Lib2),CodeCombo,LBChamps.Cells[0,Curlig]);
                          end;
                        end;
                    end;
                end;
           end;
    // 1 : Nom du Champ, 2 : Type du Champ, 3 : Valeur paramétrée 4: Intitulé Onglet
    // 10 = NbOngletMax
    //InfosOnglets: array[1..10,1..4] of string;
        if CodeFichier=CodeFichierProspect then
            ExecuteSQL('INSERT INTO CHAMPSPRO (RCL_ONGLET,RCL_PANEL'+
                ',RCL_LIBELLE,RCL_NOMCHAMP,RCL_TYPECHAMP,RCL_NOMONGLET) '+
                ' VALUES ('+IntToStr(NoOnglet)+',99,"'+InfosOnglets[OngletActif,3]+
                '","'+InfosOnglets[OngletActif,1]+'","'+
                InfosOnglets[OngletActif,2]+'","'+CheckdblQuote(InfosOnglets[OngletActif,4])+'")')
        else
            ExecuteSQL('INSERT INTO RTINFOSDESC (RDE_DESC,RDE_ONGLET,RDE_PANEL'+
                ',RDE_LIBELLE,RDE_NOMCHAMP,RDE_TYPECHAMP,RDE_NOMONGLET) '+
                ' VALUES ("'+CodeFichier+'",'+IntToStr(NoOnglet)+',99,"'+InfosOnglets[OngletActif,3]+
                '","'+InfosOnglets[OngletActif,1]+'","'+
                InfosOnglets[OngletActif,2]+'","'+CheckdblQuote(InfosOnglets[OngletActif,4])+'")')
        end;
     end;
     // maj tablette libellés champs libres
     for i := 2 to ( Nbpost-1 ) do
         begin
         if LBChamps.RowHeights[Nbpost] <> 0 then
            updateTablette (CheckdblQuote(LBChamps.Cells[2,i]),'',CodeCombo,LBChamps.Cells[0,i]);
         end;
end;

procedure Tof_ParamCL.updateTablette (stLibelle, stAbrege, stType, stCode : string);
var sql,stTable : string;
    Q : TQuery;
begin
  Q:=OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE = "'+TabletteLibelle+'"',True,-1,'DESHARE');
  if Q.Eof then
    stTable := 'CHOIXCOD'
  else
    stTable := GetBase(Q.Findfield('DS_NOMBASE').AsString,'CHOIXCOD');
 Ferme (Q);
 sql:='UPDATE '+stTable+' SET CC_LIBELLE = "'+stLibelle+'"';
 if stAbrege<>'' then
   sql:=sql+', CC_ABREGE = "'+stAbrege+'"';
 sql:=sql+' WHERE CC_TYPE = "'+stType+'" and CC_CODE = "'+stCode+ '"' ;

 ExecuteSQL(sql) ;
end;

procedure Tof_ParamCL.InitForme;
var St : string;
    P : TPanel;
    ListeActive : TScrollBox;
begin
ChargeListe;
ChargeParam;
St:='';

if ParamChamps = 0 then
     begin
        if OrigineFichier or OrigineParam then
           begin
            PGIInfo('Le paramétrage de cette saisie n''a pas été effectué');
            exit;
           end
        else
           // Création de la ScrollBox dans le premier Onglet si aucun paramétrage enregistré
           CreerOnglet (1);
     end
   else
      begin
           ListeActive:= TScrollBox(GetControl('SCB1'));
           if ListeActive <> Nil then
               begin
               P:=TPanel(ListeActive.Controls[0]);
               St:=P.Name;
               if Pos('PCHAMPS_',St)<=0 then exit;
               end;
           if St<>'' then
              CurLig:=StrToInt(Copy(St,pos('_',St)+1,Length(St)-pos('_',St)));
           //P.Color:=clTeal
      end;

//SetActiveTabSheet('ONGLETLIBRE1') ;
    if PremierOnglet <> 0 then
       begin
       St:='ONGLETLIBRE'+IntToStr(PremierOnglet);
       SetActiveTabSheet(St) ;
       end
    else
       PGIInfo('Aucun onglet n''est affichable pour cette fiche,#13#10cf les conditions d''affichage des onglets d''infos complémentaires','');

if OrigineFichier or OrigineParam then
   begin
   SetControlVisible('LBChamps',FALSE) ;
   TFVierge(ecran).width:=GetParamsocSecur(ParamSocLargeur,300) ;
   TFVierge(ecran).height:=GetParamsocSecur(ParamSocHauteur,300) ;
   end;
end;

procedure Tof_ParamCL.ChargeParam ;
Var
    R,NoOnglet,nb,i : integer;
    Q : TQuery;
    NonGrise : Boolean;
    Code : string;
    TS: TTabSheet;
    TobFille,TobChampsProFille : TOB;
BEGIN
    TobEcran:=tob.create('tob ecran',Nil,-1);
    SavTobEcran:=tob.create('Sauv tob ecran',Nil,-1);
    //TobChampsProFille:=VH_RT.TobChampsPro.detail[StrToInt(CodeFichier)];
    TobChampsProFille:=VH_RT.TobChampsPro.FindFirst(['CO_CODE'], [CodeFichier], TRUE);
    if Assigned(TobChampsProFille) then
      for i := 0 to TobChampsProFille.Detail.Count-1 do
       begin
          TobFille := TobChampsProFille.Detail[i];
          ParamChamps:=1;
          Code:=TobFille.GetValue('RDE_CODE');
          NoOnglet:=TobFille.GetValue('RDE_ONGLET');
          if TobFille.GetValue('RDE_PANEL') <> 99 then
              begin
              // Recherche du code dans la liste
              NonGrise:=True;
              if ( code <> 'LIG' ) and ( code <> 'ESP' ) then
                  begin
                      for R := 0 to (Nbpost-1) do
                      begin
                         if LBChamps.Cells[0,R] = Code then
                         begin
                         LBChamps.RowHeights[R]:=0;
                         LBChamps.Cells[5,R]:=TobFille.GetValue('RDE_OBLIGATOIRE');
                         LBChamps.Cells[6,R]:=TobFille.GetValue('RDE_CRITERESEL');
                         LBChamps.Cells[7,R]:=TobFille.GetValue('RDE_TYPETEXTE');
                         LBChamps.Cells[8,R]:=TobFille.GetValue('RDE_FILTRE');
                         LBChamps.Cells[10,R]:=TobFille.GetValue('RDE_CHPORIGINE');
                         break;
                         end;
                      end;
                  end
              else
                if code = 'LIG' then R:=PosLigne
                                else R:=PosEspace;
              if ( R <> Nbpost ) then
                  begin
                  // Création de l'Onglet s'il n'existe pas
                  nb:=TPageControl(getcontrol('PAGEONGLETS')).PageCount;
                  if NoOnglet > nb then
                     CreerOnglet (NoOnglet);
                  CreerPanel (NoOnglet,R,LBChamps.Cells[1,R],False,TobFille.GetValue('RDE_LIBELLE'),NonGrise);
                  CreerChampTob(LBChamps.Cells[0,R]);
                  end;
              end
          else
              begin
              InfosOnglets[NoOnglet,1]:=TobFille.GetValue('RDE_NOMCHAMP');
              InfosOnglets[NoOnglet,2]:=TobFille.GetValue('RDE_TYPECHAMP');
              InfosOnglets[NoOnglet,3]:=TobFille.GetValue('RDE_LIBELLE');
              InfosOnglets[NoOnglet,4]:=TobFille.GetValue('RDE_NOMONGLET');

              TS:= TTabSheet(Ecran.FindComponent('ONGLETLIBRE'+IntToStr(NoOnglet)));
              if (TS<>Nil) and (InfosOnglets[NoOnglet,4] <> '') then
                 TS.Caption:= InfosOnglets[NoOnglet,4];
              // si condition d'affichage sur champs table ACTIONS, alim TobAction
              if (CodeFichier = CodeFichierAction) and (OrigineFichier) and (TypeAction = '') then
                  begin
                  if ( copy(InfosOnglets[NoOnglet,1],1,3) = 'RAC' ) and ( TobAction = Nil ) then
                      begin
                      TobAction:=TOB.create ('ACTIONS',NIL,-1);
                      // purée faut suivre ..
                      Q:=OpenSQL('Select * from ACTIONS where RAC_AUXILIAIRE="'+copy(CodeInfos,1,(Pos(';',CodeInfos)-1))+'" and RAC_NUMACTION='+copy(CodeInfos,(Pos(';',CodeInfos)+1),length(CodeInfos)),True);
                      TobAction.selectDB ('',Q,False);
                      Ferme (Q);
                      end;
                  end;
              end;
       end;
    if (OrigineFichier) {mng 29/01/2003and (TypeAction = '') }then
       TestAffichageOnglets;
TobEcran.GetEcran(TFVierge(ecran),Nil);
SavTobEcran.GetEcran(TFVierge(ecran),Nil);

end;

function Tof_ParamCL.ExisteOngletChampOblig : boolean ;
Var TobChampsProFille,tobOblig : tob;
    TS : TTabSheet; 
begin
  result:=false;
  TobChampsProFille:=VH_RT.TobChampsPro.FindFirst(['CO_CODE'], [CodeFichier], TRUE);
  tobOblig:=TobChampsProFille.FindFirst(['RDE_DESC','RDE_OBLIGATOIRE'],[CodeFichier,'X'],TRUE);
  while Assigned(tobOblig) do
    begin
    if tobOblig.GetString('RDE_ONGLET') <> '' then
      begin
      TS:= TTabSheet(Ecran.FindComponent('ONGLETLIBRE'+tobOblig.GetString('RDE_ONGLET')));
      if Assigned(TS) and (TS.TabVisible)= true then
        begin
        result:=true;
        break;
        end;
      end;
    tobOblig:=TobChampsProFille.FindNext(['RDE_DESC','RDE_OBLIGATOIRE'],[CodeFichier,'X'],TRUE);
    end;
end;
procedure Tof_ParamCL.ChargeListe ;
Var
    R,i : integer;
BEGIN
// boucle sur les libellés des champs libres de la table Prospect

   NumSeparateur:=PosSeparateur;
   R:=PosLigne;
   for i:= 0 to ( LBChamps.ColCount - 1 ) do
      if i <> 2 then
         LBChamps.ColWidths[i]:=0;

    LBChamps.RowCount:=R+1; LBChamps.Row:=R;
    LBChamps.Cells[0,R]:='LIG';
    LBChamps.Cells[1,R]:='LIGNE';
    LBChamps.Cells[2,R]:=TraduireMemoire('Ligne de séparation');
    LBChamps.Cells[9,R]:=LBChamps.Cells[2,R];
    R:=PosEspace;
    LBChamps.RowCount:=R+1; LBChamps.Row:=R;
    LBChamps.Cells[0,R]:='ESP';
    LBChamps.Cells[1,R]:='ESPACE';
    LBChamps.Cells[2,R]:=TraduireMemoire('Espace de séparation');
    LBChamps.Cells[9,R]:=LBChamps.Cells[2,R];
    Inc(R) ;

   For i:=1 to MaxCombo do
        begin
        LBChamps.RowCount:=R+1; LBChamps.Row:=R;
        if ( i < 11 ) then
           LBChamps.Cells[0,R]:='CL'+IntToStr(i-1)
        else
           LBChamps.Cells[0,R]:='CL'+chr(i-1+55);

        LBChamps.Cells[1,R]:='COMBO';
        if ( i < 11 ) then
           LBChamps.Cells[2,R]:=RechDom(TabletteLibelle,'CL'+IntToStr(i-1),FALSE)
        else
           LBChamps.Cells[2,R]:=RechDom(TabletteLibelle,'CL'+Chr(i-1+55),FALSE);

        LBChamps.Cells[3,R]:=Prefixe+'_RPRLIBTABLE'+IntToStr(i-1);
        LBChamps.Cells[4,R]:=TabletteCombo+IntToStr(i-1);
        LBChamps.Cells[5,R]:='-';
        LBChamps.Cells[6,R]:='-';
        LBChamps.Cells[7,R]:='';
        LBChamps.Cells[9,R]:=LBChamps.Cells[2,R];
        LBChamps.Cells[10,R]:='';
        Inc(R) ;
        end;
  For i:=1 to MaxMultiChoix do
        begin
        if i = 1 then PremierMul:=R;
        LBChamps.RowCount:=R+1; LBChamps.Row:=R;

        if ( i < 11 ) then
           LBChamps.Cells[0,R]:='ML'+IntToStr(i-1)
        else
           LBChamps.Cells[0,R]:='ML'+chr(i-1+55);

        if ( i < 11 ) then
           LBChamps.Cells[2,R]:=RechDom(TabletteLibelle,'ML'+IntToStr(i-1),FALSE)
        else
           LBChamps.Cells[2,R]:=RechDom(TabletteLibelle,'ML'+Chr(i-1+55),FALSE);

        LBChamps.Cells[1,R]:='MULTI';
        LBChamps.Cells[3,R]:=Prefixe+'_RPRLIBMUL'+IntToStr(i-1);
        LBChamps.Cells[4,R]:=TabletteComboMul+IntToStr(i-1);
        LBChamps.Cells[5,R]:='-';
        LBChamps.Cells[6,R]:='-';
        LBChamps.Cells[7,R]:='';
        LBChamps.Cells[9,R]:=LBChamps.Cells[2,R];
        LBChamps.Cells[10,R]:='';
        Inc(R) ;
        end;
  For i:=1 to MaxTexte do
        begin
        LBChamps.RowCount:=R+1; LBChamps.Row:=R;
        if ( i < 11 ) then
           LBChamps.Cells[0,R]:='TL'+IntToStr(i-1)
        else
           LBChamps.Cells[0,R]:='TL'+chr(i-1+55);

        if ( i < 11 ) then
           LBChamps.Cells[2,R]:=RechDom(TabletteLibelle,'TL'+IntToStr(i-1),FALSE)
        else
           LBChamps.Cells[2,R]:=RechDom(TabletteLibelle,'TL'+Chr(i-1+55),FALSE);

        LBChamps.Cells[1,R]:='EDIT';
        LBChamps.Cells[3,R]:=Prefixe+'_RPRLIBTEXTE'+IntToStr(i-1);
        LBChamps.Cells[5,R]:='-';
        LBChamps.Cells[6,R]:='-';
        LBChamps.Cells[7,R]:='';
        LBChamps.Cells[9,R]:=LBChamps.Cells[2,R];
        LBChamps.Cells[10,R]:='';
        Inc(R) ;
        end;
  For i:=1 to MaxValeur do
        begin
        LBChamps.RowCount:=R+1; LBChamps.Row:=R;
        if ( i < 11 ) then
           LBChamps.Cells[0,R]:='VL'+IntToStr(i-1)
        else
           LBChamps.Cells[0,R]:='VL'+chr(i-1+55);

        if ( i < 11 ) then
           LBChamps.Cells[2,R]:=RechDom(TabletteLibelle,'VL'+IntToStr(i-1),FALSE)
        else
           LBChamps.Cells[2,R]:=RechDom(TabletteLibelle,'VL'+Chr(i-1+55),FALSE);

        LBChamps.Cells[1,R]:='DOUBLE';
        LBChamps.Cells[3,R]:=Prefixe+'_RPRLIBVAL'+IntToStr(i-1);
        LBChamps.Cells[5,R]:='-';
        LBChamps.Cells[6,R]:='-';
        LBChamps.Cells[7,R]:='';
        LBChamps.Cells[9,R]:=LBChamps.Cells[2,R];
        LBChamps.Cells[10,R]:='';
        Inc(R) ;
        end;
  For i:=1 to MaxBool do
        begin
        LBChamps.RowCount:=R+1; LBChamps.Row:=R;
        if ( i < 11 ) then
           LBChamps.Cells[0,R]:='BL'+IntToStr(i-1)
        else
           LBChamps.Cells[0,R]:='BL'+chr(i-1+55);

        if ( i < 11 ) then
           LBChamps.Cells[2,R]:=RechDom(TabletteLibelle,'BL'+IntToStr(i-1),FALSE)
        else
           LBChamps.Cells[2,R]:=RechDom(TabletteLibelle,'BL'+Chr(i-1+55),FALSE);

        LBChamps.Cells[1,R]:='BOOLEAN';
        LBChamps.Cells[3,R]:=Prefixe+'_RPRLIBBOOL'+IntToStr(i-1);
        LBChamps.Cells[5,R]:='-';
        LBChamps.Cells[6,R]:='-';
        LBChamps.Cells[7,R]:='';
        LBChamps.Cells[9,R]:=LBChamps.Cells[2,R];
        LBChamps.Cells[10,R]:='';
        Inc(R) ;
        end;
  For i:=1 to MaxDate do
        begin
        LBChamps.RowCount:=R+1; LBChamps.Row:=R;
        if ( i < 11 ) then
           LBChamps.Cells[0,R]:='DL'+IntToStr(i-1)
        else
           LBChamps.Cells[0,R]:='DL'+chr(i-1+55);

        if ( i < 11 ) then
           LBChamps.Cells[2,R]:=RechDom(TabletteLibelle,'DL'+IntToStr(i-1),FALSE)
        else
           LBChamps.Cells[2,R]:=RechDom(TabletteLibelle,'DL'+Chr(i-1+55),FALSE);

        LBChamps.Cells[1,R]:='DATE';
        LBChamps.Cells[3,R]:=Prefixe+'_RPRLIBDATE'+IntToStr(i-1);
        LBChamps.Cells[5,R]:='-';
        LBChamps.Cells[6,R]:='-';
        LBChamps.Cells[7,R]:='';
        LBChamps.Cells[9,R]:=LBChamps.Cells[2,R];
        LBChamps.Cells[10,R]:='';
        Inc(R) ;
        end;
// 1 bloc-notes
  LBChamps.RowCount:=R+1; LBChamps.Row:=R;
  LBChamps.Cells[0,R]:='BLO';
  LBChamps.Cells[1,R]:='BLOB';
  LBChamps.Cells[2,R]:=RechDom(TabletteLibelle,'BLO',FALSE);
  LBChamps.Cells[3,R]:=Prefixe+'_BLOCNOTE';
  LBChamps.Cells[5,R]:='-';
  LBChamps.Cells[6,R]:='-';
  LBChamps.Cells[7,R]:='';
  LBChamps.Cells[9,R]:=LBChamps.Cells[2,R];
  LBChamps.Cells[10,R]:='';

TFVierge(Ecran).Hmtrad.ResizeGridColumns(LBChamps) ;
LBChamps.Row:=0; LBChamps.SetFocus ;
END ;


procedure Tof_ParamCL.AjouterChamps();
var TC,St : String;
    NoOnglet,TagOrigine,i : integer;
    P : TPanel ;
begin
if CurLig<0 then Exit ;

P:= TPanel(Ecran.FindComponent('PCHAMPS_'+IntToStr(CurLig)));
if P=Nil then TagOrigine:=0
else
    TagOrigine:=P.Tag;

CurLig:=LBChamps.Row ;

if (LBChamps.RowHeights[CurLig]=0) then exit;

TC:=LBChamps.Cells[1,CurLig] ;

if TC='' then Exit ;
NoOnglet:=TPageControl(getControl( 'PAGEONGLETS')).ActivePage.PageIndex+1;
CreerPanel (NoOnglet,Curlig,TC,True,'',True);

P:= TPanel(Ecran.FindComponent(DernierPanel));
if P=Nil then exit;
St:=P.Name;
CurLig:=StrToInt(Copy(St,pos('_',St)+1,Length(St)-pos('_',St)));



for i := 1 to ( P.Tag - TagOrigine - 1 ) do
    DescendreChamps();
BougeChamp:=True;
end;

procedure Tof_ParamCL.EnleverChamps();
var P : TPanel;
    T : TTabOrder;
    i,j,Pos : integer;
    OngletActif : Integer ;
    ListeActive : TScrollBox ;
begin
OngletActif:=TPageControl(getControl( 'PAGEONGLETS')).ActivePage.PageIndex+1;
ListeActive:= TScrollBox(GetControl('SCB'+IntToStr(OngletActif)));

P:= TPanel(Ecran.FindComponent('PCHAMPS_'+IntToStr(CurLig)));
if P=Nil then exit;
T:=P.TabOrder;
Pos:=P.Tag;
while P.ControlCount>0 do P.Controls[0].Free;
P.Free;
if Curlig < PosSeparateur then
   LBChamps.RowHeights[CurLig]:=LBChamps.DefaultRowHeight;
for i:=0 to ListeActive.ControlCount-1 do
    if ListeActive.Controls[i] is TPanel then
    begin
    if TPanel(ListeActive.Controls[i]).TabOrder=T-1 then TPanel(ListeActive.Controls[i]).Color:=clTeal
                                             else TPanel(ListeActive.Controls[i]).Color:=CouleurPanel;
    end;
// mise à jour des tags
for i:=Pos+1 to ListeActive.ControlCount do
    for j:=0 to ListeActive.ControlCount-1 do
        begin
        if TPanel(ListeActive.Controls[j]).Tag = i then
           begin
           TPanel(ListeActive.Controls[j]).Tag:=i-1;
           break;
           end;
        end;
BougeChamp:=True;        
end;

procedure Tof_ParamCL.Panels_OnEnter(Sender: TObject);
var St     : string;
    i,IP : integer;
    P      : TPanel;
    OngletActif : Integer ;
    ListeActive : TScrollBox ;
begin
OngletActif:=TPageControl(getControl( 'PAGEONGLETS')).ActivePage.PageIndex+1;
ListeActive:= TScrollBox(GetControl('SCB'+IntToStr(OngletActif)));

St:=TComponent(Sender).Name; if Pos('PCHAMPS_',St)<=0 then exit;
CurLig:=StrToInt(Copy(St,pos('_',St)+1,Length(St)-pos('_',St)));
for i:=0 to ListeActive.ControlCount-1 do
   if ((ListeActive.Controls[i] is TPanel)) then
      BEGIN
      P:=TPanel(ListeActive.Controls[i]); St:=P.Name;
      if Pos('PCHAMPS_',St)<=0 then break;
      IP:=StrToInt(Copy(St,pos('_',St)+1,Length(St)-pos('_',St)));
      if (IP=CurLig) and (not OrigineFichier) and (not OrigineParam) then
         BEGIN
         P.Color:=clTeal;
         END else
             P.Color:=CouleurPanel;
      END;
end;

Procedure Tof_ParamCL.MonterChamps;
var
    i,IP : integer;
    P,P1 : TPanel;
    OngletActif : Integer ;
    ListeActive : TScrollBox ;
begin
OngletActif:=TPageControl(getControl( 'PAGEONGLETS')).ActivePage.PageIndex+1;
ListeActive:= TScrollBox(GetControl('SCB'+IntToStr(OngletActif)));
P1:=Nil;
P:= TPanel(Ecran.FindComponent('PCHAMPS_'+IntToStr(CurLig)));
if P=Nil then exit;
if P.Tag = ListeActive.ControlCount-1 then exit;

// recherche du panel ayant tag+1
for i:=0 to ListeActive.ControlCount-1 do
begin
   if ((ListeActive.Controls[i] is TPanel)) then
      begin
      P1:=TPanel(ListeActive.Controls[i]);
      if P1.Tag=(P.Tag+1) then break;
      end;
end;
if P1 = Nil then exit;
i:=P.Tag;

IP:=P.Top;
P.Top:=P1.Top;
P1.Top:=IP;
P.Tag:=i+1;
P1.Tag:=i;

ListeActive.Invalidate;
BougeChamp:=True;
end;

Procedure Tof_ParamCL.DescendreChamps;
var
    i,IP : integer;
    P,P1 : TPanel;
    OngletActif : Integer ;
    ListeActive : TScrollBox ;
begin
OngletActif:=TPageControl(getControl( 'PAGEONGLETS')).ActivePage.PageIndex+1;
ListeActive:= TScrollBox(GetControl('SCB'+IntToStr(OngletActif)));
P1:=Nil;
P:= TPanel(Ecran.FindComponent('PCHAMPS_'+IntToStr(CurLig)));
if P=Nil then exit;
if P.Tag = 0 then exit;

// recherche du panel ayant tag+1
for i:=0 to ListeActive.ControlCount-1 do
begin
   if ((ListeActive.Controls[i] is TPanel)) then
      begin
      P1:=TPanel(ListeActive.Controls[i]);
      if P1.Tag=(P.Tag-1) then break;
      end;
end;
if P1 = Nil then exit;
i:=P.Tag;

IP:=P.Top;
P.Top:=P1.Top;
P1.Top:=IP;
P.Tag:=i-1;
P1.Tag:=i;

ListeActive.Invalidate;
BougeChamp:=True;
end;

procedure Tof_ParamCL.NewPosition;
var i, New : integer;
BEGIN
New:=CurLig;
for i:=CurLig to LBChamps.RowCount-2 do
   if LBChamps.RowHeights[i]=0 then New:=New+1 else break;
if LBChamps.RowHeights[New]=0 then
   BEGIN
   New:=CurLig;
   for i:=CurLig downto 1 do
      if LBChamps.RowHeights[i]=0 then New:=New-1 else break;
   END;
if LBChamps.RowHeights[New]=0 then CurLig:=-1 else BEGIN CurLig:=New; LBChamps.Row:=New; END;
END;

procedure Tof_ParamCL.AjouterOnglet();
var i : Integer;
begin
// Création d'un Onglet
    i :=  TPageControl(getcontrol('PAGEONGLETS')).PageCount + 1;
    if ( i > NbOngletMax ) then
       PGIInfo(Format('On ne peut pas créer plus de %s onglets',[IntToStr(NbOngletMax)]),'')
    else
       CreerOnglet (i);
BougeChamp:=True;       
end;

procedure Tof_ParamCL.OngletAGauche();
var TS,TS2 : TTabSheet;
    SauvInfosOnglets: array[1..4] of string;
    i : integer;
    ListeActive: TScrollBox;
begin
  TS :=  TPageControl(getControl( 'PAGEONGLETS')).ActivePage;
  if not Assigned(TS) then exit;
  if TS.PageIndex = 0 then exit; // onglet le plus à droite
  for i:=1 to 4 Do
    SauvInfosOnglets[i]:=InfosOnglets[TS.PageIndex,i];
  for i:=1 to 4 Do
    InfosOnglets[TS.PageIndex,i]:=InfosOnglets[TS.PageIndex+1,i];
  for i:=1 to 4 Do
    InfosOnglets[TS.PageIndex+1,i]:=SauvInfosOnglets[i];

  { il faut renommer les scroll box }
  ListeActive:= TScrollBox(GetControl('SCB'+IntToStr(TS.PageIndex)));
  if Assigned(ListeActive) then
    ListeActive.Name:='SCBX';

  ListeActive:= TScrollBox(GetControl('SCB'+IntToStr(TS.PageIndex+1)));
  if Assigned(ListeActive) then
    ListeActive.Name:='SCB'+IntToStr(TS.PageIndex);

  ListeActive:= TScrollBox(GetControl('SCBX'));
  if Assigned(ListeActive) then
    ListeActive.Name:='SCB'+IntToStr(TS.PageIndex+1);

  { il faut renommer les onglets }
  TS2:= TTabSheet(Ecran.FindComponent('ONGLETLIBRE'+IntToStr(TS.PageIndex)));
  if Assigned(TS2) then
    TS2.Name:='ONGLETLIBREX';

  TS2:= TTabSheet(Ecran.FindComponent('ONGLETLIBRE'+IntToStr(TS.PageIndex+1)));
  if Assigned(TS2) then
    TS2.Name:='ONGLETLIBRE'+IntToStr(TS.PageIndex);

  TS2:= TTabSheet(Ecran.FindComponent('ONGLETLIBREX'));
  if Assigned(TS2) then
    TS2.Name:='ONGLETLIBRE'+IntToStr(TS.PageIndex+1);

  TS.PageIndex:=TS.PageIndex-1;

end;

procedure Tof_ParamCL.OngletADroite();
var TS,TS2 : TTabSheet;
    SauvInfosOnglets: array[1..4] of string;
    i : integer;
    ListeActive: TScrollBox;
begin

  TS :=  TPageControl(getControl( 'PAGEONGLETS')).ActivePage;
  if not Assigned(TS) then exit;
  if TS.PageIndex = Pred(TPageControl(getcontrol('PAGEONGLETS')).PageCount) then exit; // onglet le plus à gauche
  for i:=1 to 4 Do
    SauvInfosOnglets[i]:=InfosOnglets[TS.PageIndex+1,i];
  for i:=1 to 4 Do
    InfosOnglets[TS.PageIndex+1,i]:=InfosOnglets[TS.PageIndex+2,i];
  for i:=1 to 4 Do
    InfosOnglets[TS.PageIndex+2,i]:=SauvInfosOnglets[i];

  { il faut renommer les scroll box }
  ListeActive:= TScrollBox(GetControl('SCB'+IntToStr(TS.PageIndex+1)));
  if Assigned(ListeActive) then
    ListeActive.Name:='SCBX';

  ListeActive:= TScrollBox(GetControl('SCB'+IntToStr(TS.PageIndex+2)));
  if Assigned(ListeActive) then
    ListeActive.Name:='SCB'+IntToStr(TS.PageIndex+1);

  ListeActive:= TScrollBox(GetControl('SCBX'));
  if Assigned(ListeActive) then
    ListeActive.Name:='SCB'+IntToStr(TS.PageIndex+2);

  { il faut renommer les onglets }
  TS2:= TTabSheet(Ecran.FindComponent('ONGLETLIBRE'+IntToStr(TS.PageIndex+1)));
  if Assigned(TS2) then
    TS2.Name:='ONGLETLIBREX';

  TS2:= TTabSheet(Ecran.FindComponent('ONGLETLIBRE'+IntToStr(TS.PageIndex+2)));
  if Assigned(TS2) then
    TS2.Name:='ONGLETLIBRE'+IntToStr(TS.PageIndex+1);

  TS2:= TTabSheet(Ecran.FindComponent('ONGLETLIBREX'));
  if Assigned(TS2) then
    TS2.Name:='ONGLETLIBRE'+IntToStr(TS.PageIndex+2);

  TS.PageIndex:=TS.PageIndex+1;
end;
procedure Tof_ParamCL.EnleverOnglet();
var TS : TTabSheet ;
    OngletActif : Integer ;
    ListeActive : TScrollBox ;
    i : Integer ;
    P : TPanel;
    St : String;
begin
// Suppression des controles du ScrollBox associés à l'Onglet que l'on supprime
   OngletActif:=TPageControl(getControl( 'PAGEONGLETS')).ActivePage.PageIndex+1;

   ListeActive:= TScrollBox(GetControl('SCB'+IntToStr(OngletActif)));

   while ( ListeActive.ControlCount <> 0 ) do
   begin
      i := ListeActive.ControlCount-1 ;
      if ListeActive.Controls[i] is TPanel then
      begin
        P := TPanel(ListeActive.Controls[i]);
        St:= P.Name;
        if Pos('PCHAMPS_',St)<=0 then exit;
        CurLig:=StrToInt(Copy(St,pos('_',St)+1,Length(St)-pos('_',St)));
        P.Free;
        if Curlig < PosSeparateur then
                LBChamps.RowHeights[CurLig]:=LBChamps.DefaultRowHeight;
      end;
   end;
// on ne supprime pas l'Onglet mais on le rend invisible
  TS:= TTabSheet(Ecran.FindComponent('ONGLETLIBRE'+IntToStr(OngletActif)));
  if TS=Nil then exit;
  TS.TabVisible:=False;
BougeChamp:=True;  
end;

procedure Tof_ParamCL.CreerOnglet ( NoOnglet : integer) ;
var TS : TTabSheet ;
    i : Integer;
    SCB : TScrollBox;
begin
// Création d'un Onglet
    TS:=TTabSheet.create(Ecran) ;
    With TS do
    Begin
    PageControl:=TPageControl(getcontrol('PAGEONGLETS')) ;
    i :=  NoOnglet;
    Name:='ONGLETLIBRE'+IntToStr(i);
    Parent:=TPageControl(getcontrol('PAGEONGLETS'));
    ShowHint:=FALSE ;
    //Caption:='Champs Libres '+IntToStr(i);
    //Caption:=RechDom('RTLIBONGLET','ON'+IntToStr(i-1),FALSE);
    Visible:=TRUE ;
    ParentFont:=true;
    SetActiveTabSheet(Name) ;
    End ;
// Création de la ScrollBox associée
    SCB:=TScrollBox.create(Ecran) ;
    SCB.Parent:=TTabSheet(getcontrol('ONGLETLIBRE'+IntToStr(i))) ;
    SCB.Name:='SCB'+IntToStr(i);
    SCB.ParentFont:=False; SCB.Font.Color:=clWindowText;
    SCB.Align:=alClient;
end;

procedure Tof_ParamCL.CreerPanel ( OngletActif:integer; Curlig : integer; TC:String; NewPos : Boolean; Libelle : string; NonGrise : Boolean) ;
var P  : TPanel;
    L,LT  : TLabel;
    CC : TCheckBox;
    C  : THValComboBox;
    //MC : THMultiValComboBox;
    MC : TListBox;
    E  : THEdit;
    N  : THNumEdit;
    OkCase,OkCombo,OkEdit,OkMulti,OkNumEdit,OkSpace,OkLigne,OkDate,OkBlob : boolean;
    i  : integer;
    stMess,boobool,NomCode,ValChamp,Requete,Titre : string;
    ListeActive : TScrollBox ;
    B,B2 : TToolbarButton97;
    Q : TQuery;
    BB : THRichEditOle ;
    sl : HTStringList;
begin

Titre:=ReadToKenSt(Libelle);

OkCase:=(TC='BOOLEAN');
OkCombo:=(TC='COMBO');
OkMulti:=(TC='MULTI');
OkNumEdit:=(TC='DOUBLE');
OkLigne:=(TC='LIGNE');
OkSpace:=(TC='ESPACE');
OkDate:= (TC='DATE');
OkBlob:= (TC='BLOB');

OkEdit:=((TC<>'BOOLEAN')And(TC<>'COMBO')And(TC<>'MULTI')And(TC<>'BLOB')
     And(TC<>'DOUBLE')And(TC<>'ESPACE')And(TC<>'LIGNE')And(TC<>'DATE'));

if ((Not OkCase)And(Not OkCombo)And(Not OkEdit)And (Not OkBlob)And
    (Not OkMulti)And(Not OkNumEdit)And(Not OkLigne)And(Not OkSpace)And(Not OkDate)) then Exit;

ListeActive:= TScrollBox(GetControl('SCB'+IntToStr(OngletActif)));

P:=TPanel.Create(Ecran);
P.Parent:=ListeActive;
P.ParentFont:=False; P.Font.Color:=clWindowText;
P.ParentColor:=true;
P.Width:=TFVierge(ecran).width;
if CouleurPanel=clAqua then CouleurPanel:=P.Color;
MC:=Nil;
if OkLigne then
begin
   P.BevelOuter:=bvRaised;
   P.Height:=HauteurLigneSeule;
end
else
begin
   P.BevelOuter:=bvNone;
   if OkMulti then
   begin
        MC:=TListBox.Create(Ecran);
        if (not NewPos) and (Titre <> '') then
           begin
           P.Height:=StrToInt(Titre);
           MC.Tag:=P.Height div HauteurLigneListe;
           end
        else
           begin
           P.Height:=NbLigneMulti*HauteurLigneListe;
           // ajout mng 15-11-01
           MC.Tag :=NbLigneMulti;
           end;
   end
   else
      P.Height:=HauteurPanel;
end;
P.Align:=alTop; P.BevelInner:=bvNone ;
P.OnEnter:=Panels_OnEnter; P.OnClick:=Panels_OnEnter;

P.Tag:=ListeActive.ControlCount - 1;

if ( not OkSpace ) and ( not OkLigne ) then
begin
  if Assigned(TPanel(Ecran.FindComponent('PCHAMPS_'+IntToStr(CurLig)))) then
    begin
    stMess:='Ce composant existe déjà : PCHAMPS_'+IntToStr(CurLig);
    PgiInfo(stMess, 'Composant existant');
    exit;
    end;
  P.Name:='PCHAMPS_'+IntToStr(CurLig);
  if OkCase then
     begin
     CC:=TCheckBox.Create(Ecran); CC.Parent:=P; CC.Name:=LBChamps.Cells[0,CurLig];
     CC.Left:=8; CC.Width:=164;
     CC.Alignment:= taLeftJustify;
     CC.Top:=4;
     CC.Caption:=LBChamps.Cells[2,CurLig];
     CC.Anchors:=[akLeft];
     CC.Enabled:=NonGrise;
     if ExisteInfos then
       begin
       boobool:=TobInfos.GetValue(LBChamps.Cells[3,CurLig]);
       if boobool = 'X' then
           CC.state:=cbChecked
       else
           CC.state:=cbUnChecked;
       end;
     if OrigineFichier then
         CC.OnExit := ChampExit;
     end
  else
     begin
     L:=THLabel.Create(Ecran); L.Parent:=P; L.Name:='L'+IntToStr(CurLig);
     L.Left:=8; L.Width:=150;
     L.Top:=(P.Height-L.height) div 2;
     L.AutoSize:=false;
     L.Caption:=LBChamps.Cells[2,CurLig];
     L.Anchors:=[akTop,akLeft];
     L.Alignment:=TaLeftJustify;
     L.Enabled:=NonGrise;
     if OkCombo then
        begin
        C:=THValComboBox.Create(Ecran); C.Parent:=P; C.Name:=LBChamps.Cells[0,CurLig];
        C.Left:=L.Left+L.Width+1; C.Width:=P.Width-C.Left-8; C.Top:=4; C.Text:='' ;
        C.Style:=csDropDownList ;
        C.ItemIndex:=0; L.FocusControl:=C;
        C.Anchors:=[akTop,akRight,akLeft];
        //C.Anchors:=[akTop,akLeft];
        C.Enabled:=NonGrise;
        C.Vide:=True;
        C.VideString:=TraduireMemoire('<<Aucun>>');
        C.DataType:=LBChamps.Cells[4,CurLig];
        If GetParamsocSecur ('So_ChampObliAst',false) and (LBChamps.Cells[5,CurLig]='X') then
          begin  //mcd 03/01/2007 13528 GC
          C.obligatory:=true;
          end;
        if ExisteInfos then
           C.value:=TobInfos.GetValue(LBChamps.Cells[3,CurLig]);
        if OrigineFichier then
           C.OnExit := ChampExit;
        end;
     if OkMulti then
        begin
        B:=Nil;
        if OrigineFichier then
            begin
            B:= TToolbarButton97.Create(Ecran);
            B.Parent:=P;
            B.ParentColor:=true;
            B.Name:='B'+LBChamps.Cells[0,CurLig]+'_'+IntToStr(CurLig);
            B.Left:=L.Left+L.Width+1;
            B.Height:=BHauteur;
            B.Top:=L.Top-((B.Height-L.Height) div 2);
            B.Width:=BLargeur;
            B.Flat:=False;
            B.Visible:=True;
            B.Onclick:=SelectionMulti;
            B2:= TToolbarButton97(GetControl('TITREONGLET'));
            B.Glyph:=B2.Glyph;
            B.GlyphMask:=B2.GlyphMask;
            B.Enabled:=NonGrise;
            B.Anchors:=[akLeft];
            end;

        //MC:=TListBox.Create(Ecran);
        MC.Parent:=P;
        MC.Name:=LBChamps.Cells[0,CurLig];
        MC.MultiSelect:=True;
        MC.ExtendedSelect:=False;
        MC.Anchors:=[akTop,akRight,akLeft];
        if OrigineFichier then
           MC.Left:=B.width+1+L.Left+L.Width+1
        else
           MC.Left:=L.Left+L.Width+1;
        MC.Enabled:=NonGrise;
        MC.Width:=P.Width-MC.Left-8; MC.Top:=0;
        MC.Height:=P.Height;
(*        If GetParamsocSecur ('So_ChampObliAst',false) and (LBChamps.Cells[5,CurLig]='X') then
          begin  //mcd 03/01/2007 13528 GC   a remettre quand OK pour  TListeBox(agl 12757)
          mc.obligatory:=true;
          end; *)
        NomCode:='RM'+copy(MC.Name,3,1);
        L.FocusControl:=MC;

        if ExisteInfos then
           ValChamp:=TobInfos.GetValue(LBChamps.Cells[3,CurLig])
        else
           ValChamp:=Libelle;
        LBChamps.Cells[8,Curlig]:=ValChamp;
        //ChargeListBox (TabletteComboMul, MC,( not OrigineFichier) ,NomCode,ValChamp,ExisteInfos);
        ChargeListBox (TabletteComboMul, MC,( not OrigineFichier) ,NomCode,ValChamp,true);
        {if MC.Items.Count > 0 then
           if MC.Items.Count > MC.Tag then
              P.Height:=MC.Tag*HauteurLigneListe
           else
               if MC.Items.Count < NbLignesMini then
                  P.Height:=(NbLignesMini)*HauteurLigneListe
               else
                  P.Height:=(MC.Items.Count)*HauteurLigneListe
        else   }
        if MC.Items.Count = 0 then
           begin
           {P.Height:=HauteurLigneListe*2;}
           MC.Items.Add (TraduireMemoire('<<Aucun>>'));
           end;
        MC.Height:=P.Height;
        L.Top:=(P.Height-L.height) div 2;
        if OrigineFichier then
           B.Top:=L.Top-((B.Height-L.Height) div 2);

        end;
     if OkEdit or OkDate then
        begin
        E:=THEdit.Create(Ecran); E.Parent:=P; E.Name:=LBChamps.Cells[0,CurLig];
        if TC='DATE' then
           begin E.EditMask:=DateMask;
                 E.OpeType := otDate;E.Defaultdate:=od1900 ;
                 E.ControlerDate:=true;
                 end
        else
            E.OpeType := otString;
        E.Left:=L.Left+L.Width+1; E.Width:=P.Width-E.Left-8; E.Top:=4;
        E.Text:=''; L.FocusControl:=E;
        E.Anchors:=[akTop,akRight,akLeft];
        E.MaxLength:=35;
        If GetParamsocSecur ('So_ChampObliAst',false) and (LBChamps.Cells[5,CurLig]='X') then
          begin  //mcd 03/01/2007 13528 GC
          E.obligatory:=true;
          end;
        if ExisteInfos then
              E.text:=TobInfos.GetValue(LBChamps.Cells[3,CurLig])
        else  if TC='DATE' then E.text:=DateToStr(iDate1900) ;
        E.Enabled:=NonGrise;
        //if (TypeAction <> '') and (LBChamps.Cells[7,CurLig] <> '') then E.Enabled:=False
        //else
        if (OkEdit) and (LBChamps.Cells[7,CurLig] <> '' ) then
           begin
           E.MaxLength:=17;
           E.ElipsisButton:=True;
           E.OnElipsisClick := TexteClick;
           if ExisteInfos then
              E.OnExit := TestExiste;
           if E.Width > ( LargeurTexte + 50 ) then
               begin
               E.Width:= LargeurTexte;
               if LBChamps.Cells[7,CurLig] = 'C' then E.Width:= LargeurContact;
               E.Anchors:=[akLeft];
               LT:=THLabel.Create(Ecran); LT.Parent:=P; LT.Name:='LT'+IntToStr(CurLig);
               LT.FocusControl:=E;
               LT.Left:=E.Left+E.Width+5;
               LT.Width:=P.Width-LT.Left-8;
               LT.Top:=L.Top;
               LT.height:=L.height;
               LT.AutoSize:=false;
               //LT.Caption:=LBChamps.Cells[2,CurLig];
               LT.Alignment:=TaLeftJustify;
               LT.Anchors:=[akTop,akRight,akLeft];
               LT.Caption:='';
               LT.Enabled:=NonGrise;
               if (ExisteInfos) and (E.Text <> '') then
                   begin
                   Requete:='';
                   // Tiers
                   if LBChamps.Cells[7,CurLig] = 'T' then
                       begin
                       if GetControlText ('NATUREAUXI')='FOU' then
                         Requete:='Select T_LIBELLE from TIERS where T_TIERS="'+E.text
                            +'" and (T_NATUREAUXI = "FOU")'
                       else
                       {$ifdef GIGI}
                         Requete:='Select T_LIBELLE from TIERS where T_TIERS="'+E.text
                            +'" and (T_NATUREAUXI = "CLI" or T_NATUREAUXI="PRO" or T_NATUREAUXI="NCP" or T_NATUREAUXI="CON")';
                       {$else}
                         Requete:='Select T_LIBELLE from TIERS where T_TIERS="'+E.text
                            +'" and (T_NATUREAUXI = "CLI" or T_NATUREAUXI="PRO" or T_NATUREAUXI="CON")';
                       {$endif}
                       end
                   else
                   // ressource
                   if LBChamps.Cells[7,CurLig] = 'R' then
                      Requete:=RequeteRessource+E.text+'"'
                   else
                     if LBChamps.Cells[7,CurLig] = 'O' then
                        Requete:=RequeteOperation+E.text+'"'
                     else
                       if LBChamps.Cells[7,CurLig] = 'P' then
                          Requete:=RequeteProjet+E.text+'"'
                       else
                     // Contact
                        if ( LBChamps.Cells[7,CurLig] = 'C' ) and ( IsNumeric(E.Text) ) then
                            begin
                            if CodeFichier = CodeFichierProspect then
                               Requete:='Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'
                                  +CodeInfos+'" AND C_NUMEROCONTACT='+E.text
                            else
                               if CodeFichier = CodeFichierAction then
                                 Requete:='Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'
                                  +copy(CodeInfos,1,(Pos(';',CodeInfos)-1))+'" AND C_NUMEROCONTACT='+E.text
                               else
                                 if (CodeFichier = CodeFichierParc) and (Assigned(TobParc)) then
                                   begin
                                   if TobParc.GetString('WPC_TIERS') <> '' then
                                     Requete:='Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'
                                        +TiersAuxiliaire(TobParc.GetString('WPC_TIERS'),false)+'" AND C_NUMEROCONTACT='+E.text;
                                   end
                                 else
                                 if (CodeFichier = CodeFichierProposition) and (Assigned(TobProposition)) then
                                   begin
                                   if TobProposition.GetString('RPE_TIERS') <> '' then
                                     Requete:='Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'
                                        +TiersAuxiliaire(TobProposition.GetString('RPE_TIERS'),false)+'" AND C_NUMEROCONTACT='+E.text;
                                   end
                                 else
                                   if (CodeFichier = CodeFichierPiece) then
                                      Requete:='Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'
                                        +TiersAuxiliaire(TobPiece.GetString('GP_TIERS'),false)+'" AND C_NUMEROCONTACT='+E.text
                                   else
                                     Requete:='Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'
                                        +copy(CodeInfos,3,(pos(';',copy(CodeInfos,3,length(CodeInfos))))-1)+'" AND C_NUMEROCONTACT='+E.text
                            end
                        else E.Text:='';
                   if ( Requete<>'' ) then
                       begin
                       Q := OpenSQL(Requete, True);
                       if not Q.Eof then
                           begin
                           if LBChamps.Cells[7,CurLig] = 'T' then
                              LT.Caption:=Q.FindField('T_LIBELLE').AsString
                           else
                              if LBChamps.Cells[7,CurLig] = 'R' then
                                 LT.Caption:=Q.FindField('ARS_LIBELLE').AsString
                              else
                                if LBChamps.Cells[7,CurLig] = 'O' then
                                   LT.Caption:=Q.FindField('ROP_LIBELLE').AsString
                                else
                                  if LBChamps.Cells[7,CurLig] = 'P' then
                                     LT.Caption:=Q.FindField('RPJ_LIBELLE').AsString
                                  else
                                     LT.Caption:=Q.FindField('C_NOM').AsString;
                           end;
                       Ferme(Q) ;
                       end;
                   end;
               end;
           end;
        end;
     if OkNumEdit then
        begin
        N:=THNumEdit.Create(Ecran); N.Parent:=P; N.Name:=LBChamps.Cells[0,CurLig];
        N.Left:=L.Left+L.Width+1; N.Width:=P.Width-N.Left-8; N.Top:=4;
        N.Text:=''; L.FocusControl:=N;
        N.Anchors:=[akTop,akRight,akLeft];
        N.Decimals:=2;
        N.Digits:=12;
        N.UseRounding:=True;
        N.Masks.PositiveMask:='#,##0.00';
        N.AutoSize:=True;
        N.BorderStyle:=bsSingle;
        N.Cursor:=crDefault;
        N.Debit:=False;
        N.DragCursor:=crDrag;
        N.DragMode:=dmManual;
        N.Enabled:=NonGrise;
        N.HideSelection:=True;
        N.ParentColor:=True;
        N.ParentCtl3D:=True;
        N.ParentFont:=False;
        N.ParentShowHint:=True;
        N.TabStop:=True;
        N.Validate:=False;
        //N.NumericType:=ntDecimal ;
        N.NumericType:=ntGeneral ;
(*        If GetParamsocSecur ('So_ChampObliAst',false) and (LBChamps.Cells[5,CurLig]='X') then
          begin  //mcd 03/01/2007 13528 GC     a remettre quand OK pour ThNUmEdit (agl 12757)
          N.obligatory:=true;
          end; *)
        if ExisteInfos then
           N.Value:=TobInfos.GetValue(LBChamps.Cells[3,CurLig]);
        end;
     if OkBlob then
        begin
        if Titre <> '' then P.Height:=StrToInt(Titre);
        L.Top:=(P.Height-L.height) div 2;
        BB := THRichEditOle.Create(Ecran); BB.Parent:=P ;
        BB.Anchors:=[akLeft,akRight];
        if OrigineFichier then
          BB.ReCreatePopup;
        BB.Name:=LBChamps.Cells[0,CurLig];
        BB.Left:=L.Left+L.Width+1;
        BB.Width:=P.Width-BB.Left-8 ;
        if ExisteInfos then
           begin
           sl := HTStringList.Create;
           sl.Add(TobInfos.GetValue(LBChamps.Cells[3,CurLig]));
           THRichEditOle(BB).SetLinesRTF (sl);
           freeandnil(sl);
           end
           //THRichEditOle(BB).LinesRTF.Text:=TobInfos.GetValue(LBChamps.Cells[3,CurLig])
        else
           BB.Lines.Text:='';
        BB.Height:=P.Height;
        L.FocusControl:=BB;
        end;
     end;
     LBChamps.RowHeights[CurLig]:=0;
     P.Caption:='';
end
else
begin
     Inc(NumSeparateur);

     if Assigned(TPanel(Ecran.FindComponent('PCHAMPS_'+IntToStr(NumSeparateur)))) then
        begin
        stMess:='Ce composant existe déjà : PCHAMPS_'+IntToStr(NumSeparateur);
        PgiInfo(stMess, 'Composant existant');
        exit;
        end;

     P.Name:='PCHAMPS_'+IntToStr(NumSeparateur);
     if OkSpace then
        begin
        P.Caption:=Titre;
        P.Font.Color:=clActiveCaption;
        end
     else
        P.Caption:='';
end;
DernierPanel:=P.Name;

if ( NewPos ) then
    begin
    NewPosition;
    for i:=0 to ListeActive.ControlCount-1 do if ListeActive.Controls[i] is TPanel then
        BEGIN
        if TPanel(ListeActive.Controls[i])=P then TPanel(ListeActive.Controls[i]).Color:=clTeal
                                      else
                                          TPanel(ListeActive.Controls[i]).Color:=CouleurPanel;
        END;
     end;
P.enabled := not VerrouModif;  // Confidentialité

if (not ExisteInfos) and (not OkMulti) and (not OkBlob)  and (Libelle <> '') then
    begin
    SetControlText(LBChamps.Cells[0,Curlig],Libelle); // Valeurs par défaut
    if (OrigineFichier) and ( OkCase or OkCombo ) then // alim Tob pour tests affichage Onglets
        TobInfos.PutValue(LBChamps.Cells[3,CurLig],Libelle);
    end;
end;

Procedure Tof_ParamCL.CreerChampTob (NomChamp : String);
begin
TobEcran.AddChampSup(NomChamp,True);
SavTobEcran.AddChampSup(NomChamp,True);
end;

procedure TOF_ParamMulti.ListeExit(Sender: TObject);
begin
if TlistBox(Sender).SelCount > NbElemListe then
   begin
   PGIInfo('Vous ne pouvez pas sélectionner plus de 26 éléments dans cette liste','');
   TlistBox(Sender).SetFocus;
   end;
end;

procedure TOF_ParamMulti.ListeClick(Sender: TObject);
begin
if TlistBox(Sender).SelCount > NbElemListe then
   begin
   PGIInfo('Vous ne pouvez pas sélectionner plus de 26 éléments dans cette liste','');
   if TlistBox(Sender).Selected[TlistBox(Sender).ItemIndex] = True then
      TlistBox(Sender).Selected[TlistBox(Sender).ItemIndex]:=False;
   TlistBox(Sender).SetFocus;
   end;
end;

procedure Tof_ParamCL.SelectionMulti(Sender: TObject);
Var Param,Valeur,Bouton : String;
    TL : TListBox;
    P : TPanel;
    iPosition : integer;
begin
Bouton:=TToolbarButton97(Sender).Name;
iPosition:=StrToInt(Copy(Bouton,pos('_',Bouton)+1,Length(Bouton)-pos('_',Bouton)));

if iPosition > 100 then exit;

TL:= TListBox(GetControl(LBChamps.Cells[0,iPosition]));
ChargeChampMulti (TListBox(TL),Valeur,True,True);
Param:='VALCHAMP='+Valeur+'|NOMCODE='+'RM'+copy(LBChamps.Cells[0,iPosition],3,1);
Param:=Param+'|PROSPECTEXISTE=1';
Param:=Param+'|TABLETTEMULTI='+TabletteComboMul ;
Valeur:=AGLLanceFiche('RT','RTPARAMMULTI','','',Param);
TL.Clear;
Repeat
  Param:=ReadTokenSt(Valeur) ;
  if Param<>'' then
     TL.Items.Add(Param);
until (Param = '') ;
P:=TPanel(TL.Parent);
TaillePanelMulti(P,TL,iPosition);
end;

procedure Tof_ParamCL.TaillePanelMulti(P : TPanel; TL : TListBox; iPosit : integer);
Var B : TToolbarButton97;
    L : TLabel;
begin
if TL.Items.Count > 0 then
   if TL.Items.Count > TL.Tag  then
      P.Height:=(TL.Tag )*HauteurLigneListe
   else
       if TL.Items.Count < NbLignesMini then
          P.Height:=(NbLignesMini)*HauteurLigneListe
       else
          P.Height:=(TL.Items.Count)*HauteurLigneListe
else
   begin
   P.Height:=HauteurLigneListe*2;
   TL.Items.Add (TraduireMemoire('<<Aucun>>'));
   end;
TL.Height:=P.Height;
L:= TLabel(GetControl('L'+IntToStr(iPosit)));
if Assigned (L) then
  begin
  B:= TToolbarButton97(GetControl('B'+LBChamps.Cells[0,iPosit]+'_'+IntToStr(iPosit)));
  L.Top:=(P.Height-L.height) div 2;
  if Assigned (B) then
    B.Top:=L.Top-((B.Height-L.Height) div 2);
  end;
end;

procedure ChargeListBox ( TabletteComboMul: string; MC:TListBox; Complet : Boolean; CleChoixCode : String; Valchamp : String; ProspectExiste : boolean);
var ii,indice,iElement,Pos1 : integer;
    Critere,Liste,Valeur,StString,stCombos : String ;
    Trouve : Boolean;
    StListe : HTStringList;
    StChamp : array[1..3] of string;
begin
stCombos:=TabletteComboMul+CleChoixCode[3];
ii:=TTToNum(stCombos) ;
if ii>0 then
  begin
  if (V_PGI.DECombos[ii].Valeurs = Nil) then RemplirListe (stCombos,'');
  StListe := V_PGI.DECombos[ii].Valeurs;
  if StListe = Nil then exit;
  iElement := 0;
  While iElement < StListe.count do
     begin
     Liste:=Valchamp;
     StString := StListe.Strings[iElement];
     for indice := 0 to 2 do
         begin
         Pos1 := Pos (#9,StString);
         if Pos1 > 1 then
            begin
            StChamp[indice+1] := Copy(StString,1,Pos1-1);Delete(StString,1,Pos1);
            end
         else
            begin
            StChamp[indice+1] := '';
            end;
         end;
     Valeur:=StChamp[1]+' '+StChamp[2];
     if Valeur <> ' ' then
         begin
         if Complet then
           MC.Items.Add(Valeur);
         if ProspectExiste then
            begin
            Trouve:=False;
            Repeat
              Critere:=Trim(ReadTokenSt(Liste)) ;
              Pos1:=Pos(' ',Critere);
              if (Complet) and (Pos1>1) then
                  // pos1>1 pour traiter le cas de valeurs par défaut
                  Critere := copy(Critere,1,(Pos1 -1)) ;
              if Critere = copy(Valeur,1,( Pos(' ',Valeur)-1)) then
                 begin
                 if not Complet then
                    MC.Items.Add(Valeur)
                 else
                    MC.Selected[iElement]:=True;
                 Trouve:=True;
                 end;
            until (Critere = '') or Trouve ;
            end;
         end;
     inc (iElement)
     end;
  end;

end;


procedure ChargeChampMulti (C:TListBox ; var Valeur: String; LibComplet : boolean; Tous : boolean);
var i,NbSel : integer;
begin
Valeur:='';
NbSel:=0;
For i:=0 to C.Items.Count - 1 do
    begin
    if C.Selected[i] or Tous then
       begin
       if LibComplet then
          Valeur:=Valeur+C.Items.Strings[i]+';'
       else
          Valeur:=Valeur+copy(C.Items.Strings[i],1,( Pos(' ',C.Items.Strings[i])-1))+';';
       Inc(NbSel);
       if NbSel = NbElemListe then break;
       end;
    end;
// 12-06-02
if Valeur = ';' then Valeur:='';
end;
// comparaison de 2 tobs virtuelles de structure identique
Function RTCompareTobs(Tob1,tob2 : tob) : boolean;
var i : integer;
begin
Result:=True;
for i := 1000 to (1000+(Tob1.ChampsSup.count-1)) do
    begin
    if (Tob1.FieldExists(Tob1.GetNomChamp(i))) and
        (Tob2.FieldExists(Tob2.GetNomChamp(i))) then
      if Tob1.GetValue(Tob1.GetNomChamp(i)) <> Tob2.GetValue(Tob2.GetNomChamp(i)) then
         begin
         Result:=False;
         exit;
         end;
    end;
end;

Procedure AddChpsProprietes(TobP : Tob);
Var Critere,Liste : string;
begin
    Liste:=ChampsProprietes;
    Repeat
    Critere:=ReadToKenSt(Liste);
    if Critere<>'' then
        TobP.AddChampSup(Critere,True);
    until Critere='';
end;

procedure Tof_ParamCL.Propriete();
var    P : TPanel;
    OngletActif,j,INum : Integer ;
    ListeActive : TScrollBox ;
        C : TControl;
        LC : TLabel;
        Param,StLib,Code : String;
begin
OngletActif:=TPageControl(getControl( 'PAGEONGLETS')).ActivePage.PageIndex+1;
ListeActive:= TScrollBox(GetControl('SCB'+IntToStr(OngletActif)));
P:= TPanel(Ecran.FindComponent('PCHAMPS_'+IntToStr(CurLig)));
if (P = Nil ) or (P.Height = HauteurLigneSeule) then exit; // hauteur 6 = ligne : aucun paramètre à renseigner

TobPropriete:=tob.create('Propriétés champ',Nil,-1);
AddChpsProprietes(TobPropriete);
StLib:='';
if ( Curlig < PosSeparateur ) then
  begin
  TobPropriete.PutValue( 'INTITULE', LBChamps.Cells[2,Curlig]);
  TobPropriete.PutValue( 'OBLIGATOIRE',LBChamps.Cells[5,Curlig]);
  TobPropriete.PutValue( 'CRITERESEL',LBChamps.Cells[6,Curlig]);
  TobPropriete.PutValue( 'TYPECHAMP',LBChamps.Cells[7,Curlig]);
  TobPropriete.PutValue( 'CHPASSOCIE',LBChamps.Cells[10,Curlig]);
  if LBChamps.Cells[1,Curlig] = 'COMBO' then StLib:=' ( '+TraduireMemoire('table n° ');
  if LBChamps.Cells[1,Curlig] = 'MULTI' then StLib:=' ( '+TraduireMemoire('multi-choix n° ');
  if LBChamps.Cells[1,Curlig] = 'EDIT' then StLib:=' ( '+TraduireMemoire('texte n° ');
  if LBChamps.Cells[1,Curlig] = 'DOUBLE' then StLib:=' ( '+TraduireMemoire('valeur n° ');
  if LBChamps.Cells[1,Curlig] = 'BOOLEAN' then StLib:=' ( '+TraduireMemoire('booléen n° ');
  if LBChamps.Cells[1,Curlig] = 'DATE' then StLib:=' ( '+TraduireMemoire('date n° ');
  if LBChamps.Cells[1,Curlig] = 'BLOB' then StLib:=' '+TraduireMemoire('bloc-notes');
  if LBChamps.Cells[1,Curlig] <> 'BLOB' then
  begin
    Code:= LBChamps.Cells[0,Curlig];
    INum:=Ord(Code[3]);
    if INum < 65 then
       INum:=StrToInt(Code[3])
    else
       INum:=Ord(Code[3])-55;
    StLib:=StLib+IntToStr(INum)+' )';
  end;
  TobPropriete.PutValue( 'NOMCHAMP',TraduireMemoire(StLib));
  TobPropriete.PutValue( 'CHOIXFILTRE',LBChamps.Cells[8,Curlig]);
  end
else
  TobPropriete.PutValue( 'INTITULE', P.Caption);

TobPropriete.PutValue( 'NBLIGNES',(P.Height div HauteurLigneListe));

//TobPropriete.PutValue( 'INTITULE', LBChamps.Cells[2,Curlig]);
TheTob:=TobPropriete;
if Curlig >= PosSeparateur then
   Param:='NON'
else
   Param:=LBChamps.Cells[1,Curlig];
Param:=Param+';'+CodeFichier ;
AGLLanceFiche('RT','RTPARAMCHAMP','','',Param);

if ( theTob<>Nil ) and ( Curlig < PosSeparateur ) then
    begin
    BougePropriete:=true;
    TobPropriete:=TheTob;
    TheTob:=Nil;
    LBChamps.Cells[2,Curlig]:=TobPropriete.GetValue( 'INTITULE' );
    LBChamps.Cells[5,Curlig]:=TobPropriete.GetValue( 'OBLIGATOIRE');
    LBChamps.Cells[6,Curlig]:=TobPropriete.GetValue( 'CRITERESEL');
    LBChamps.Cells[7,Curlig]:=TobPropriete.GetValue( 'TYPECHAMP');
    LBChamps.Cells[10,Curlig]:=TobPropriete.GetValue( 'CHPASSOCIE');
    if LBChamps.Cells[7,Curlig] <> '' then
       LBChamps.Cells[8,Curlig]:=TobPropriete.GetValue( 'CHOIXFILTRE')
    else
       LBChamps.Cells[8,Curlig]:='';
    LC:=Nil;
    For j:=0 to P.ControlCount-1 do
      begin
          C:=P.Controls[j];
          if (C is TLabel) and ( Copy(THLabel(C).Name,1,2) <> 'LT' ) then
          begin
               LC:=THLabel(C);
               THLabel(C).Caption:=LBChamps.Cells[2,Curlig];
          end
          else
             if C is TCheckBox then
                TCheckBox(C).Caption:=LBChamps.Cells[2,Curlig]
             else
                 if ( C is TListBox ) or ( C is THRichEditOle ) then
                 begin
                      P.Height:=TobPropriete.GetValue('NBLIGNES')*HauteurLigneListe;
                      if C is TListBox then
                         TListBox(C).Height:=P.Height
                      else
                         THRichEditOle(C).Height:=P.Height ;
                      LC.Top:=(P.Height-LC.height) div 2;
                 end
                 else
                     if C is THEdit then
                     begin
                     if LBChamps.Cells[7,Curlig] <> '' then
                         begin
                         THEdit(C).ElipsisButton:=True;
                         THEdit(C).Enabled:=true;
                         end
                     else
                         begin
                         THEdit(C).ElipsisButton:=false;
                         THEdit(C).Enabled:=false;
                         end;
                     end;
      end;
    end;

if ( Curlig >= PosSeparateur ) and ( TheTob <> nil ) then
   begin
    TobPropriete:=TheTob;
    TheTob:=Nil;

   if (P<>Nil) and (TobPropriete.GetValue( 'INTITULE' ) <> '') then
      P.Caption:= TobPropriete.GetValue( 'INTITULE' );
   end;
if assigned (TobPropriete) then
  FreeAndNil(TobPropriete);

ListeActive.Invalidate;
end;

procedure Tof_ParamCL.FicheSaisie();
begin
SetControlVisible('TPDIRECTION',FALSE) ;
SetControlVisible('AJOUONGLET',FALSE) ;
SetControlVisible('ONGLETGAUCHE',FALSE) ;
SetControlVisible('ONGLETDROITE',FALSE) ;
SetControlVisible('SUPPONGLET',FALSE) ;
//SetControlVisible('TITREONGLET',FALSE) ;
SetControlVisible('PROPRIETE',FALSE) ;
SetControlVisible('LBChamps',FALSE) ;
WidthEcran:=TFVierge(ecran).width;
HeightEcran:=TFVierge(ecran).height;
TFVierge(ecran).width:=GetParamsocSecur(ParamSocLargeur,300) ;
TFVierge(ecran).height:=GetParamsocSecur(ParamSocHauteur,300) ;
end;

procedure Tof_ParamCL.FicheParam();
begin
SetControlVisible('TPDIRECTION',TRUE) ;
SetControlVisible('AJOUONGLET',TRUE) ;
SetControlVisible('ONGLETGAUCHE',TRUE) ;
SetControlVisible('ONGLETDROITE',TRUE) ;
SetControlVisible('SUPPONGLET',TRUE) ;
//SetControlVisible('TITREONGLET',TRUE) ;
SetControlVisible('PROPRIETE',TRUE) ;
SetControlVisible('LBChamps',TRUE) ;
TFVierge(ecran).width:=WidthEcran ;
TFVierge(ecran).height:=HeightEcran ;
end;

procedure Tof_ParamCL.BoutonsParamInvisibles;
begin
SetControlVisible('TPDIRECTION',FALSE) ;
SetControlVisible('AJOUONGLET',FALSE) ;
SetControlVisible('ONGLETGAUCHE',FALSE) ;
SetControlVisible('ONGLETDROITE',FALSE) ;
SetControlVisible('SUPPONGLET',FALSE) ;
SetControlVisible('PROPRIETE',FALSE) ;
SetControlVisible('BTESTFICHE',FALSE) ;
SetControlVisible('BPARAMONGLET',FALSE) ;
end;

// ajout cd 221100
procedure Tof_ParamCL.SetArgumentsCL(StSQL : string);
var Critere,ChampMul,ValMul,StControl{,StCritere,Liste} : string ;
    x,r{,i} : integer ;
    Ctrl,C : TControl;
    Fiche : TFVierge;
    LB : TListBox;
    P  : TPanel;
begin
SetControlVisible('BSTOP',TRUE);
//SetControlVisible('BTIERS',TRUE);
//DS.Edit;
Fiche := TFVierge(ecran);
Repeat
    Critere:=Trim(ReadTokenPipe(StSQL,'|')) ;
    if Critere<>'' then
        begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           if copy(ValMul,1,1)='"' then ValMul:=copy(ValMul,2,length(ValMul));
           if copy(ValMul,length(ValMul),1)='"' then ValMul:=copy(ValMul,1,length(ValMul)-1);
           StControl := '';
           for r:=0 to LBCHAMPS.RowCount-1 do
               begin
               { mng 03/03/04 : on ne controle que la fin du champ pour gérer
                 correctement les infos complémentaires des contacts }
               if copy(LBChamps.Cells[3,R],8,length(LBChamps.Cells[3,R])) = copy(ChampMul,8,length(ChampMul)) then
                 begin
                 StControl := LBChamps.Cells[0,R];
                 break;
                 end;
              end;
           if StControl <> '' then
               begin
               Ctrl:=TControl(Fiche.FindComponent(StControl));
               if Ctrl=nil then continue;
               if (Pos('ML',StControl) = 0) then
                   SetControlText(StControl,ValMul)
               else
                   begin
                   LB := TListBox(GetControl(StControl));
                   LB.Clear;
                   ChargeListBox (TabletteComboMul, LB,( not OrigineFichier) ,StControl,ValMul,true);
                   end;
               if (Ctrl is TCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is THEdit) Or (Ctrl is THNumEdit) Or (Ctrl is TListBox)then
                  TEdit(Ctrl).Font.Color:=clRed
               else if Ctrl is TSpinEdit then
                  TSpinEdit(Ctrl).Font.Color:=clRed;
    // pour mettre le libelle du controle en rouge aussi
               P := TPanel(Ctrl.Parent);
               for r:=0 to P.ControlCount-1 do
                    begin
                    C:=P.Controls[r];
                    if C is TLabel then
                    begin
                    TLabel(C).Font.Color:=clRed;
                    break;
                    end;
                   end;
               end;
           end;
        end;
until  Critere='';
end;

procedure Tof_ParamCL.ChampExit(Sender: TObject) ;
begin
if (OrigineFichier) and (TypeAction = '') and ( Curlig < PosSeparateur ) then
    begin
    TobInfos.PutValue(LBChamps.Cells[3,CurLig],GetControlText(LBChamps.Cells[0,CurLig]));
    TestAffichageOnglets;
    end;
end;
// ajout cd 221100

procedure AGLAjouterChamps(parms:array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  if (ToTof is Tof_ParamCL) then Tof_ParamCL(ToTof).AjouterChamps else exit;
end;

procedure AGLEnleverChamps(parms:array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  if (ToTof is Tof_ParamCL) then Tof_ParamCL(ToTof).EnleverChamps else exit;
end;

procedure AGLMonterChamps(parms:array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  if (ToTof is Tof_ParamCL) then Tof_ParamCL(ToTof).MonterChamps else exit;
end;

procedure AGLDescendreChamps(parms:array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  if (ToTof is Tof_ParamCL) then Tof_ParamCL(ToTof).DescendreChamps else exit;
end;

procedure AGLAjouterOnglet(parms:array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  if (ToTof is Tof_ParamCL) then Tof_ParamCL(ToTof).AjouterOnglet else exit;
end;

procedure AGLOngletADroite(parms:array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  if (ToTof is Tof_ParamCL) then Tof_ParamCL(ToTof).OngletADroite else exit;
end;

procedure AGLOngletAGauche(parms:array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  if (ToTof is Tof_ParamCL) then Tof_ParamCL(ToTof).OngletAGauche else exit;
end;
procedure AGLEnleverOnglet(parms:array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  if (ToTof is Tof_ParamCL) then Tof_ParamCL(ToTof).EnleverOnglet else exit;
end;

procedure AGLFicheSaisie(parms:array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  if (ToTof is Tof_ParamCL) then Tof_ParamCL(ToTof).FicheSaisie else exit;
end;

procedure AGLFicheParam(parms:array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  if (ToTof is Tof_ParamCL) then Tof_ParamCL(ToTof).FicheParam else exit;
end;

procedure AGLPropriete(parms:array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  if (ToTof is Tof_ParamCL) then Tof_ParamCL(ToTof).Propriete else exit;
end;

procedure AGLRTOuvreFiche(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     code : string;
     TobChampsProFille : tob;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (copy(F.Name,1,7) = 'GCTIERS') then code:=CodeFichierProspect
  else if F.Name = getParamSocSecur('SO_AFFICHETIERS','GCTIERS') then code:=CodeFichierProspect //mcd 28/06/07 pour fiche tiers personalisée
  else if (copy(F.Name,1,13) = 'GCFOURNISSEUR') then code:=CodeFichierFournisseur
  else if (copy(F.Name,1,9) = 'YYCONTACT') then code:=CodeFichierContact
  else if (copy(F.Name,1,16) = 'RTTYPESCHAINAGES') or (copy(F.Name,1,13) = 'RTTYPEACTIONS') then code:=CodeFichierAction;

  VH_RT.TobChampsPro.Load;
  VH_RT.TobTypesAction.Load;

  TobChampsProFille:=VH_RT.TobChampsPro.FindFirst(['CO_CODE'], [code], TRUE);
  if (TobChampsProFille = Nil ) or (TobChampsProFille.detail.count = 0 ) then
      begin
      PGIInfo('Vous ne pouvez pas accéder à cette fiche.#13#10Le paramétrage de cette saisie n''a pas été effectué','');
      exit;
      end;
  AGLLanceFiche(Parms[1],Parms[2],Parms[3],Parms[4],Parms[5]);
end;


Initialization
registerclasses([Tof_ParamCL]);
registerclasses([Tof_ParamChamp]);
registerclasses([Tof_ParamMulti]);
RegisterAglProc('AjouterChamps',true,0,AGLAjouterChamps);
RegisterAglProc('EnleverChamps',true,0,AGLEnleverChamps);
RegisterAglProc('AjouterOnglet',true,0,AGLAjouterOnglet);
RegisterAglProc('EnleverOnglet',true,0,AGLEnleverOnglet);
RegisterAglProc('MonterChamps',true,0,AGLMonterChamps);
RegisterAglProc('DescendreChamps',true,0,AGLDescendreChamps);
RegisterAglProc('FicheSaisie',true,0,AGLFicheSaisie);
RegisterAglProc('FicheParam',true,0,AGLFicheParam);
RegisterAglProc('Propriete',true,0,AGLPropriete);
RegisterAglProc('RTOuvreFiche',true,5,AGLRTOuvreFiche);
RegisterAglProc('OngletAGauche',true,0,AGLOngletAGauche);
RegisterAglProc('OngletADroite',true,0,AGLOngletADroite);
end.


