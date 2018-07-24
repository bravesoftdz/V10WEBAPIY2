{***********UNITE*************************************************
Auteur  ...... : ?????
Créé le ...... : 17/05/2000
Modifié le ... :   /  /
Description .. : TOM de la fiche Contacts multi types
Suite ........ :
Suite ........ : arguments : TYPE = TIE pour les tiers
Suite ........ :                 TYPE2 = SAL pour les salariés
Suite ........ :                 TYPE2 = CLI pour les clients
Suite ........ :                 TYPE2 = FOU pour les fournisseurs
Suite ........ :            TYPE = ET pour les entreprises
*****************************************************************}
unit UTomContact;
 
interface

Uses
 ParamSoc,AglIsoflex,
{$IFDEF EAGLCLIENT}
  spin,Efiche,
{$ELSE}
  DB,Hqry, Fiche,
  {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  utoc,
{$IFDEF GCGC}
  EntGC,UtilGC,
{$IFDEF CTI}
  CtiGrc,UtilCti,CtiPcb,
{$ENDIF}
{$IFNDEF CCS3}
  UtilConfid,
{$ENDIF}
{$IFDEF GRC}
  EntRT, UtilRT,
{$ENDIF}
{$ENDIF}
  HTB97,

  UTOM, Forms, controls,Classes,StdCtrls,Graphics,Menus,UTableFiltre,windows,AglInit,
  wCommuns, UTob
{$IF Defined(CRM) or Defined(GRCLIGHT)}
  ,UtofRTSaisieInfosEsker,EskerInterface
{$IFEND}

 ;

type
    TOM_Contact = class(TOM)
      private
       LibTitre : String;
       type_tiers :String;
       GuidPer :String;   //mcd 20/06/2006 mis Guid au lie de codeper
       Particulier,InfoParticulier :Boolean;
       IsAnnuaire: boolean; //mcd 14/09/07 bureau 11677
       NomContact :String;
       Type_contact: string;
       FromSaisie : boolean ;
       FirstTime : boolean ;
       SaisieNom   : String ;
       ModifLot : boolean;
       StSQL : string;
       StTiers : string ;  // renseigné qd appel des contacts depuis GCTIERS (rend les boutons GRC actifs)
       CodeTiers : string; // pr MAJ c_tiers en création de contact
       Origin : string;
       IsModify : Boolean;
       CanCreate : boolean;
       {$ifdef GIGI}
       CodeAuxiliaire : string; // mcd 07/09/2006 12599
       TobValTiers:tob;
       {$endif}
       LaToc : Toc;
       {$IFDEF CTI}
       AppelCtiOk,SerieCTI,ContexteCti,CtiCCA : Boolean;
       CtiHeureDebCom,CtiHeureFinCom : TDateTime;
       WhereUpdate : String;
       BPHONE,BTELBUR,BTELDOM,BTELPORT,BAttente,BStop : TToolBarButton97;
       procedure SetEtatBoutonCti;
       procedure FindContactBTP(TypeContact, Auxiliaire: string;Numero: integer);
       procedure ChangeInfoSUp(Sender: Tobject);
       procedure SetEvents(Etat: boolean);
       procedure DeleteContactBTP(TypeContact, Auxiliaire: string;Numero: integer);
       {$ENDIF}
      private
       TF : TTableFiltre;
       AlerteInitTob : tob; // tob valeurs initiales pour gestion des alertes (tables complémentaires)
      TOBCONTACTBTP : TOB;
      {$IFDEF BTP}
       procedure FLISTE_OnDblClick( Sender: TObject );
      {$ENDIF}
       procedure SetArguments(StSQL : string);
       procedure AddRemoveItemFromPopup(stPopup,stItem : string; bVisible : boolean);     //PCS
       procedure SetLastError (Num : integer; ou,valeur : string );
       Procedure AffContactIndisponible ;
       procedure FermeContact(Sender: TObject);
       function ExisteContact (NomTable, NomChamp1, NomChamp2, TypeRech: string ; NumErr : integer) : boolean;
       function ExisteContact_InfosCompl (Prefixe, NoDesc : string ; NumErr : integer) : boolean;
       {$IFDEF CTI}
       procedure RTCTIDecrocheAppelContact ;
       procedure RTCTIRaccrocheAppelContact ;
       procedure RTCTINumeroterContact ;
       procedure RTCTIAppelAttenteContact ;
       procedure RTCTIRepriseAppelContact ;
       {$ENDIF}
       {$IFDEF GRC}
       procedure RTAppelParamCLContact;
       {$ENDIF}
      function TestAlerteContact (CodeA : String) : boolean;
      procedure ListAlerte_OnClick_C(Sender: TObject);
      procedure Alerte_OnClick_C(Sender: TObject);
      {$IFDEF GCGC}
      procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
      {$ENDIF}
      procedure GereIsoflex;
      procedure BContactClick     (Sender : TObject);
      {$IFDEF GIGI}
      procedure AffectValTiers (CodeTiers: string) ;
      procedure AFAffectVal(Champ: string;QQ:Tquery) ;
      procedure ChargeTobValTiers ;
      {$ENDIF}
      procedure SetButtonEna;

      public
{$IFDEF CTI}
         procedure RTCTIAfficheMessageCTI;
{$ENDIF}
       procedure OutLookInsertContact;
       procedure OnChangeField (F : TField)  ; override ;
       procedure OnUpdateRecord  ; override ;
       procedure OnLoadRecord  ; override ;
       procedure OnLoadAlerte  ; override ;
       procedure OnNewRecord  ; override ;
       procedure OnClose   ; override ;
       procedure OnArgument (Arguments : String ); override ;
       procedure OnCancelRecord             ; override ;
       function PrincipalUnique : Boolean;
       function LastNum : LongInt;
       Procedure OnDeleteRecord ; override ;
{$IFDEF CTI}
      procedure BTELBUR_OnClick(Sender: TObject);
      procedure BTELDOM_OnClick(Sender: TObject);
      procedure BTELPORT_OnClick(Sender: TObject);
      procedure BAttenteClick(Sender: TObject);
      procedure BPhoneClick(Sender: TObject);
      procedure BStopClick(Sender: TObject);
{$ENDIF CTI}
{$IF Defined(CRM) or Defined(GRCLIGHT)}
         procedure BSMS_OnClick(Sender: TObject);
{$IFEND}
     END ;

const
     TexteMessage : Array[1..18] of String = (
          {1} 'Vous avez déjà indiqué un contact principal.'
          {2} ,'Le nom du contact doit être renseigné.'
          {3} ,'Suppression impossible : il existe des actions pour ce contact.'
          {4} ,'Suppression impossible : il existe des propositions pour ce contact.'
          {5} ,'Suppression impossible : il existe des projets pour ce contact.'
          {6} ,'Suppression impossible : il existe des affaires pour ce contact.'
          {7} ,'Suppression impossible : ce contact est utilisé en historique de propositions.'
          {8} ,'Suppression impossible : ce contact est utilisé dans les informations complémentaires.'
          {9} ,'Suppression impossible : ce contact est utilisé dans les informations complémentaires actions.'
         {10} ,'Suppression impossible : il existe des contrats pour ce contact.'
         {11} ,'Suppression impossible : il existe des dossiers clients pour ce contact.'
         {12} ,'Suppression impossible : il existe des lignes de dossiers clients pour ce contact.'
         {13} ,'Suppression impossible : Ce contact est utilisé dans les adresses.'
         {14} ,'Suppression impossible : Ce contact est utilisé dans les compléments d''en-tête de pièces.'
         {15} ,'Ce contact est le contact principal.'
         {16} ,'La saisie du champ suivant est obligatoire : '
         {17} ,'Suppression impossible : ce contact est utilisé dans les informations complémentaires fournisseurs.'
         {18} ,'Suppression impossible : ce contact est affecté dans une intervention ou intervenant d''une affaire %s.'
           );

implementation

Uses
  SysUtils
  , HCtrls
  , HEnt1
  , HDB
  {$IFDEF EAGLCLIENT}
  , MaineAGL
  {$ELSE}
  , dbctrls
  , Fe_Main
  {$ENDIF}
  , M3FP
  , HMsgBox
  , UtilPGI
  , SaisieList
  , TiersUtil
  , UtofContact
  {$IFDEF AFFAIRE}
  , Dicobtp
  , CalcOleGenericAff
  {$ENDIF AFFAIRE}
  {$IFDEF DP}
  , entDP
  {$ENDIF}
  , CtiAlerte
  , UtilAlertes
  , YAlertesConst
  , Entpgi
  , BtpUtil
  , HrichOle
  , CbpMCD
  , CbpEnumerator
  , HPanel
  , CommonTools
  , HSysMenu
  ;
function GetAllSelectFields(TableN : String) : String ;
var
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
begin
    //récupération des champs de la table annulien pour Ordre "select"
    Result := '*';
    table := Mcd.getTable(TableN);
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
    begin
      if trim((FieldList.Current as IFieldCOM).name) <> '' then
      begin
        if Result <> '' then
           Result := Result + ',' ;
        Result := Result+(FieldList.Current as IFieldCOM).name;
      end;
    end;
end;

procedure TOM_CONTACT.SetLastError (Num : integer; ou,valeur : string );
begin
if ou<>'' then SetFocusControl(ou);
LastError:=Num;
{$IFDEF AFFAIRE}
if (Valeur <> '') and (Num = 18) then
  LastErrorMsg := format (TraduitGA (TexteMessage [LastError] ), [CodeAffaireAffiche(valeur)])
else
if (Valeur <> '') then
  LastErrorMsg := format (TraduitGA (TexteMessage [LastError] ), [valeur])
else
  LastErrorMsg:=TraduitGA(TexteMessage[LastError]);
{$ELSE}
if Valeur <> '' then
  LastErrorMsg := format (TraduireMemoire (TexteMessage [LastError] ), [valeur])
else
  LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$ENDIF AFFAIRE}
end ;

procedure TOM_CONTACT.SetArguments(StSQL : string);
var Critere,ChampMul,ValMul : string ;
    x,y : integer ;
    Ctrl : TControl;
    Fiche : TFSaisieList;
begin
SetControlVisible('BSTOP',TRUE);
//DS.Edit;
if TFSaisieList(Ecran).LeFiltre.State=dsBrowse then TFSaisieList(Ecran).LeFiltre.Edit;
Fiche := TFSaisieList(ecran);
Repeat
    //Critere:=uppercase(Trim(ReadTokenPipe(StSQL,'|'))) ;
    Critere:=Trim(ReadTokenPipe(StSQL,'|')) ;
    if Critere<>'' then
        begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           y := pos(',',ValMul);
           if y<>0 then ValMul:=copy(ValMul,1,length(ValMul)-1);
           if copy(ValMul,1,1)='"' then ValMul:=copy(ValMul,2,length(ValMul));
           if copy(ValMul,length(ValMul),1)='"' then ValMul:=copy(ValMul,1,length(ValMul)-1);
           SetField(ChampMul,ValMul);
           Ctrl:=TControl(Fiche.FindComponent(ChampMul));
           if Ctrl=nil then exit;
{$IFDEF EAGLCLIENT}
           if (Ctrl is TCustomCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is TCustomEdit) then TEdit(Ctrl).Font.Color:=clRed
           else if Ctrl is TSpinEdit then TSpinEdit(Ctrl).Font.Color:=clRed
              //mcd 25/01/01 ajout test TH pour zones YTC de tierscompl
           else if (Ctrl is TCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is THEdit)Or (Ctrl is THNumEdit)then
              begin
              TSpinEdit(Ctrl).Font.Color:=clRed;
              SetControlText(ChampMul,ValMul);
              end;
{$ELSE}
           if (Ctrl is TDBCheckBox) or (Ctrl is THDBValComboBox) Or (Ctrl is THDBEdit) then TEdit(Ctrl).Font.Color:=clRed
           else if Ctrl is THDBSpinEdit then THDBSpinEdit(Ctrl).Font.Color:=clRed
              //mcd 25/01/01 ajout test TH pour zones YTC de tierscompl
           else if (Ctrl is TCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is THEdit)Or (Ctrl is THNumEdit)then
              begin
              THDBSpinEdit(Ctrl).Font.Color:=clRed;
              SetControlText(ChampMul,ValMul);
              end;
{$ENDIF}
           end;
        end;
until  Critere='';
end;

procedure TOM_CONTACT.OnArgument (Arguments : String );
var
  Critere : string ;
  ChampMul,ValMul,Range,CodeAuxi,S : string;
  x,wi : integer ;
  bComplement,bAnnuaire : boolean;
  pop : TPopupMenu;
  CanInsertChoixCod : boolean;
begin
  inherited ;
  TOBCONTACTBTP := TOB.Create ('BCONTACT',nil,-1);

  AppliqueFontDefaut (THRichEditOle(GetControl('C_BLOCNOTE')));
  S := Arguments;
  {$IFDEF CTI}
  WhereUpdate:='';
  ContexteCti:=False;
  CtiCCA:=False;
  {$ENDIF}

  {$ifdef GIGI}
  chargeTobValTiers;	//mcd 11/07/06
  {$ENDIF}
  IsAnnuaire :=False; //mcd 14/09/07 11677
  x := pos('MODIFLOT',Arguments);
  ModifLot := x<>0;
  if ModifLot then
    begin
  //  TFSaisieList(Ecran).MonoFiche:=true;
    StSQL := copy(Arguments,x+9,length(Arguments));
    end;

  LibTitre := TraduireMemoire('Contacts : ');
  {récup des arguments }
  Type_contact := 'T';
  type_tiers := '';
  GUidPer:=''; //mcd 12/2005
  Particulier := False;
  InfoParticulier := False;
  FromSaisie:=False;
  StTiers:='';
  bComplement:=false;                                                           
  bAnnuaire:=false;
  Origin := GetArgumentString(Arguments, 'ORIGIN');
  Repeat
     Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
     if Critere<>'' then
        begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           end else
           begin
           ChampMul:=Critere; ValMul:='';
           end;

         if (ChampMul='TYPE') then Type_contact := ValMul
         else if (ChampMul='GUIDPER') then
            BEGIN //mcd 12/2005
            Guidper:=ValMul;
            if Assigned(GetControl('C_GUIDPER')) then SetField ('C_GUIDPER',GuidPer);
            END
         else if (ChampMul='ANNUAIRE') then begin if ValMul = 'TRUE' then Isannuaire:=True;end //mcd 14/09/07
         else if (ChampMul='TYPE2') then  type_tiers := ValMul
         else if (ChampMul='TITRE') then LibTitre   := ValMul
         else if (ChampMul='PART') then
              begin
                InfoParticulier := True;
                if (ValMul='X')  then Particulier := True;
              end
         else if (ChampMul='TIERS') then  StTiers := ValMul
         else if (ChampMul='ANNUAIRE') then
              begin
                SetControlVisible ('BZOOM',True);
                if (ValMul = 'PRO') or (ValMul='CLI') then
                  bComplement:=true;
                if (ValMul = 'FOU') then
                  bAnnuaire:=true;
              end
         else if (ChampMul='COMPLEMENT') then
                bComplement:=true
         else if (ChampMul='ALLCONTACT') then SetControlVisible ('BALLCONTACT',True)
         else if (ChampMul='FROMSAISIE') then begin FromSaisie:=True; SaisieNom:=ValMul end
         ;
  {$IFDEF CTI}
         if (critere='DECROCHE') and (CtiHeureDeb<>0) then
             begin
             AppelCtiOk:=True;
             CtiHeureDebCom:=CtiHeureDeb;
             end;
         if (critere='COUPLAGE') or (critere='NUMEROTER') then
             RTFormCti:=Ecran;
         if (critere='CTI') then ContexteCti:=True;
         if critere='CCA' then
            CtiCCA:=true;
  {$ENDIF}
        end;
  until  Critere='';

  if Ecran <> nil then
  begin
    if InfoParticulier = False then
    begin
      CodeAuxi := TiersAuxiliaire(GetArgumentValue(S, 'TIERS'));
      if CodeAuxi = '' then
      begin
        Range := TFsaisieList(ecran).FRange;
        CodeAuxi := ReadTokenSt(Range);      // Type de contact
        CodeAuxi := ReadTokenSt(Range);      // code Auxiliaire
      end;
      Particulier := ExisteSQL('SELECT T_PARTICULIER FROM TIERS WHERE T_AUXILIAIRE = "' + CodeAuxi + '" AND T_PARTICULIER = "X"');
    end;
    if Type_tiers = 'CLI' then
      CanCreate := Tools.CanInsertedInTable('CONTACT'{$IFDEF APPSRV}, '', '' {$ENDIF APPSRV})
    else
      CanCreate := True;
    if not CanCreate then
    begin
      FicheReadOnly(TFSaisieList(ecran));
      TFSaisieList(ecran).TypeAction := taConsult;
    end;
    SetButtonEna;
  end;
  if (Type_contact = 'T') and (Type_tiers = '') then Type_tiers := GetField ('C_NATUREAUXI');

  if FromSaisie then begin SetControlChecked('FROMSAISIE',true); SetControlProperty('BVALIDER','ModalResult',mrOK); end;;
  FirstTime:=true;

  { mng 18-02-04 : on vient de l'annuaire clients/prospects, on peut consulter les données complémentaires }
  if ((ctxGRC in V_PGI.PGIContexte) and
      (((StTiers <> '') and ((type_tiers = 'CLI') or (type_tiers = 'PRO'))) or (bComplement=true)) )
          or ( bAnnuaire=true ) or (type_tiers = 'FOU')
  {$IFDEF CTI}
          or (ContexteCti)
  {$ENDIF}
     then
     begin
     if (StTiers <> '') or (type_tiers = 'FOU') then
        SetControlVisible ('BACTION',True);
     SetControlVisible ('BCOMPLGRC',True);
     SetControlText ('TIERS',StTiers);

     if (GetParamSocSecur('SO_RTPROJGESTION',False) = False) or (type_tiers = 'FOU') or (bAnnuaire )
  {$IFDEF GRCLIGHT}
       or ( not GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False) )
  {$ENDIF GRCLIGHT}
     then
        AddRemoveItemFromPopup ('POPM2','MNPROJETS',False);

  {$IFNDEF CCS3}
     if (type_tiers = 'FOU') or (bAnnuaire ) then
  {$ENDIF}
        begin
        AddRemoveItemFromPopup ('POPM2','MNPROPOSITIONS',False);
  {$IFNDEF CCS3}
  {$IFDEF GCGC}
        if VH_GC.GRFSeria = false then
  {$ENDIF}
  {$ENDIF}
  {CRM_GRF10049_CD}
           begin
           AddRemoveItemFromPopup ('POPM2','MNACTIONS',False);
  {CRM_GRF10049_CD_DEB}
           SetControlVisible ('BACTION',False);
           end;
  {CRM_GRF10049_CD_FIN}
        end;
  {$IFDEF GRC}
     if (GetParamSocSecur('SO_RTGESTINFOS006',False) = False) or (type_tiers = 'FOU') or (bAnnuaire ) then
       AddRemoveItemFromPopup ('POPM2','MNFICHEINFOS',False);
  {$ENDIF}
     end;
  {$IFDEF CTI}
    if Assigned(GetControl('BPHONE')) then
    begin
      BPhone := TToolbarButton97(GetControl('BPHONE'));
      TToolbarButton97(GetControl('BPHONE')).OnClick := BPhoneClick;
    end;

    if Assigned(GetControl('BSTOP')) then
    begin
      BStop := TToolbarButton97(GetControl('BSTOP'));
      TToolbarButton97(GetControl('BSTOP')).OnClick := BStopClick;
    end;

    if Assigned(GetControl('BATTENTE')) then
    begin
      BAttente := TToolbarButton97(GetControl('BATTENTE'));
      TToolbarButton97(GetControl('BATTENTE')).OnClick := BAttenteClick;
    end;

    if Assigned(GetControl('BTELPORT')) then
    BEGIN
      BTELPORT := TToolbarButton97(GetControl('BTELPORT'));
      TToolbarButton97(GetControl('BTELPORT')).OnClick := BTELPORT_OnClick;
    END;

    if Assigned(GetControl('BTELBUR')) then
    begin
      BTELBUR := TToolbarButton97(GetControl('BTELBUR'));
      TToolbarButton97(GetControl('BTELBUR')).OnClick := BTELBUR_OnClick;
    end;

    if Assigned(GetControl('BTELDOM')) then
    begin
      BTELDOM := TToolbarButton97(GetControl('BTELDOM'));
      TToolbarButton97(GetControl('BTELDOM')).OnClick := BTELDOM_OnClick;
    end;

  {$ENDIF CTI}

  if Not (CtxGrc in V_PGI.PGIContexte) then //mcd 03/02/2005 10287 on cache option GRC
    begin
    AddRemoveItemFromPopup ('POPM2','MNPROPOSITIONS',False);
    AddRemoveItemFromPopup ('POPM2','MNPROJETS',False);
    AddRemoveItemFromPopup ('POPM2','MNACTIONS',False);
    AddRemoveItemFromPopup ('POPM2','MNFICHEINFOS',False);
    end;

{ mng 15-09-04 : en S3, pas de GRF }
{$IFDEF CCS3}
   if GetField ('C_NATUREAUXI') = 'FOU' then
     begin
     SetControlVisible ('BACTION',false);
     SetControlVisible ('BCOMPLGRC',false);
     end;
{$ENDIF}

  if (Type_contact = 'T') then
  begin
    if LibTitre <> 'Contacts : ' then
    begin
      if (type_tiers = 'SAL') then LibTitre := TraduireMemoire('Contacts du salarié : ')     + LibTitre ;
      if (type_tiers = 'CLI') then LibTitre := TraduireMemoire('Contacts du client : ')      + LibTitre;
      if (type_tiers = 'FOU') then LibTitre := TraduireMemoire('Contacts du fournisseur : ') + LibTitre;
    end
  end
  else if (Type_contact = 'ET') then
  begin
    LibTitre := TraduireMemoire('Contacts de la Société : ') + LibTitre ;
    type_tiers := 'ET';
  end
  else if (Type_contact = 'GCL') then
  begin
    LibTitre := TraduireMemoire('Contacts du commercial : ') + LibTitre ;
    type_tiers := 'GCL';
  end;

  {$IFDEF GCGC}
  //¨Paramétrage des libellés des zones libres
  if Ecran <> nil then
  begin
    if (TFSaisieList(Ecran).name = 'YYCONTACT') then
    begin
      setcontrolVisible('Page2',true) ;  // avt P_INFO
      x := 0;
      if (GCMAJChampLibre(TForm (Ecran), False, 'COMBO', 'C_LIBRECONTACT', 10, '') = 0) then
        SetControlVisible('GB_TABLELIBRE', False) else inc(x);
      if (GCMAJChampLibre(TForm (Ecran), False, 'EDIT', 'C_VALLIBRE', 3, '_') = 0) then
        SetControlVisible('GB_MONTANTS', False) else inc(x);
      if (GCMAJChampLibre(TForm (Ecran), False, 'EDIT', 'C_DATELIBRE', 3, '_') = 0) then
        SetControlVisible('GB_DATES', False) else inc(x);
      if (GCMAJChampLibre(TForm (Ecran), False, 'BOOL', 'C_BOOLLIBRE', 3, '_') = 0) then
        SetControlVisible('GB_DECISIONS', False) else inc(x);
      if (GCMAJChampLibre(TForm (Ecran), False, 'EDIT', 'C_TEXTELIBRE', 3, '_') = 0) then
        SetControlVisible('GB_TEXTES', False) else inc(x);
      if (x = 0) then
        SetControlVisible('Page2', False);
      if (type_tiers = 'CLI') or (type_tiers = 'PRO') then
        CanInsertChoixCod := Tools.CanInsertedInTable('CHOIXEXT'{$IFDEF APPSRV}, '', '' {$ENDIF APPSRV})
      else
        CanInsertChoixCod := True;
      for x := 1 to 9 do
        SetControlProperty('C_LIBRECONTACT' + IntToStr(x), 'DataTypeParametrable', CanInsertChoixCod);
      SetControlProperty('C_LIBRECONTACTA', 'DataTypeParametrable', CanInsertChoixCod);
    end;
    TFSaisieList(Ecran).OnKeyDown:=FormKeyDown ;
  end;
  {$ENDIF}
  {$IFDEF BTP}
  if Assigned(GetControl('FLISTE')) then
      THGrid(GetControl('FLISTE' )).OnDblClick := FLISTE_OnDblClick;
  {$ENDIF}


  {$IFDEF GRC} // provisoirement uniquement si GRC
  {$IFDEF CTI}
  SetControlVisible('BZOOM',(ctxGRC in V_PGI.PGIContexte));
  {$ENDIF}
  if (ctxGRC in V_PGI.PGIContexte)
  {$IFDEF GCGC}
      or (VH_GC.GRFSeria = true)
  {$ENDIF}
  then SetControlVisible('BOUTLOOK',True );
  //SetControlVisible('C_LIENTIERS',(ctxGRC in V_PGI.PGIContexte));
  //SetControlVisible('TC_LIENTIERS',(ctxGRC in V_PGI.PGIContexte));
  //SetControlVisible('TC_LIENTIERS_',(ctxGRC in V_PGI.PGIContexte));
  {$ENDIF}

  SetActiveTabSheet('Page1');  // avt PGeneral
  if Ecran <> nil then 
    TCheckBox(GetControl('C_FERME')).OnClick := FermeContact;
  CodeTiers := stTiers;
  {$ifdef GIGI}
  CodeAuxiliaire:=TiersAuxiliaire(CodeTiers,false); //mcd 07/09/2006 12599
  {$endif}
  GereIsoflex;

// gm 15/10/04 pour eviter en S3 , sans  options d'avoir un bouton sans popup
  pop := TPopupMenu(GetControl('POPM2'));
  if pop <> nil then
  begin
    SetControlVisible('BCOMPLGRC',false);
    for wi := 0 to pop.items.count - 1 do
      if pop.items[wi].visible = true then
      begin
        SetControlVisible('BCOMPLGRC',true);
        break;
      end;
  end;

// $$$ JP 13/06/05 - CTI du bureau (pas de define CTI surtout!!)
//mcd {$IFDEF DP}
{$IFDEF BUREAU}
  SetControlVisible ('BTELDOM',  VH_DP.ctiAlerte <> nil);
  SetControlVisible ('BTELPORT', VH_DP.ctiAlerte <> nil);
  SetControlVisible ('BTELBUR',  VH_DP.ctiAlerte <> nil);
{$ENDIF}
  //mcd 14/09/07 11677 bureau: si appel depuis fiche annauire, il ne faut pas voir le bouton pour accéder au contact tiers
//if (ctxscot in V_PGI.PgiContexte) and (Guidper <>'' ) and (Not IsAnnuaire) then      //mcd 28/03/2007 13847 ajout test guiper pour avoir bouton que si appel depuis tiers
  if (ctxscot in V_PGI.PgiContexte) and (Guidper <>'' ) then      //md 26/09/2007 FQ 13847 revue => le bouton est de nouveau accessible
  begin //mcd 22/02/2006 12599 GAGI
  if Assigned (GetControl('BCONTACT'))  then
    begin
    SetControlVisible ('BCONTACT',true);
    TToolbarButton97(GetControl('BContact')).OnClick := BContactClick;
    end;
  end;

  if Assigned(GetControl('MnAlerte')) then
    if AlerteActive('C') then
      TMenuItem(GetControl('MnAlerte')).OnClick := Alerte_OnClick_C
    else
      TMenuItem(GetControl('MnAlerte')).visible:=false;

  if Assigned(GetControl('MnListAlerte')) then
    if AlerteActive('C') then
      TMenuItem(GetControl('MnListAlerte')).OnClick := ListAlerte_OnClick_C
    else
      TMenuItem(GetControl('MnListAlerte')).visible:=false;

  if Assigned(GetControl('MnGestAlerte')) and Assigned(GetControl('MnAlerte'))
     and Assigned(GetControl('MnListAlerte')) then
         TMenuItem(GetControl('MnGestAlerte')).visible := (TMenuItem(GetControl('MnAlerte')).visible)
          and (TMenuItem(GetControl('MnListAlerte')).visible);

{$IF Defined(CRM) or Defined(GRCLIGHT)}
  if GetParamSocSecur ('SO_RTGESTIONFLYDOC',False) = True then SetControlVisible ('BSMS',True);
  if Assigned(GetControl('BSMS')) then
    TToolbarButton97(GetControl('BSMS')).OnClick := BSMS_OnClick;
{$IFEND}
  if not GetParamSocSecur('SO_RTCTIGESTION', False) then
  begin
    BPhone.Visible := GetParamSocSecur('SO_RTCTIGESTION', False);
    BAttente.Visible:= GetParamSocSecur('SO_RTCTIGESTION', False);
    BStop.Visible:= GetParamSocSecur('SO_RTCTIGESTION', False);
  end else
  begin
  	SetEtatBoutonCti;
    If TFFiche(Ecran).fTypeAction = TaModif then
    Begin
       BTELDOM.Visible := true;
       BTELBUR.Visible := true;
       BTELPORT.Visible := true;
    end;
  end;
end;


// Vérifications avant la mise à jour d'un contact
procedure TOM_Contact.OnUpdateRecord;
var Q : TQuery;
    numero : Integer;
{$IFDEF GCGC}
{$IFNDEF CCS3}
    NomChamp : string;
{$ENDIF}
{$ENDIF}
begin
inherited;
    // La vérification de C_NOM = '' se fait toute seule
    if not PrincipalUnique then begin SetLastError(1,'C_PRINCIPAL','') ; exit ; end ;

    // La vérification de déjà-existence du nom se fait également toute seule...

    if GetField ('C_NUMEROCONTACT') = 0 then
      begin
      Q := OpenSql ('SELECT MAX(C_NUMEROCONTACT) FROM CONTACT WHERE C_TYPECONTACT="'+Type_Contact+'" AND C_AUXILIAIRE = "'+getfield ('C_AUXILIAIRE')+'"', True);
      if not Q.Eof then
         begin
         numero := Q.Fields[0].AsInteger;
         SetField ('C_NUMEROCONTACT', numero + 1);
         end else SetField ('C_NUMEROCONTACT', 1);
         Ferme (Q);
      end;
    If (GetField ('C_NOM')= '') then begin SetLastError(2,'C_NOM','') ; exit ; end ;
    if GetField('C_TELEPHONE') <> '' then
        SetField ('C_CLETELEPHONE',  CleTelephone (GetField ('C_TELEPHONE')))
    else
        SetField ('C_CLETELEPHONE', '');
    if GetField('C_FAX') <> '' then
       SetField('C_CLEFAX',CleTelephone(GetField('C_FAX')))
    else
       SetField('C_CLEFAX','');

    // $$$ JP 13/06/05
    if GetField('C_TELEX') <> '' then
        SetField ('C_CLETELEX', CleTelephone (GetField ('C_TELEX')))
    else
        SetField ('C_CLETELEX', ''); // $$$ JP 13/06/05
    // $$$ JP fin
    
{$IFDEF GCGC}
{$IFNDEF CCS3}
    NomChamp:=VerifierChampsObligatoires(Ecran,'');
    if NomChamp<>'' then
      begin
      NomChamp:=ReadTokenSt(NomChamp);
      SetFocusControl(NomChamp) ;
      LastError:=16; LastErrorMsg:=TexteMessage[LastError]+champToLibelle(NomChamp);
      exit;
      end;
{$ENDIF}
{$ENDIF}
{$IFDEF GRC}
    if (ds <> Nil) and (DS.State=dsInsert) then    //mcd 19/04/2006 ajout test Nil, sinon plante depuis Pgiside
       begin
       if (ctxGRC in V_PGI.PGIContexte) and (GetParamSocSecur ('SO_RTCREATMESSCONT',False)) and ((Type_Tiers='CLI') or (Type_Tiers='PRO')) then OutLookInsertContact;
       if
{$IFDEF GCGC}
       (VH_GC.GRFSeria = true) and
{$ENDIF}
       (GetParamSocSecur ('SO_RFCREATMESSCONT',False)) and ((Type_Tiers='FOU') ) then OutLookInsertContact;
       end;
{$ENDIF}
    if ModifLot then TFSaisieList(ecran).BFermeClick(nil);

  if (ds<>nil) and (not V_Pgi.SilentMode) and AlerteActive('C') then
    if (copy(TFSaisieList(Ecran).name,1,9) = 'YYCONTACT') and ( not ModifLot )  then
      if (DS.State<>dsInsert) then
        begin
        if not TestAlerteContact(CodeModification+';'+CodeModifChamps)then
           begin
           LastError:=99;
           exit;
           end;
        end
        else
           if not TestAlerteContact (CodeCreation) then
             begin
             LastError:=99;
             exit;
             end;

  TOBCONTACTBTP.GetEcran(ecran);
  TOBCONTACTBTP.SetString('BC1_NUMEROCONTACT',GetField ('C_NUMEROCONTACT'));
  TOBCONTACTBTP.SetAllModifie(true);
  TOBCONTACTBTP.InsertOrUpdateDB(false);
  IsModify := (Origin = 'RGPD');
end;

// Initialisation du n° de contact
procedure TOM_Contact.OnNewRecord;
var Q : TQuery;
begin
inherited;
if (Type_contact = 'T') and ((type_tiers = '') or (CodeTiers = '')) then
  begin
  Q:=OpenSql('SELECT T_NATUREAUXI,T_TIERS From TIERS WHERE T_AUXILIAIRE ="'+GetField('C_AUXILIAIRE')+'"',True);
  if Not Q.EOF then
    begin
    if type_tiers = '' then type_tiers:=Q.Fields[0].AsString;
    if CodeTiers = '' then CodeTiers:=Q.Fields[1].AsString;
    end;
  Ferme(Q);
  end;
SetField('C_NUMEROCONTACT',LastNum+1);
SetField ('C_NATUREAUXI', type_tiers);
SetField ('C_TYPECONTACT',Type_Contact);
SetField('C_TIERS',CodeTiers);
if Assigned(GetControl('C_GUIDPER')) then SetField('C_GUIDPER',GuidPer);    //mcd 12/2005
// spécifique MODE /////////////////
if (ctxMode in V_PGI.PGIContexte) and (NomContact <> '') then
   SetField ('C_NOM', NomContact);
{$ifdef GIGI}
 if CodeTiers <>'' then AffectValTiers(CodeTiers);
{$endif}
end;

// Fonction vérifiant si le contact actuellement saisi dans la fiche est le seul
//  contact principal de l'auxiliaire (dans le cas où principal est coché)
function TOM_Contact.PrincipalUnique : Boolean;
begin
  if not FromSaisie then
  begin
    if not (GetField('C_PRINCIPAL') = 'X')
      then result := true
      else result := not ExisteSQL('SELECT C_PRINCIPAL FROM CONTACT '+
                           'WHERE C_TYPECONTACT="'+Type_Contact+'" AND C_AUXILIAIRE="'+String(GetField('C_AUXILIAIRE'))+'" '+
                             'AND C_PRINCIPAL="X" '+
                             'AND C_NUMEROCONTACT<>'+String(GetField('C_NUMEROCONTACT')));
  end
  else
  begin
    if not TDBCheckBox(GetControl('C_PRINCIPAL')).Checked
      then result := true
      else result := not ExisteSQL('SELECT C_PRINCIPAL FROM CONTACT '+
                           'WHERE C_TYPECONTACT="'+Type_Contact+'" AND C_AUXILIAIRE="'+String(GetField('C_AUXILIAIRE'))+'" '+
                             'AND C_PRINCIPAL="X" '+
                             'AND C_NUMEROCONTACT<>'+String(GetField('C_NUMEROCONTACT')));
  end;
end;

// Recherche du plus grand n° de contact pour l'auxiliaire courant
function TOM_Contact.LastNum : LongInt;
var Q : TQuery;
begin
{$ifdef GIGI}
if (Type_Contact='ANN') then
Q := OpenSQL('SELECT MAX(C_NUMEROCONTACT) FROM CONTACT '+
             'WHERE C_TYPECONTACT="'+Type_Contact+'" AND C_AUXILIAIRE="'+String(GetField('C_AUXILIAIRE'))+'"',true)
else
Q := OpenSQL('SELECT MAX(C_NUMEROCONTACT) FROM CONTACT '+
             'WHERE C_TYPECONTACT="'+Type_Contact+'" AND C_AUXILIAIRE="'+CodeAuxiliaire+'"',true);      //mcd 07/09/2006 12599
{$else}
Q := OpenSQL('SELECT MAX(C_NUMEROCONTACT) FROM CONTACT '+
             'WHERE C_TYPECONTACT="'+Type_Contact+'" AND C_AUXILIAIRE="'+String(GetField('C_AUXILIAIRE'))+'"',true);
{$endif}
if not Q.EOF then result := Q.Fields[0].AsInteger
             else result := -1;
Ferme(Q);
if (result <= 0) then result := 9;   // laisser de la marge pour les créations automatiques  
end;

procedure TOM_CONTACT.OnLoadAlerte  ;
begin
inHerited;
if (not V_Pgi.SilentMode) and AlerteActive('C') and (not AfterInserting) then
    begin
    FirstTime:=false;
    if (copy(TFSaisieList(Ecran).name,1,9) = 'YYCONTACT') and ( not ModifLot )  then
      TestAlerteContact(CodeOuverture+';'+CodeDateAnniv);
    end;
end;

procedure TOM_CONTACT.OnLoadRecord  ;
var
  i,NoRec : integer;
begin
  SetEvents (false);
  {$IFDEF CTI}
  if StTiers = '' then
    begin
    StTiers:=TiersAuxiliaire (GetField('C_AUXILIAIRE'), True);
    SetControlText ('TIERS',StTiers);
    end;
  if (RTFormCti=Ecran) and (AppelCtiOk=True) then
     // cas du 'decroche;COUPLAGE' = Bouton Appel Sortant
     CtiNumAct:=RTCTIGenereAction(True,StTiers,GetField('C_AUXILIAIRE'),CtiHeureDeb,0,'',CtiModeAppel,GetField('C_NUMEROCONTACT'));
  {$ENDIF}
  {$ifdef GIGI}
  if CodeAuxiliaire =''  then CodeAuxiliaire :=getcontrolText ('C_AUXILIAIRE'); //mcd 21/05/2007 14077 pas OK si acces depuis mul contact
  {$endif}
  if FromSaisie and FirstTime then
   BEGIN
   FirstTime:=false;
   TF := TFSaisieList(Ecran).LeFiltre;
   if TF.state=dsBrowse then
      begin
      NoRec := 1;
      for i := 0 to TF.TOBFiltre.detail.count-1 do
         begin
         if TF.TOBFiltre.detail[i].GetValue('C_NOM') = SaisieNom then
            begin
            NoRec := i+1;
            break;
            end;
         end;
      TF.SelectRecord (NoRec);
      end;
   END ;

  inHerited;
  if GetField ('C_RVA') <>'' then
     begin
     SetControlEnabled ('BNEWMAIL',true);
     end
  else
     begin
     SetControlEnabled ('BNEWMAIL',false);
     end;
  SetControlText ('DATECREATION',FormatDateTime('dd mmmm yyyy',GetField ('C_DATECREATION')));
  SetControlText ('DATEMODIF',FormatDateTime('dd mmmm yyyy',GetField ('C_DATEMODIF')));
  AffContactIndisponible() ;                // Gestion des affichages relatifs à C_FERME
  if (GetField ('C_NUMEROCONTACT') = 1) and (Particulier = True) then
      begin
      SetControlEnabled('C_CIVILITE',False);
      SetControlEnabled('C_NOM',False);
      SetControlEnabled('C_PRENOM',False);
      SetControlEnabled('C_JOURNAIS',False);
      SetControlEnabled('C_MOISNAIS',False);
      SetControlEnabled('C_ANNEENAIS',False);
      SetControlEnabled('C_SEXE',False);
      SetControlEnabled('C_RVA',False);
      SetControlEnabled('C_TELEPHONE',False);
      SetControlEnabled('C_FAX',False);
      SetControlEnabled('C_TELEX',False);
      SetControlEnabled('C_PRINCIPAL',False);
      SetControlVisible('BDelete',False);
      end else
      begin
      if TFSaisieList(ecran).TypeAction <> TaConsult then
         begin
         SetControlEnabled('C_CIVILITE',True);
         SetControlEnabled('C_NOM',True);
         SetControlEnabled('C_PRENOM',True);
         SetControlEnabled('C_JOURNAIS',True);
         SetControlEnabled('C_MOISNAIS',True);
         SetControlEnabled('C_ANNEENAIS',True);
         SetControlEnabled('C_SEXE',True);
         SetControlEnabled('C_RVA',True);
         SetControlEnabled('C_TELEPHONE',True);
         SetControlEnabled('C_FAX',True);
         SetControlEnabled('C_TELEX',True);
         SetControlEnabled('C_PRINCIPAL',True);
         SetControlVisible('BDelete',True);
         end;
      end;
  if ModifLot then SetArguments(StSQL);
  If (Ecran<>nil) and (LibTitre<>'') then Ecran.Caption:= TraduireMemoire(LibTitre) ;
  if (Type_contact = 'GCL') then SetControlText ('TC_SERVICE', TraduireMemoire('&Société')) ;
  {$IFDEF GCGC}
  {$IFNDEF CCS3}
  AppliquerConfidentialite(Ecran,'');
  {$ENDIF}
  {$ENDIF}
  if (ctxscot in V_PGI.PgiContexte) and (ecran<>Nil)and (Guidper <>'' )then      //mcd 28/03/2007 13847 ajout test guiper pour avoir bouton que si appel depuis tiers
      begin //mcd 22/02/2006 12599 GAGI
      if Assigned (GetControl('BCONTACT'))  then
       if ds.state = DsBrowse then
        begin
        SetControlVisible ('BCONTACT',true);
        SetcontrolEnabled('BCONTACT',true);
        end
        else SetControlVisible ('BCONTACT',false);
      end;
    //
  FindContactBTP (GetField('C_TYPECONTACT'),GetField('C_AUXILIAIRE'),GetField('C_NUMEROCONTACT'));
  TOBCONTACTBTP.PutEcran(ecran);
  SetEvents (true);
  if Origin = 'RGPD' then
  begin
    THGrid(GetControl('FLISTE')).Enabled := False;
    SetControlVisible('BDelete', False);
  end;
  SetButtonEna;
end;

procedure TOM_CONTACT.OnClose ;
begin
inherited;
  TOBCONTACTBTP.free;
{$IFDEF CTI}
// Appels en série : on sort sans avoir raccroché : on empêche de sortir
if (AppelCtiOk) and (SerieCTI) and (not CtiRaccrocherFiche) then
    begin
    LastError := 1;
    exit;
    end ;
{$ENDIF}
{$IFDEF CTI}
  if not ( (AppelCtiOk) and (SerieCTI) and (not CtiRaccrocherFiche) ) then
    begin
{ CRM_20080901_MNG_FQ;012;10843_FIN }
    // Serie Appels CTI : pas d'Appel, on génère l'action
    if (GetParamSocSecur('SO_RTCTIGESTION',False)) and (SerieCTI) and (not AppelCtiOk) then
        RTCTIGenereAction(AppelCtiOk,GetControlText('TIERS'),GetField('C_AUXILIAIRE'),CtiHeureDebCom,CtiHeureFinCom-CtiHeureDebCom,'',CtiModeAppel,GetField('C_NUMEROCONTACT'));
    // Appels Sortant seul : on sort sans avoir raccroché : on mémorise l'heure de début de Communication
    if (AppelCtiOk) and (not SerieCTI) and (not CtiRaccrocherFiche) then
        CtiHeureDeb:=CtiHeureDebCom;
    end;

  if RTFormCti<> Nil then RTFormCti:=Nil;
  // Pour affichage infos CTI, on retourne Code Tiers et Raison sociale
  if ContexteCti then
    TFSaisieList(Ecran).Retour:=GetControlText('TIERS')+'|'+GetField('C_PRENOM')+' '+GetField('C_NOM') ;
{$ENDIF CTI}

{$IFDEF BUREAU}
{$IFDEF EAGLCLIENT}                   //TJA 11/12/2007   //FQ 11889
If GetControlText('C_AUXILIAIRE') <> '' Then
{$ELSE}
if GetField('C_AUXILIAIRE')<>'' then
{$ENDIF EAGLCLIENT}
   TFSaisieList(Ecran).Retour:= GetField('C_TYPECONTACT')+';'+GetField('C_AUXILIAIRE')+';'+IntToStr(GetField('C_NUMEROCONTACT'));
{$ENDIF}
  if Assigned( AlerteInitTob) then AlerteInitTob.free;
  if Origin = 'RGPD' then
  begin
    if IsModify then
      TFSaisieList(Ecran).Retour := GetField('C_TYPECONTACT')+';'+GetField('C_AUXILIAIRE')+';'+IntToStr(GetField('C_NUMEROCONTACT'))
    else
      TFSaisieList(Ecran).Retour := '';
  end;
end;

procedure TOM_CONTACT.OnChangeField(F: TField);
begin
Inherited;
{$IFDEF CTI}
if CtiErreurConnexion then
    RTCTIAfficheMessageCTI;
if CtiRaccrocherFiche then
    SetControlText('INFOSCTI',TraduireMemoire('le correspondant a raccroché'));
{$ENDIF}
// spécifique MODE ////////
if (ctxMode in V_PGI.PGIContexte) then
   begin
   if (GetField ('C_NUMEROCONTACT') = 1) and (Particulier = True) then
        begin
        NomContact := GetField ('C_NOM');
        SetControlEnabled('C_CIVILITE',False);
        SetControlEnabled('C_NOM',False);
        SetControlEnabled('C_PRENOM',False);
        SetControlEnabled('C_JOURNAIS',False);
        SetControlEnabled('C_MOISNAIS',False);
        SetControlEnabled('C_ANNEENAIS',False);
        SetControlEnabled('C_SEXE',False);
        SetControlEnabled('C_RVA',False);
        SetControlEnabled('C_TELEPHONE',False);
        SetControlEnabled('C_FAX',False);
        SetControlEnabled('C_TELEX',False);
        end else
        begin
        if TFSaisieList(ecran).TypeAction <> TaConsult then
           begin
           SetControlEnabled('C_CIVILITE',True);
           SetControlEnabled('C_NOM',True);
           SetControlEnabled('C_PRENOM',True);
           SetControlEnabled('C_JOURNAIS',True);
           SetControlEnabled('C_MOISNAIS',True);
           SetControlEnabled('C_ANNEENAIS',True);
           SetControlEnabled('C_SEXE',True);
           SetControlEnabled('C_RVA',True);
           SetControlEnabled('C_TELEPHONE',True);
           SetControlEnabled('C_FAX',True);
           SetControlEnabled('C_TELEX',True);
           end;
        end;
   end;

// cd 22/03/04 Suppression automatisme de MAJ entre combo et libellé
   {  if      (F.FieldName = 'C_FONCTIONCODEE') and (F.Value <> '') then
  begin
    if (GetField('C_FONCTION')='') then
      SetField('C_FONCTION',RECHDOM('TTFONCTION',GetField('C_FONCTIONCODEE'),FALSE));
  end
  else if (F.FieldName = 'C_SERVICECODE')   and (F.Value <> '') then
  begin
    if GetField('C_SERVICE')='' then
      SetField('C_SERVICE',RECHDOM('YYSERVICE',GetField('C_SERVICECODE'),FALSE));
  end;   }
end;

procedure TOM_Contact.OnDeleteRecord  ;
var Q : TQuery;
    StSql,StOr : string;
begin
Inherited ;
  if GetField('C_TYPECONTACT') <> 'T' then exit;

  //Controle pour GRC
  if ExisteSQL('SELECT RAC_LIBELLE FROM ACTIONS WHERE RAC_AUXILIAIRE="'
       +GetField('C_AUXILIAIRE')+'" AND (RAC_NUMEROCONTACT='+IntToStr(GetField('C_NUMEROCONTACT'))
       +' OR RAC_DESTMAIL like "%;'+IntToStr(GetField('C_NUMEROCONTACT'))+';%")') then
     begin
     SetLastError(3,'','') ;
     exit ;
     end ;
  if ExisteContact ('PERSPECTIVES','RPE_AUXILIAIRE','RPE_NUMEROCONTACT','A',4) then exit;
  if ExisteContact ('PERSPHISTO','RPH_AUXILIAIRE','RPH_NUMEROCONTACT','A',7) then exit;
  if ExisteContact ('PROJETS','RPJ_AUXILIAIRE','RPJ_NUMEROCONTACT','A',5) then exit;
  // Vérification existence contact ds infos complémentaires prospects
  Q:=OpenSQL('SELECT RCL_NOMCHAMP From CHAMPSPRO WHERE RCL_TYPETEXTE="C"',True,-1,'',True);
  if not Q.Eof then
     begin
     StSql := 'SELECT RPR_AUXILIAIRE from PROSPECTS WHERE RPR_AUXILIAIRE ="' + GetField('C_AUXILIAIRE')+'" AND (';
     StOr := '';
     while Not Q.EOF do
        begin
        StSql := StSql + StOr + Q.FindField('RCL_NOMCHAMP').asstring + ' = "'+IntToStr(GetField('C_NUMEROCONTACT')) + '"';
        StOr := ' OR ';
        Q.Next;
        end;
     StSql := StSql + ')';
     if ExisteSQL(StSql) then
        begin
        SetLastError(8,'','') ;
        Ferme(Q) ;
        exit ;
        end ;
     end;
  Ferme(Q) ;
  // Vérification existence contact ds infos complémentaires actions
  if ExisteContact_InfosCompl ('RD1','1',9) then exit;
  if ExisteContact_InfosCompl ('RD3','3',17) then exit;

  // Vérification existence contact ds les adresses
  if ExisteContact('ADRESSES', 'ADR_REFCODE', 'ADR_NUMEROCONTACT', 'T', 13) then exit;

  // Vérification existence contact ds les compléments d'entête pièce
  if GetParamSocSecur('SO_GCPIECEADRESSE',True) then
  begin
    if ExisteSQL('SELECT GPA_NUMEROCONTACT FROM PIECEADRESSE'
                + ' LEFT JOIN PIECE ON (GPA_NATUREPIECEG = GP_NATUREPIECEG'
                                      + ' AND GPA_SOUCHE = GP_SOUCHE'
                                      + ' AND GPA_NUMERO =  GP_NUMERO'
                                      + ' AND GPA_INDICEG = GP_INDICEG)'
                + ' WHERE ((GP_TIERSLIVRE = "'+ TiersAuxiliaire(GetField('C_AUXILIAIRE'),True) + '"'
                + ' AND GPA_TYPEPIECEADR="001"'
                + ' AND GPA_NUMEROCONTACT = ' + IntToStr(GetField('C_NUMEROCONTACT')) + ')'
                + ' OR (GP_TIERSFACTURE = "' + TiersAuxiliaire(GetField('C_AUXILIAIRE'),True) + '"'
                + ' AND GPA_TYPEPIECEADR="002"'
                + ' AND GPA_NUMEROCONTACT = ' + IntToStr(GetField('C_NUMEROCONTACT')) + '))' ) then
    begin
      SetLastError(14,'','');
      exit;
    end
  end
  else
    if ExisteContact ('ADRESSES','ADR_REFCODE','ADR_NUMEROCONTACT','T',14) then exit;

  //Controle pour AFFAIRES
  if ExisteContact ('AFFAIRE','AFF_TIERS','AFF_NUMEROCONTACT','T',6) then exit;
  //Controle pour AFFTIERS intervention dans une affaire
  StSql :='SELECT AFT_AFFAIRE FROM AFFTIERS WHERE AFT_TYPEINTERV ="CON"'
        + ' AND AFT_NUMEROCONTACT = ' + IntToStr(GetField('C_NUMEROCONTACT'))
        + ' AND AFT_AUXILIAIRE = "' + GetField('C_AUXILIAIRE') + '"';
  Q:=OpenSQL(StSql,True);
  if not Q.Eof then
  begin
    SetLastError(18,'',Q.Fields [0].asstring);
    Ferme (Q);
    exit;
  end;
  Ferme (Q);
  //Controle pour CHR
  if ExisteContact ('HRCONTRAT','HCO_TIERS','HCO_NUMEROCONTACT','T',10) then exit;
  if ExisteContact ('HRDOSSIER','HDC_TIERS','HDC_NUMEROCONTACT','T',11) then exit;
  if ExisteContact ('HRDOSRES','HDR_TIERS','HDR_NUMEROCONTACT','T',12) then exit;
  if ((ctxGRC in V_PGI.PGIContexte) and (StTiers <> '') and ((type_tiers = 'CLI') or (type_tiers = 'PRO')))
  {$IFDEF CTI}
          or (ContexteCti)
  {$ENDIF}
     then
     begin
     if TFSaisieList(Ecran).LeFiltre.TOBFiltre.detail.count-1 = 0 then
       begin
       SetControlEnabled ('BACTION',False);
       SetControlEnabled ('BCOMPLGRC',False);
       end;
     end;
{ GC/GRC : MNG / gestion des alertes }
  if (not V_Pgi.SilentMode) and (not AfterInserting)
    and ( not ModifLot ) and (AlerteActive (TableToPrefixe(TableName))) then
      if (not TestAlerteContact (CodeSuppression) ) then
        begin
        LastError := 99;
        exit;
        end;
{ fin GC/GRC : MNG / gestion des alertes }

{$IFDEF GRC}
  if GetParamSocSecur('SO_RTGESTINFOS006',False) = True then
     ExecuteSQL('DELETE FROM RTINFOS006 where RD6_CLEDATA="T;'+GetField('C_AUXILIAIRE')+';'+IntToStr(GetField('C_NUMEROCONTACT'))+'"') ;
{$ENDIF}
  DeleteContactBTP (GetField('C_TYPECONTACT'),GetField('C_AUXILIAIRE'),GetField('C_NUMEROCONTACT'));
end;

function TOM_Contact.ExisteContact (NomTable, NomChamp1, NomChamp2, TypeRech: string ; NumErr : integer) : boolean;
begin
  if TypeRech = 'A' then
    result := ExisteSQL('SELECT ' + NomChamp1 +
                        '  FROM ' + NomTable  +
                        ' WHERE ' + NomChamp1 + ' = "' + GetField('C_AUXILIAIRE') +
                        '"  AND ' + NomChamp2 + ' ='   +IntToStr(GetField('C_NUMEROCONTACT')))
  else
    result := ExisteSQL('SELECT ' + NomChamp1 +
                        '  FROM ' + NomTable  +
                        ' WHERE ' + NomChamp1 + ' = "' + TiersAuxiliaire(GetField('C_AUXILIAIRE'),True) +
                        '"  AND ' + NomChamp2 + '= '   + IntToStr(GetField('C_NUMEROCONTACT')));

  if result then SetLastError(NumErr,'','') ;

end;

function TOM_Contact.ExisteContact_InfosCompl (Prefixe, NoDesc: string ; NumErr : integer) : boolean;
var Q : TQuery;
    StSql,StOr,NomChamp : string;
begin
Result := False;
Q:=OpenSQL('SELECT RDE_NOMCHAMP From RTINFOSDESC WHERE RDE_DESC = "'+ NoDesc + '" AND RDE_TYPETEXTE="C"',True,-1,'',True);
if not Q.Eof then
   begin
   if Prefixe = 'RD1' then StSql := 'SELECT RD1_CLEDATA from RTINFOS001 WHERE RD1_CLEDATA like"' + GetField('C_AUXILIAIRE')+';%" AND ('
   else StSql := 'SELECT RD3_CLEDATA from RTINFOS003 WHERE RD3_CLEDATA ="' + GetField('C_AUXILIAIRE')+'" AND (';
   StOr := '';
   while Not Q.EOF do
      begin
      NomChamp := FindEtReplace(Q.FindField('RDE_NOMCHAMP').asstring,'RPR',Prefixe,True);
      StSql := StSql + StOr + NomChamp + ' = "'+IntToStr(GetField('C_NUMEROCONTACT')) + '"';
      StOr := ' OR ';
      Q.Next;
      end;
   StSql := StSql + ')';
   if ExisteSQL(StSql) then
      begin
      Result := True;
      SetLastError(NumErr,'','') ;
      end ;
   end;
Ferme(Q) ;
end;

procedure TOM_Contact.OutLookInsertContact;
var nomPrenom : string ;
begin
// Fin VERRUE
Try
LaToc:=Toc.create(True);
except
  LaToc := nil;
end;
if assigned(LaToc) then
  begin
  LaToc.clear;

   // LaToc.PutValue ('SOL_ORGANIZATIONALIDNUMBER',  GetField('C_AUXILIAIRE')) ;
   // LaToc.PutValue ('SOL_CUSTOMERID',  GetField('C_TYPECONTACT')+';'+GetField('C_AUXILIAIRE')+';'+intToStr(GetField('C_NUMEROCONTACT'))) ;
    nomPrenom :=trim(GetField('C_PRENOM')+' '+ GetField('C_NOM'));
//    LaToc.PutValue ('SOL_COMPANYNAME', GetField('C_SERVICE')) ;
    LaToc.PutValue ('SOL_DEPARTMENT', GetField('C_SERVICE')) ;
    LaToc.PutValue ('SOL_LASTNAME', GetField('C_NOM')) ;
    LaToc.PutValue ('SOL_FULLNAME', NomPrenom) ;
    LaToc.PutValue ('SOL_FIRSTNAME', GetField('C_PRENOM')) ;
    LaToc.PutValue ('SOL_EMAIL1ADDRESS', GetField('C_RVA')) ;
    LaToc.PutValue ('SOL_PROFESSION', GetField('C_FONCTION')) ;
    LaToc.PutValue ('SOL_BUSINESSTELEPHONENUMBER', GetField('C_TELEPHONE')) ;
    LaToc.PutValue ('SOL_BUSINESSFAXNUMBER', GetField('C_FAX')) ;
    LaToc.PutValue ('SOL_MOBILETELEPHONENUMBER', GetField('C_TELEX')) ;
  LaToc.Insert;Latoc.free;
  end;
end;

Procedure TOM_Contact.AffContactIndisponible ;
//var Ctrl : TControl;
begin
if (GetField('C_FERME')='X') then
    begin
    SetControlText ('DATEFERMETURE',FormatDateTime('dd mmmm yyyy',GetField ('C_DATEFERMETURE')));
    SetControlVisible('DATEFERMETURE',True);
    SetControlVisible('TC_DATEFERMETURE',True);
//    TCheckBox(GetControl('C_FERME')).Font.Color:= clRed;
    end else
    begin
    SetControlText ('DATEFERMETURE',FormatDateTime('dd mmmm yyyy',GetField ('C_DATEFERMETURE')));
    SetControlVisible('DATEFERMETURE',False);
    SetControlVisible('TC_DATEFERMETURE',False);
//    TCheckBox(GetControl('C_FERME')).Font.Color:= clBlack;
    end;
end;

procedure TOM_Contact.FermeContact(Sender: TObject);
begin
  if TCheckBox(GetControl('C_FERME')).Checked then
  begin
    SetField('C_FERME','X') ;
    SetField('C_DATEFERMETURE', V_PGI.DateEntree);
  end else
  begin
    SetField('C_FERME','-') ;
    Setfield('C_DATEFERMETURE',iDate1900) ;
  end;
AffContactIndisponible() ;
if (TCheckBox(GetControl('C_FERME')).Checked) and (GetField('C_PRINCIPAL') = 'X') then
   PGIBox(TraduireMemoire(TexteMessage[15]),'Attention!');
end;

procedure TOM_Contact.AddRemoveItemFromPopup(stPopup,stItem : string; bVisible : boolean);     //PCS
var pop : TPopupMenu ;
    i : integer;
    st : string;
begin
pop :=TPopupMenu(GetControl(stPopup) ) ;
if pop<> nil  then
   for i:=0 to pop.items.count-1 do
       begin
       st:=uppercase(pop.items[i].name);
       if st=stItem then  begin pop.items[i].visible:=bVisible; break; end;
       end;
end ;

function TOM_Contact.TestAlerteContact (CodeA : String) : boolean;
var TobContact,TobAlertes,TOBInfosCompl : tob;
    i : integer;
    stCle : String;
    Q : TQuery;
begin
  Result := true;
  { mng 27/09/2007 : cas particulier de la fiche contact ou l'on peut arrivé sans contact, mais pas en mode création}
  if (GetField('C_AUXILIAIRE') = '') or (GetField('C_NUMEROCONTACT') = 0) then exit;
  TobContact:=TOB.Create ('CONTACT', Nil, -1);
  TobContact.GetEcran (TFSaisieList(Ecran),Nil);
  TobAlertes:=TOB.create ('les alertes',NIL,-1);
  TobAlertes.Dupliquer(TobContact, False, True);
  TobContact.free;
  if GetParamSocSecur('SO_RTGESTINFOS006',false) then
    begin
    TOBInfosCompl:= TOB.Create('RTINFOS006', nil, -1);
    if CodeA<>CodeCreation then
      begin
      stCle:='"T;'+GetField('C_AUXILIAIRE')+';'+IntToStr(GetField('C_NUMEROCONTACT'))+ '"';
      Q := OpenSQL('Select * from RTINFOS006 Where RD6_CLEDATA='+stCle,True) ;
      if Not Q.EOF then
        TOBInfosCompl.SelectDB('',Q);
      Ferme(Q);
      end;
    for i := 1 to Pred(TOBInfosCompl.NbChamps) do
      TobAlertes.AddChampSupValeur(TOBInfosCompl.GetNomChamp(i), TOBInfosCompl.GetValue(TOBInfosCompl.GetNomChamp(i)));
    if (pos(CodeOuverture,CodeA) > 0) then
      begin
      if assigned( AlerteInitTob) then AlerteInitTob.free;
      AlerteInitTob:=TOB.create ('les initiales',NIL,-1);
      AlerteInitTob.Dupliquer(TobAlertes, False, True);
      end;
    TOBInfosCompl.free;
    end;
  result:=FMonAlerte.ExecuteAlerte(Ecran,'C',CodeA,TobAlertes,AlerteInitTob);
  TobAlertes.free;
end;

procedure TOM_Contact.ListAlerte_OnClick_C(Sender: TObject);
begin
if (GetField('C_NUMEROCONTACT') <> 0) then
   AGLLanceFiche('Y','YALERTES_MUL','YAL_PREFIXE=C','','ACTION=CREATION;MONOFICHE;CHAMP=C_TYPECONTACT|C_AUXILIAIRE|C_NUMEROCONTACT;VALEUR='
      +GetField('C_TYPECONTACT')+'|'+GetField('C_AUXILIAIRE')+'|'+IntToStr(GetField('C_NUMEROCONTACT'))
      +';LIBELLE='+GetField('C_NOM')) ;
end ;

procedure TOM_Contact.Alerte_OnClick_C(Sender: TObject);
begin
  if (GetField('C_NUMEROCONTACT') <> 0) then
     AGLLanceFiche('Y','YALERTES','','','ACTION=CREATION;MONOFICHE;CHAMP=C_TYPECONTACT|C_AUXILIAIRE|C_NUMEROCONTACT;VALEUR='
      +GetField('C_TYPECONTACT')+'|'+GetField('C_AUXILIAIRE')+'|'+IntToStr(GetField('C_NUMEROCONTACT'))
      +';LIBELLE='+GetField('C_NOM')) ;
  VH_EntPgi.TobAlertes.ClearDetail;
end;

{$IFDEF BTP}
procedure TOM_CONTACT.FLISTE_OnDblClick(Sender: TObject);
begin
  TFSaisieList(Ecran).retour := GetField('C_NUMEROCONTACT');
  TFSaisieList(Ecran).Close;
end;
{$ENDIF}

procedure AGLRTOLKInsertContact( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFSaisieList) then OM:=TFSaisieList(F).OM else exit;
if (OM is TOM_Contact) then TOM_Contact(OM).OutLookInsertContact else exit;
end;
{$IFDEF CTI}
procedure TOM_CONTACT.RTCTIDecrocheAppelContact ;
begin
Fo_PBC.DecrocherClick (TForm (Ecran)) ;
CtiHeureDebCom := Time;
AppelCtiOk:=True;
Fo_PBC.AfficheCoordonnees (true,GetControlText('TIERS'),'',GetField('C_NATUREAUXI'),GetField('C_AUXILIAIRE'),GetField('C_NUMEROCONTACT'),GetField('C_PRENOM')+' '+GetField('C_NOM')) ;
CtiNumAct:=RTCTIGenereAction(AppelCtiOk,GetControlText('TIERS'),GetField('C_AUXILIAIRE'),CtiHeureDebCom,0,WhereUpdate,CtiModeAppel,GetField('C_NUMEROCONTACT'));
end;

procedure TOM_CONTACT.RTCTIRaccrocheAppelContact ;
begin
CtiHeureFinCom:=Time;
SetControlProperty('BPHONE','Down',False);
Fo_PBC.RaccrocherClick (TForm (Ecran)) ;
// Serie ou Appel Sortant Seul : on génère l'action
// même sur appel entrant if (SerieCTI) or (CtiModeAppel=1) then
//RTCTIGenereAction(AppelCtiOk,GetField('T_TIERS'),GetField('T_AUXILIAIRE'),CtiHeureDebCom,CtiHeureFinCom-CtiHeureDebCom,WhereUpdate,CtiModeAppel);
RTCTIMajDureeAction (GetField('C_AUXILIAIRE'),CtiNumAct,CtiHeureFinCom-CtiHeureDebCom);
CtiHeureDeb:=0;
end;
procedure TOM_CONTACT.RTCTINumeroterContact ;
begin
if GetField('C_CLETELEPHONE') <> '' then
    begin
    Fo_PBC.RTCTINumeroteClient (TForm (Ecran),GetField('C_CLETELEPHONE'),GetControlText('TIERS'),'',GetField('C_NATUREAUXI'),GetField('C_AUXILIAIRE')) ;
    if CtiErreurConnexion=True then
        begin
        RTCTIAfficheMessageCTI;
        SetControlProperty('BPHONE','Down',False);
        CtiErreurConnexion:=False;
        //SetControlVisible('BPHONE',False);
        end
    else
        begin
        if not CtiCCA then
          begin
          SetControlProperty('BPHONE','Hint',TraduireMemoire('Raccrocher'));
          SetControlProperty('BPHONE','Down',True);
          SetControlText('INFOSCTI',TraduireMemoire('Communication en cours'));
          end
        else
          SetControlVisible('BATTENTE',False);
        AppelCtiOk:=True;
        CtiHeureDebCom := Time;
        Fo_PBC.AfficheCoordonnees (False,GetControlText('TIERS'),'',GetField('C_NATUREAUXI'),GetField('C_AUXILIAIRE'),GetField('C_NUMEROCONTACT'),GetField('C_PRENOM')+' '+GetField('C_NOM')) ;
        CtiNumAct:=RTCTIGenereAction(AppelCtiOk,GetControlText('TIERS'),GetField('C_AUXILIAIRE'),CtiHeureDebCom,0,WhereUpdate,CtiModeAppel,GetField('C_NUMEROCONTACT'));
        end;
    end
else
    SetControlProperty('BPHONE','Down',False);
end;

procedure TOM_CONTACT.RTCTIAppelAttenteContact ;
begin
Fo_PBC.AttenteClick (TForm (Ecran)) ;
end;
procedure TOM_CONTACT.RTCTIRepriseAppelContact ;
begin
Fo_PBC.RepriseClick (TForm (Ecran)) ;
end;


procedure TOM_CONTACT.RTCTIAfficheMessageCTI;
var MessageCti: string;
begin
MessageCti := TraduireMemoire('Communication Impossible');
case TResultatAppel(CtiResultatAppel) of
    raOccupe  : MessageCti:=MessageCti+' : '+TraduireMemoire('corresp. occupé');
    raRefus   : MessageCti:=MessageCti+' : '+TraduireMemoire('Appel refusé');
    raAbort   : MessageCti:=MessageCti+' : '+TraduireMemoire('Appel interrompu');
    end;
SetControlText('INFOSCTI',MessageCti);
end;


procedure AGLRTCTIDecrocheAppelContact( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFSaisieList) then OM:=TFSaisieList(F).OM else exit;
if (OM is TOM_CONTACT) then TOM_CONTACT(OM).RTCTIDecrocheAppelContact() else exit;
end;

procedure AGLRTCTINumeroterContact( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFSaisieList) then OM:=TFSaisieList(F).OM else exit;
if (OM is TOM_CONTACT) then TOM_CONTACT(OM).RTCTINumeroterContact() else exit;
end;

procedure AGLRTCTIRaccrocheAppelContact( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFSaisieList) then OM:=TFSaisieList(F).OM else exit;
if (OM is TOM_CONTACT) then TOM_CONTACT(OM).RTCTIRaccrocheAppelContact() else exit;
end;

procedure AGLRTCTIAppelAttenteContact( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFSaisieList) then OM:=TFSaisieList(F).OM else exit;
if (OM is TOM_CONTACT) then TOM_CONTACT(OM).RTCTIAppelAttenteContact() else exit;
end;

procedure AGLRTCTIRepriseAppelContact( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFSaisieList) then OM:=TFSaisieList(F).OM else exit;
if (OM is TOM_CONTACT) then TOM_CONTACT(OM).RTCTIRepriseAppelContact() else exit;
end;
{$ENDIF}

{$IFDEF GRC}
procedure TOM_CONTACT.RTAppelParamCLContact;
var StAction : string;
    TobChampsProFille : tob;
begin
if ( TToolbarButton97(GetControl('BCOMPLGRC')).Enabled = false ) then exit;

  VH_RT.TobChampsPro.Load;

TobChampsProFille:=VH_RT.TobChampsPro.FindFirst(['CO_CODE'], ['6'], TRUE);
if (TobChampsProFille = Nil ) or (TobChampsProFille.detail.count = 0 ) then
    begin
    PGIInfo(TraduireMemoire('Le paramétrage de cette saisie n''a pas été effectué'),'');
    exit;
    end;
StAction:='';
if (not self.ModifAutorisee) or (TFSaisieList(ecran).TypeAction=taConsult) then
   StAction:='ACTION=CONSULTATION;';
if (Ecran <> Nil) and ( Ecran is TFSaisieList) then
    begin
    { mng gestion des alertes : qu'en creation }
    if DS.State = dsInsert then
        TFSaisieList(Ecran).LeFiltre.Post      ;
    AglLancefiche('RT','RTPARAMCL','','',StAction+'FICHEPARAM=YYCONTACT;FICHEINFOS='+GetField('C_TYPECONTACT')+'|'+GetField('C_AUXILIAIRE')+'|'+IntToStr(GetField('C_NUMEROCONTACT'))) ;
    { il faut valider la fiche pour que les modifs soit prise en compte dans le grid}
    TFSaisieList(Ecran).LeFiltre.Edit;
    TFSaisieList(Ecran).LeFiltre.Post;
    end;
end;

procedure AGLRTAppelParamCLContact( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
if (OM is TOM_Contact) then TOM_Contact(OM).RTAppelParamCLContact else exit;
end;
{$ENDIF GRC}

{$IFDEF GCGC}
procedure TOM_CONTACT.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
Case Key of
    VK_F6 : {Infos compl.} if (ssAlt in Shift) then
            if ( GetParamSocSecur('SO_RTGESTINFOS006',False) = True) and (type_tiers <> 'FOU') then RTAppelParamCLContact ;
    VK_F7 : {Actions} if (ssAlt in Shift) then
                       begin Key:=0 ;
                       if (Type_contact = 'T') and ((type_tiers = 'CLI') or (type_tiers = 'PRO') or (type_tiers = '')) then
                           AGLLanceFiche('RT','RTACT_MUL_CONTACT','RAC_TIERS='+GetField('C_TIERS')+';NUMEROCONTACT='+IntToStr(GetField('C_NUMEROCONTACT')),'',ActionToString(TFSaisieList(Ecran).FtypeAction)+';NOCHANGEPROSPECT') ;
                       if (Type_contact = 'T') and (type_tiers = 'FOU') then
                           AGLLanceFiche('RT','RFACT_MUL_CONTACT','RAC_TIERS='+GetField('C_TIERS')+';NUMEROCONTACT='+IntToStr(GetField('C_NUMEROCONTACT')),'',ActionToString(TFSaisieList(Ecran).FtypeAction)+';NOCHANGEPROSPECT') ;
                       end;
    VK_F8 : {Propositions} if (ssAlt in Shift) then
                       begin Key:=0 ;
                       if (Type_contact = 'T') and ((type_tiers = 'CLI') or (type_tiers = 'PRO')) then
                         AGLLanceFiche('RT','RTPERSP_MUL_TIERS','RPE_TIERS='+GetField('C_TIERS')+';RPE_NUMEROCONTACT='+IntToStr(GetField('C_NUMEROCONTACT')),'',ActionToString(TFSaisieList(Ecran).FtypeAction)+';NOCHANGEPROSPECT') ;  end;
    81 : {Ctrl + Q - Création d'1 alerte} if (ssCtrl in Shift) then
          begin
          Key:=0 ;
          Alerte_OnClick_C(Sender);
          end;
    85 : {Ctrl + U - liste des alertes du tiers} if (ssCtrl in Shift) then
          begin
          Key:=0 ;
          ListAlerte_OnClick_C(Sender);
          end;

    END;
if (ecran <> nil) then
if ecran is TFSaisieList then
   TFSaisieList(ecran).FormKeyDown(Sender,Key,Shift);
end;
{$ENDIF}

// *****************************************************************************
// ********************** gestion Isoflex **************************************
// *****************************************************************************

procedure TOM_Contact.GereIsoflex;
var bIso: Boolean;
  MenuIso: TMenuItem;
begin
  MenuIso := TMenuItem(GetControl('mnSGED'));
  bIso := AglIsoflexPresent;
  if MenuIso <> nil then MenuIso.Visible := bIso;
end;


procedure TOM_Contact.BContactClick(Sender : TObject);
var  ret,TypeCont,Auxi, Num,StSql:string;
  TobForm:tob;
//	QQ: Tquery;
begin       //mcd 22/02/2006 12599
    //le bouton d'accès aux contacts, pour sélection d'un contact d'un autre client et réplication
    //dans la saisie en cours, n'est fait que pour la GI, et mis sur le bouton "phone", qui ne sera pa sutiliser en GI
    // car CTI pas gérer en GI, mais dans le bureau
		// ce bouton, permet de récupérer un contact d'un autre client, et de le dupliquer pour la saisie en cours
	if (ds.state=dsinsert) or (ds.state=dsEdit) then
		begin
		PgiInfo ('Vous devez valider vos modifications avant d''accéder à la sélection d''un contact',Ecran.caption);
		exit;
		end;
  ret := AGLLanceFiche('AFF', 'AFCONTACT_MUL', 'C_TYPECONTACT='+Type_Contact,'', 'CONSULTATION');

  //FQ 11920 - GHA 01/2008 ajout condition Or car si l'utilisateur valide la fiche AFCONTACT_MUL,
  //alors qu'il n'exite pas de contact, exécution du script sur le bouton OK
  //retour:=GetChamp('C_TYPECONTACT')+';'+GetChamp('C_AUXILIAIRE')+';'+GetChamp('C_NUMEROCONTACT');
  // du coup retour = ';;'
  if (ret = '') or (ret = ';;')
    then exit;

    //on recupère les valeurs
  TypeCont := readTokenst (Ret);
  Auxi := readTokenst (Ret);
  Num := readTokenst (Ret);
		//il faut tout prendre, car duplication
  StSql :='SELECT '+GetAllSelectFields('CONTACT')+' FROM CONTACT where C_TYPECONTACT="'+TypeCOnt+'" and C_AUXILIAIRE="'+Auxi+'" AND C_NUMEROCONTACT='+Num;
  TobForm := TOB.Create('CONTACT', nil, -1);
  TobForm.LoadDetailDBFromSQL('CONTACT',StSQL);
//  TOBLoadFromSelectDB('CONTACT',StSQL,TobForm) ;
//   QQ := OpenSQL(StSql, True);
//   TObForm.loadDetailDb ('CONTACT','','',QQ,false);
//   Ferme(QQ);
    //on change les champs de la clé. Un seul enrgt ans la tob
  TobForm. Detail[0].putValue ('C_PRINCIPAL','-'); //il ne faut pas récupérer l'info 'principal'
  TobForm. Detail[0].putValue ('C_TYPECONTACT',Type_Contact);
{$ifdef GIGI}
  TobForm. Detail[0].putValue ('C_AUXILIAIRE',CodeAuxiliaire);   //mcd 07/09/2006 12899
{$else}
  TobForm. Detail[0].putValue ('C_AUXILIAIRE',getfield ('C_AUXILIAIRE'));
  If   getfield ('C_AUXILIAIRE') ='' then
    TobForm. Detail[0].putValue ('C_AUXILIAIRE',TiersAuxiliaire(CodeTiers,false));
{$endif}
  TobForm. Detail[0].putValue ('C_NUMEROCONTACT',LastNum+1);
  TobForm. Detail[0].putValue ('C_NATUREAUXI', type_tiers);
  If Type_Contact='ANN' then TobForm. Detail[0].putValue ('C_AUXILIAIRE',GuidPer);
  TobForm. Detail[0].putValue('C_TIERS',CodeTiers);
  if Assigned(GetControl('C_GUIDPER')) then TobForm. Detail[0].putValue('C_GUIDPER',GuidPer);
	TobForm.setAllModifie(true);
	TobForm.InsertOrUpdateDb(false);
  TobForm.Free;
  TFSaisieList(Ecran).LeFiltre.RefreshLignes;
end;

procedure c_AppelIsoFlex(parms: array of variant; nb: integer);
var F: TForm;
  Cle1: string;
begin
  F := TForm(Longint(Parms[0]));
  //GP_DS_GP14537_20071226
  if not pos('YYCONTACT', UpperCase(F.Name)) = 1 then exit;
  Cle1 := 'T;'+string(Parms[1])+';'+string(Parms[2]);
  AglIsoflexViewDoc(NomHalley, F.Name, 'CONTACT', 'C_CLE1', 'C_TYPECONTACT,C_AUXILIAIRE,C_NUMEROCONTACT', Cle1, '');
end;

{$ifdef GIGI}
procedure TOM_Contact.AffectValTiers (CodeTiers: string) ;
var
  QQ: TQuery;
	ii : integer;
begin
		//fct qui va affecter les valeurs du tiers dans le contact si Ok en paramétrage
  if (TobValTiers=Nil)  then exit;
  QQ := nil;
  try
    QQ := OPENSQL ('SELECT YTC_TABLELIBRETIERS1,YTC_TABLELIBRETIERS2,YTC_TABLELIBRETIERS3,YTC_TABLELIBRETIERS4,YTC_TABLELIBRETIERS5,YTC_TABLELIBRETIERS6'+
				',YTC_TABLELIBRETIERS7,YTC_TABLELIBRETIERS8,YTC_TABLELIBRETIERS9,YTC_TABLELIBRETIERSA'+
				'  FROM TIERSCOMPL WHERE YTC_TIERS="' +
      CodeTiers + '"', False) ;
    if  QQ.EOF then exit;
			//on traite les tables libres
    if TobValTiers=Nil then exit;
    if not QQ.EOF then
    begin
    for ii := 0 to 9 do
       AfAffectVal('C_LIBRECONTACT' + IntToStr(ii + 1),QQ);
    AfAffectVal('C_LIBRECONTACTA',QQ);
    end;
  finally
    Ferme (QQ) ;
  end;

end;
procedure TOM_Contact.AFAffectVal(Champ: string;QQ:Tquery) ;
var
  ChampMono: ThValCombobox;
  Liste: string;
  vtob: tob;
begin
		//mcd traite la zone en cours pour voir si il y a correspondance
  ChampMono := THValComboBox(GetControl(Champ));
	Liste :='';
  if ChampMono <> nil then
    Liste := ChampMono.datatype ;
  vTob := TobValTiers.FIndFirst(['YLT_DESTINATIONTAB'], [Liste], false);
  if vTob <> nil then  SetField(Champ,QQ.FIndField(Vtob.GetValue('NomChamp')).AsString);
 Liste:='';
end;

procedure TOM_CONTACT.ChargeTobValTiers ;
var StSql,Champ1, CHamp2,Champ3 : string;
  ii : integer;
	TobDet, TobBis, Vtob,TobNouv:TOb;
begin
		//fct qui regarde dans les tablettes identiques, si il y a des corresdpondances
		//tiers et Contact.  Dans ce cas, lors de la création d'une affaire, reportera info du tiers
		//dans le contact

     //seulement 2 champs, on peut tout charger
 Stsql := 'SELECT * FROM YLIENTABLETTE WHERE YLT_ORIGINETAB like "GCLIBRETIE%"';
 TobValTiers:=TOB.Create('',Nil,-1) ;
 TobValTiers.LoadDetailFromSQL (stSql);
 for ii:=0 to  TobValTiers.detail.count-1 do
  begin
  TobDet := TobValTiers.detail[ii];
  StSql := COpy (TobDet.Getvalue('YLT_ORIGINETAB'),13,1); //prend N° à la fin de GcLibreTiers* ...
  TobDet.addChampSupValeur ('NOMCHAMP','YTC_TABLELIBRETIERS'+StSql);
	end;
		//il faut maintenant regarder si on des lien table libre ressource sur tiers et affaire
 if (Vh_GC.AfTobAffectChmp <> nil) then // PL le 19/06/07 : inutile d'aller plus loin dans le cas contraire
 begin
  StSql := 'select ylt_originetab from ylientablette where ylt_originetab like "aftlibreres%" group by ylt_originetab having (count(ylt_originetab)>1)';
  TobBis:=TOB.Create('',Nil,-1) ;
  TobBis.LoadDetailFromSQL (stSql);
		//on boucle sur tous les cahmps libre ressource qui existe plusiers fois dans la table
  if (TobBis.detail.count <>0) then
  begin
    for ii:=0 to  TobBis.detail.count-1 do
     begin
     TobDet := TobBis.detail[ii];
     vtob := Vh_GC.AfTobAffectChmp.FIndFirst(['YLT_ORIGINETAB'], [TobDET.getValue('YLT_ORIGINETAB')], false);
     if vTob <> nil then
       begin
          //on e les trizte que si associé à table libre tiers et affaire
       Champ1:= Vtob.GetValue('YLT_DESTINATIONTAB');
       if (Copy(Champ1,1,12) <> 'GCLIBRETIERS') and (Copy(Champ1,1,10) <> 'YYLIBRECON') then champ1:='';
       champ2:='';
       champ3:='';
       vtob := Vh_GC.AfTobAffectChmp.FIndNext(['YLT_ORIGINETAB'], [TobDET.getValue('YLT_ORIGINETAB')], false);
       While vtob  <>Nil do
          begin
          Champ3:= Vtob.GetValue('YLT_DESTINATIONTAB');
          if (Copy(Champ3,1,12) <> 'GCLIBRETIERS') and (Copy(Champ3,1,10) <> 'YYLIBRECON') then champ3:='';
          if champ3 <>'' then
            begin
            if champ1='' then champ1:=champ3
              else champ2:=champ3;
            end;
          vtob := Vh_GC.AfTobAffectChmp.FIndNext(['YLT_ORIGINETAB'], [TobDET.getValue('YLT_ORIGINETAB')], false);
          end;
        if (champ2 <>'') and (champ1 <>'') then
          begin  //il existe 2enrgt tiers et affaire, on ajoute dans la tob
          TobNouv:= tob.create ('YLIENTABLETTE',TobValTiers,-1);
          if (Copy(Champ1,1,12) = 'GCLIBRETIERS') then
            begin
            TobNouv.putvalue ('YLT_ORIGINETAB',champ1);
            TobNouv.putvalue ('YLT_DESTINATIONTAB',champ2);
            StSql := COpy (champ1,13,1); //prend N° à la fin de GcLibreTiers* ...
            Tobnouv.addChampSupValeur ('NOMCHAMP','YTC_TABLELIBRETIERS'+StSql);
            end
          else
            begin
            TobNouv.putvalue ('YLT_ORIGINETAB',champ2);
            TobNouv.putvalue ('YLT_DESTINATIONTAB',champ1);
            StSql := COpy (champ2,13,1); //prend N° à la fin de GcLibreTiers* ...
            Tobnouv.addChampSupValeur ('NOMCHAMP','YTC_TABLELIBRETIERS'+StSql);
            end;
          end;
       end;
     end; // for
  end; //if (TobBis.detail.count <>0)
  FreeAndNil (TobBis);
 end;

 If TobValTiers.detail.count=0 then FreeAndNil(TobValTiers);
end;
{$endif}

{$IFDEF GRC}
procedure TOM_CONTACT.BTELBUR_OnClick(Sender: TObject);
begin
{$IFNDEF BUREAU}
 {$IFNDEF GIGI} // $$$ JP 09/08/07: pas de CTI en GI (PCL), car fait dans le Bureau exclusivement
  FonctionCTI ('MAKECALL', GetControlText ('C_FAX'));
 {$ENDIF GIGI}
{$ENDIF BUREAU}
end;

procedure TOM_CONTACT.BTELDOM_OnClick(Sender: TObject) ;
begin
{$IFNDEF BUREAU}
 {$IFNDEF GIGI} // $$$ JP 09/08/07: pas de CTI en GI (PCL), car fait dans le Bureau exclusivement
  FonctionCTI ('MAKECALL', GetControlText ('C_TELEPHONE'));
 {$ENDIF GIGI}
{$ENDIF BUREAU}
end ;

procedure TOM_CONTACT.BTELPORT_OnClick(Sender: TObject) ;
begin
{$IFNDEF BUREAU}
 {$IFNDEF GIGI} // $$$ JP 09/08/07: pas de CTI en GI (PCL), car fait dans le Bureau exclusivement
  FonctionCTI ('MAKECALL', GetControlText ('C_TELEX'));
 {$ENDIF GIGI}
{$ENDIF BUREAU}
end ;

{$ENDIF GRC}

{$IF Defined(CRM) or Defined(GRCLIGHT)}
procedure TOM_CONTACT.BSMS_OnClick(Sender: TObject) ;
var Infos : string;
    TobInfos : Tob;
begin
  if GetField ('C_TELEX') = '' then exit;
  if LoginEsker = False then Exit;
  TobInfos:=Tob.create('LesInfos',Nil , -1);
  Infos := RTLanceFiche_RTSaisieInfosEsker ('RT','RTSAISIESMS','','','TobInfos='+intToStr(longint(TobInfos)));
  if Infos <> '' then
  begin
    TobInfos.AddChampSupValeur ('TEL',CleTelephone (GetField ('C_TELEX')),False);
    TobInfos.AddChampSupValeur ('NOMEMETTEUR',VH_RT.RTNomResponsable,False);
    SendEsker (3,TobInfos,'',Nil);
  end;
  LogoutEsker;
  FreeAndNil (TobInfos);
end ;
{$IFEND}

{$IFDEF CTI}
procedure TOM_CONTACT.SetEtatBoutonCti;
begin
  BPhone.Visible := (VH_RT.ctiAlerte.GetState<>TCtiDisable)
  									and (VH_RT.ctiAlerte.GetState<>TctiDecroche)
                    and (VH_RT.ctiAlerte.GetState<>TctiNothingToDo)   ;
  BAttente.Visible:= (VH_RT.ctiAlerte.GetState<>TCtiDisable)
  										and (VH_RT.ctiAlerte.GetState=TctiDecroche);
  BStop.Visible:= (VH_RT.ctiAlerte.GetState<>TCtiDisable)
  								and (VH_RT.ctiAlerte.GetState=TctiDecroche) ;
end;

procedure TOM_Contact.BAttenteClick(Sender: TObject);
begin
  if BATTENTE.hint = 'Mise en attente' then
  begin
  SetControlProperty('INFOSCTI','Caption','Mise en attente');
  FonctionCTI ('CALLWAIT','');
  BATTENTE.hint := 'Reprise de l''appel';
  end else
  begin
  SetControlProperty('INFOSCTI','Caption','Communication en cours');
  FonctionCTI ('CALLWAIT','');
  BATTENTE.hint := 'Mise en attente';
  end;
end;

procedure TOM_Contact.BPhoneClick(Sender: TObject);
begin
  SetControlProperty('INFOSCTI','Caption','Communication en cours');
  FonctionCTI ('GETCALL','');
  BPHONE.Down := True;
  BPHONE.enabled := false;
  BAttente.Visible:= True;
  BStop.Visible:= True;
end;

procedure TOM_Contact.BStopClick(Sender: TObject);
begin
  SetControlProperty('INFOSCTI','Caption','Communication terminée');
  FonctionCTI ('CALLBYE','');
  BPHONE.Down := false;
  BPHONE.enabled := true;
  BAttente.Visible:= false;
  BStop.Visible:= false;
	SetEtatBoutonCti;
end;
{$ENDIF}

procedure TOM_Contact.ChangeInfoSUp (Sender : Tobject);
begin
  if not (DS.State in [dsInsert, dsEdit]) then
  begin
    DS.edit;
    SetControlEnabled('bPost',true);
    SetControlEnabled('bRechercher',false);
    SetControlEnabled('bFirst',false);
    SetControlEnabled('bPrev',false);
    SetControlEnabled('bNext',false);
    SetControlEnabled('bLast',false);
    SetControlEnabled('bDefaire',true);
    SetControlEnabled('bInsert',false);
    SetControlEnabled('bDelete',false);
    SetControlEnabled('bDupliquer',false);
  end;
end;

procedure TOM_Contact.DeleteContactBTP(TypeContact, Auxiliaire: string; Numero: integer);
var
  SQl : string;
begin
  SQl := 'DELETE FROM BCONTACT WHERE '+
         'BC1_TYPECONTACT="'+ TypeContact +'" AND '+
         'BC1_AUXILIAIRE="'+Auxiliaire+'" AND '+
         'BC1_NUMEROCONTACT='+InttoStr(Numero);
  ExecuteSql (SQL);
end;

procedure TOM_Contact.FindContactBTP(TypeContact, Auxiliaire: string; Numero: integer);
var QQ : TQuery;
    SQl : string;
begin
  TOBCONTACTBTP.InitValeurs(false);
  TOBCONTACTBTP.SetString('BC1_TYPECONTACT',TypeContact);
  TOBCONTACTBTP.SetString('BC1_AUXILIAIRE',Auxiliaire);
  if numero = 0 then exit;

  SQl := 'SELECT * FROM BCONTACT WHERE '+
         'BC1_TYPECONTACT="'+ TypeContact +'" AND '+
         'BC1_AUXILIAIRE="'+Auxiliaire+'" AND '+
         'BC1_NUMEROCONTACT='+InttoStr(Numero);
  QQ := OpenSql (SQL,true,1,'',true);
  if not QQ.eof then
  begin
    TOBCONTACTBTP.SelectDB('',QQ);
  end;
  ferme (QQ);
end;

procedure TOM_Contact.SetEvents (Etat : boolean);
begin
  if Etat then
  begin
    THEdit(GetControl ('BC1_BATIMENT')).OnChange := ChangeInfoSUp;
    THEdit(GetControl ('BC1_ETAGE')).OnChange := ChangeInfoSUp;
    THEdit(GetControl ('BC1_ESCALIER')).OnChange := ChangeInfoSUp;
    THEdit(GetControl ('BC1_PORTE')).OnChange := ChangeInfoSUp;
  end else
  begin
    THEdit(GetControl ('BC1_BATIMENT')).OnChange := nil;
    THEdit(GetControl ('BC1_ETAGE')).OnChange := nil;
    THEdit(GetControl ('BC1_ESCALIER')).OnChange := nil;
    THEdit(GetControl ('BC1_PORTE')).OnChange := nil;
  end;
end;



procedure TOM_Contact.OnCancelRecord;
begin
  inherited;
  SetEvents (false);
  FindContactBTP (GetField('C_TYPECONTACT'),GetField('C_AUXILIAIRE'),GetField('C_NUMEROCONTACT'));
  TOBCONTACTBTP.PutEcran(ecran);
  SetEvents (true);

end;

procedure TOM_Contact.SetButtonEna;
begin
  TtoolbarButton97(GetControl('bInsert')).Visible := CanCreate;
  TtoolbarButton97(GetControl('bDelete')).Visible := CanCreate;
end;

Initialization
RegisterClasses([TOM_Contact]);
RegisterAglProc( 'RTOLKInsertContact', TRUE , 0, AGLRTOLKInsertContact);
{$IFDEF CTI}
RegisterAglProc( 'RTCTIDecrocheAppelContact', TRUE , 0, AGLRTCTIDecrocheAppelContact);
RegisterAglProc( 'RTCTIRaccrocheAppelContact', TRUE , 0, AGLRTCTIRaccrocheAppelContact);
RegisterAglProc( 'RTCTINumeroterContact', TRUE , 0, AGLRTCTINumeroterContact);
RegisterAglProc( 'RTCTIAppelAttenteContact', TRUE , 0, AGLRTCTIAppelAttenteContact);
RegisterAglProc( 'RTCTIRepriseAppelContact', TRUE , 0, AGLRTCTIRepriseAppelContact);
{$ENDIF}
{$IFDEF GRC}
RegisterAglProc( 'RTAppelParamCLContact', True,0,AGLRTAppelParamCLContact) ;
{$ENDIF}
RegisterAglProc('c_AppelIsoFlex', TRUE, 2, c_AppelIsoFlex);
end.
