{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 20/02/2003
Modifié le ... : 26/03/2007
Description .. : Source TOM de la TABLE : SECTION (SECTION)
Suite ........ :
Suite ........ : JP 17/01/06 : FQ 17293 : Remplacement De ZOOMTABLE
Suite ........ : par DataType dans des SetControlProperty
Suite ........ : JP 17/01/06 : Suppression de la fonction Ergo35
Suite ........ : 
Suite ........ : JP 18/01/06 : FQ 17201 : gestion de l'état en eAgl
Suite ........ : 
Suite ........ : SBO 26/03/2007 
Suite ........ : -> DEV 5942 : Gestion des nouvelles zones + refonte de 
Suite ........ : l'interface
Mots clefs ... : TOM;SECTION
*****************************************************************}

Unit CPSECTION_TOM ;

Interface

Uses
    StdCtrls,
    Controls,
    Classes,
    Windows,         // VK_
    LookUp,          // LookupValueExist
{$IFDEF VER150} {D7}
    Variants,
{$ENDIF}
{$IFDEF EAGLCLIENT}
    eFiche,          // TFFiche
    UtileAGL,        // TNavigateBtn
    MaineAGL,        // AGLLanceFiche
{$IFNDEF PGIIMMO}
    CPSTRUCTURE_TOF, // ParamPlanAnal
{$ENDIF}
{$ELSE}
    db,
    edtREtat,
   {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF}
    ,dbCtrls,
    Fiche,         // TFFiche
    FE_Main,       // AGLLanceFiche
{$IFNDEF PGIIMMO}
    // FQ 20198 Structur,      // ParamPlanAnal
{$IFNDEF CMPGIS35}
    CPSTRUCTURE_TOF,
{$ENDIF}
{$ENDIF}

{$ENDIF}
    forms,
    sysutils,
    ComCtrls,
    Graphics,      // clWindow, clBtnFace, clRed
    HCtrls,
    HEnt1,         // IsFieldNull
    HMsgBox,       // JaiLeDroitConcept, THMsgBox
    Buttons,       // TBitBtn
    Menus,         // TPopupMenu
    HCompte,       // ExisteCpte, THCpteEdit
    HTB97,         // TToolBarButton97, TToolWindow97
    HRichOle,      // THRichEditOle
    CRITEDT,       // TCritEdt
{$IFNDEF PGIIMMO}
{$IFNDEF CCADM}
    CodeSect,      //
{$ENDIF}
{$ENDIF}
    UTOM,
    UTob,
    Ent1,          // VH
    AglInit,       // ActionToString
    UTILPGI
{$IFDEF COMPTA}
    ,CUMMENS
{$IFNDEF CMPGIS35}
    ,CPPARAMTRANCHECOMPTE_TOF
{$ENDIF}
{$ENDIF}
{$IFDEF CCSTD}
    ,CPPARAMTRANCHECOMPTE_TOF
{$ENDIF}
{$IFDEF MODENT1}
    , CPTypeCons
{$ENDIF MODENT1}
,uEntCommun
    ;

procedure FicheSection(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
procedure FicheSectionMZS(Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : string);

function  SectionRetrie(CodRupt : String ; Axe : String ; ListeCodes : TStringList) : TSectRetri;
function  EstDansAnalytiq(Stc,Sta : String) : Boolean;
function  EstDansAxe(Stc,Sta : String) : Boolean;
function  EstCorresp(Stc,Sta : String) : Boolean;
function  EstDansVentil(Stc,Sta : String) : Boolean;
function  EstDansGuideAna(Stc,Sta : String) : Boolean;
function  CodeStructure(Caxe,Csec : String) : Boolean ;

Type
    TOM_SECTION = Class (TOM)
    public
        Pages              : TPageControl;
        BMenuZoom          : TToolBarButton97;
        PopZ               : TPopupMenu;

        S_SECTION          : THEdit;

        {$IFNDEF PGIIMMO}
        {$IFNDEF GCGC}
        {JP 14/10/05 : FQ 16863 : Branchement de la balance section / général}
        ACritEdt           : ClassCritEdt;
        {$ENDIF PGIIMMO}
        {$ENDIF GCGC}

        procedure OnNewRecord              ; override ;
        procedure OnDeleteRecord           ; override ;
        procedure OnUpdateRecord           ; override ;
        procedure OnAfterUpdateRecord      ; override ;
        procedure OnLoadRecord             ; override ;
        procedure OnChangeField(F: TField) ; override ;
        procedure OnArgument(S: String)    ; override ;
        procedure OnClose                  ; override ;
        procedure OnCancelRecord           ; override ;

        procedure FormCloseQuery           (Sender: TObject; var CanClose: Boolean);

        procedure OnClickBInsert           (Sender: TObject);
        procedure OnClickBImprimer         (Sender: TObject);

        procedure OnClickBCodSec           (Sender: TObject);
        procedure OnClickBSaiCodsec        (Sender: TObject);
        procedure OnChangeS_AXE             (Sender: TObject);
        procedure OnClickBCopierSect       (Sender: TObject);

        procedure OnEnterBMenuZoom         (Sender: TObject);
        procedure OnClickBCumul            (Sender: TObject);
        procedure OnClickBZecrimvt         (Sender: TObject);
{$IFDEF COMPTA}
        procedure OnClickBBalSecGen        (Sender: TObject);
{$ENDIF}
        procedure OnClickConfidentiel      (Sender: TObject);

        procedure FormKeyDown              (Sender: TObject; var Key: Word;Shift: TShiftState);

        procedure OnEnterS_SECTION         (Sender: TObject);
        procedure OnExitS_SECTION          (Sender: TObject);

        procedure OnMouseDownS_BlocNote    (Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
        procedure OnClickBParamTranche     (Sender : TObject);
    private
        Lequel,LesModif    : String;
        QuelAxe            : String;
        Mode               : TActionFiche;
        LeFb               : TFichierBase;
        LgCode             : Byte;
        IsChantier         : Boolean;
        KelPage            : string;         //fb 19/05/2006 FQ17680

        FTobCumuls         : TOB ;           // Contient les infos de cumuls locales en mode multisoc

        FAGLCloseQuery     : TCloseQueryEvent ;

        function  OkConfidentiel       : Boolean;
        function  VerifiSiExiste       : Boolean;
        function  VerifCoherenceTL     : Boolean;
        procedure SwapttRepart         ;
        procedure EstUnAxeChantier     ;
        // pour les modif en série
        procedure InitModifEnSerie     (StModif : string);
        procedure AffecteLe            (Champ, Valeur : string);
        // MULTISOC
        function  GetChampsMS          ( vStChamp : String ) : double ;   // Retourne le cumul demandé soit issu de la table commune, soit de la table locale en multisoc
        procedure ChargeCumulsMultiSoc ;                                  // Renseigne FTobCumuls avec les valeurs locales en mode multisoc
        procedure AfficheInfos         ( vBoInit : boolean = False ) ;
        procedure OnAfterFormShow      ; //fb 31/05/2006 FQ17680
        procedure AfficheLeSolde       ( vStControl : string ; vDebit : double ; vCredit : double ) ; 

    end;

Implementation

uses
  ParamSoc, {FQ 16569}
  {$IFDEF COMPTA}
{$IFNDEF CMPGIS35}
  CPBALANALGEN_TOF, {JP 14/10/05 : FQ 16863}
{$ENDIF}
  {$ENDIF}
    {$IFDEF eAGLCLIENT}
    MenuOLX
    {$ELSE}
    MenuOLG
    {$ENDIF eAGLCLIENT}
    {$IFDEF COMPTA}
      { FQ 20370 BVE 02.10.07 }
{$IFNDEF CMPGIS35}
      ,CPMULANA_TOF // MultiCritereAna
{$ENDIF}
      { END FQ 20370 }
    {$ENDIF}
    {$IFDEF MODENT1}
    , CPProcMetier
    , CPProcGen
    {$ENDIF MODENT1}
    ,UFonctionsCBP;    


{// OLD : Anciens messages du THMessageBox
    MsgBox.Mess.Add('0;Sections analytiques;Voulez-vous enregistrer les modifications ?;Q;YNC;Y;C;');
    MsgBox.Mess.Add('1;Sections analytiques;Confirmez-vous la suppression de l''enregistrement ?;Q;YNC;N;C;');
    MsgBox.Mess.Add('2;Sections analytiques;Vous devez renseigner un code.;W;O;O;O;');
    MsgBox.Mess.Add('3;Sections analytiques;Vous devez renseigner un libellé.;W;O;O;O;');
    MsgBox.Mess.Add('4;Sections analytiques;Le code que vous avez saisi existe déjà. Vous devez le modifier.;W;O;O;O;');
    MsgBox.Mess.Add('5;Sections analytiques;Vous ne pouvez pas supprimer cette section : elle possède des écritures analytiques.;W;O;O;O;');
    MsgBox.Mess.Add('6;Sections analytiques;Vous ne pouvez pas supprimer cette section : elle est déclarée en tant que section d''attente pour cet axe.;W;O;O;O;');
    MsgBox.Mess.Add('7;Sections analytiques;Vous ne pouvez pas supprimer cette section : elle est déclarée en tant que section de correspondance.;W;O;O;O;');
    MsgBox.Mess.Add('8;Sections analytiques;Vous ne pouvez pas supprimer cette section : elle est utilisée dans une ventilation type ou une ventilation par défaut d''un compte.;W;O;O;O;');
    MsgBox.Mess.Add('9;Sections analytiques;Vous ne pouvez pas supprimer cette section : elle est utilisée comme modèle pour la création.;W;O;O;O;');
    MsgBox.Mess.Add('L''enregistrement est inaccessible');
    MsgBox.Mess.Add('11;Sections analytiques;Vous ne pouvez pas supprimer cette section : elle est référencée dans un guide.;W;O;O;O;');
    MsgBox.Mess.Add('L''accès à cette section est interdit.');
    MsgBox.Mess.Add('13;Sections analytiques;La section de correspondance n°1 n''existe pas. Vous devez le modifier.;W;O;O;O;');
    MsgBox.Mess.Add('14;Sections analytiques;La section de correspondance n°2 n''existe pas. Vous devez le modifier.;W;O;O;O;');
    MsgBox.Mess.Add('15;Sections analytiques;ATTENTION. Il y a des codes de tables libres qui n''existent pas. Désirez vous continuer ?;Q;YNC;N;C;');
    MsgBox.Mess.Add('16;Sections analytiques;Vous devez renseigner une fin de fourchette de comptes.;W;O;O;O;');
    MsgBox.Mess.Add('17;Sections analytiques;Vous devez renseigner un début de fourchette de comptes.;W;O;O;O;');
}
{*****************************************************************}
function  EstDansAnalytiq(Stc,Sta : String) : Boolean;
begin
    result := ExisteSQL('Select Y_SECTION From ANALYTIQ Where Y_SECTION="'+Stc+'" and Y_AXE="'+Sta+'"');
end;

{*****************************************************************}
function  EstDansAxe(Stc,Sta : String) : Boolean;
begin
    result := (Stc = VH^.Cpta[AxeToFb(Sta)].Attente);
end;

{*****************************************************************}
function  EstCorresp(Stc,Sta : String) : Boolean;
begin
    result := ExisteSQL('Select S_CORRESP1,S_CORRESP2 From SECTION where (S_CORRESP1="'+Stc+'" OR S_CORRESP2="'+Stc+'") and (S_AXE="'+Sta+'")');
end;

{*****************************************************************}
function  EstDansVentil(Stc,Sta : String) : Boolean;
begin
    result := Presence('VENTIL','V_SECTION',Stc);
end;

{*****************************************************************}
function EstDansGuideAna(Stc,Sta : String) : Boolean;
begin
    result := ExisteSQL('Select AG_SECTION From ANAGUI Where AG_SECTION="'+Stc+'" and AG_AXE="'+Sta+'"');
end;

{*****************************************************************}
function SectionRetrie(CodRupt : String ; Axe : String ; ListeCodes : TStringList) : TSectRetri;
var
    Q : TQuery;
    Deb,Lon : TabByte;
    St,StTemp,Ru : String;
    i,Lg : Byte;
    Trie,Execute : Boolean;
    Tous : Boolean;
begin
    result := srOk;

    if (not VH^.Cpta[AxeToFb(Axe)].Structure) then
    begin
        result := srNonStruct;
        exit;
    end;

    Ru := 'RU'+Axe[2];
    Q := OpenSQL('SELECT CC_LIBRE FROM CHOIXCOD WHERE CC_TYPE="'+Ru+'" and CC_CODE="'+CodRupt+'"',true,-1,'',true);
{$IFDEF EAGLCLIENT}
    if (Q.Fields[0].AsString = '') or (Q.Eof) then
{$ELSE}
    if (Q.Fields[0].AsString = '') or (IsFieldNull(Q,'CC_LIBRE')) or (Q.Eof) then
{$ENDIF}
    begin
        result := srPasEnchainement;
        exit;
    end
    else StTemp := Q.Fields[0].AsString;

    St := '';

    Ferme(Q);

    FillChar(Deb,SizeOF(Deb),#0);
    FillChar(Lon,SizeOF(Lon),#0);

    i := 1;

    While (StTemp <> '') do
    begin
        St := ReadTokenSt(StTemp);
        Q := OpenSQL('SELECT SS_DEBUT,SS_LONGUEUR FROM STRUCRSE WHERE SS_AXE="'+Axe+'" and SS_SOUSSECTION="'+St+'"',true,-1,'',true);
        Deb[i] := Q.Fields[0].AsInteger;
        Lon[i] := Q.Fields[1].AsInteger;
        Ferme(Q);
        Inc(i);
    end;

    Lg := VH^.Cpta[AxeToFb(Axe)].Lg;
    Q := OpenSQL('SELECT S_SECTION,S_SECTIONTRIE FROM SECTION WHERE S_AXE="'+Axe+'"',true,-1,'',true);
    Trie := false;
    Execute := false;
    Tous := PresenceComplexe('SECTION',['S_AXE','S_SECTIONTRIE'],['=','='],[Axe,''],['S','S']);

    While ((not Q.EOF) and ((not Trie) or Tous)) do
    begin
        St := '';
        for i := 1 to Lg do
        begin
            StTemp := Copy(Q.Fields[0].AsString, Deb[i], Lon[i]);
            if (StTemp = '') then Break;
            St := St+StTemp;
        end;
        if Tous and (Q.Fields[1].AsString='') Then Execute := false;
        if not Execute then begin
{$IFDEF EAGLCLIENT}
        if (Q.Fields[1].AsString='') then
{$ELSE}
        if (Q.Fields[1].AsString='') or IsFieldNull(Q,'S_SECTIONTRIE') then
{$ENDIF}
        begin
            Trie := false;
            Execute := true;
        end
        else
            if Q.Fields[0].AsString<>VH^.Cpta[AxeToFb(Axe)].Attente then
            begin
                Execute := true;
                if (Q.Fields[1].AsString <> St) then Trie := false
                else Trie := true;
            end;
        end;

        if (not Trie) then ExecuteSQL('UPDATE SECTION SET S_SECTIONTRIE="'+St+'" WHERE S_SECTION="'+Q.Fields[0].AsString+'" and S_AXE="'+Axe+'"');
        Q.Next;
    end;
    Ferme(Q);
end;

{*****************************************************************}
function  CodeStructure(Caxe,Csec : String) : Boolean ;
var
    D1,L1     : TabByte;
    C1        : TabSt3;
    Li        : TabSt35;
    i,Lg      : Byte;
    St,Lib    : String;
    Trouver   : Boolean;
    QLoc      : TQuery;
begin
    result := false;

    if ((Caxe='') or (Csec='')) then exit;
    if (not VH^.Cpta[AxeToFb(CAxe)].Structure) then exit;

    ChargeTablo(Caxe,D1,L1,C1,Li);
    Lg := VH^.Cpta[AxeToFb(Caxe)].Lg;
    Lib := Caxe+' '+Csec+' '+V_PGI.UserName+' '+DateToStr(Date);
    if (Length(Lib) > 35) then Lib := Copy(Lib,1,35);

    for i := 1 to Lg do
    begin
        St := Copy(CSec,D1[i],L1[i]);
        Trouver := false;
        if (St = '') then Break ;
        QLoc := OpenSQL('Select PS_CODE from SSSTRUCR where PS_AXE="'+Caxe+'" and '+ 'PS_SOUSSECTION="'+C1[i]+'"',true,-1,'',true);
        While (not QLoc.EOF) and (not Trouver) do
        begin
            if (St = QLoc.Fields[0].AsString) then Trouver := true
            else QLoc.Next;
        end;
        if (not Trouver) then
        begin
            ExecuteSQL('Insert Into SSSTRUCR (PS_AXE,PS_SOUSSECTION,PS_CODE,PS_LIBELLE,PS_SOCIETE) '+'Values("'+Caxe+'","'+C1[i]+'","'+St+'","'+Lib+'","'+V_PGI.CodeSociete+'")');
            result := true;
        end;
        Ferme(QLoc);
    end;
end;

{*****************************************************************}
procedure FicheSectionMZS(Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : string);
var lStAction : String ;
begin
    // Lancement de la fiche selon le contexte de creation modification ou consultation
    lStAction := ActionToString( Comment ) ;

    Case Comment of
        // creation
        taCreat,taCreatEnSerie,taCreatOne :
            begin
                if (not ExJaiLeDroitConcept(TConcept(ccSecCreat),true)) then exit;
                if lStAction='' then lStAction := 'ACTION=CREATION' ;
                AGLLanceFiche('CP','CPSECTION','','',lStAction+';'+Compte+';'+Axe+';'+LesModif);
            end;
        // modification
        taModif,taModifEnSerie :
            begin
                if (not ExJaiLeDroitConcept(TConcept(ccSecModif),true)) then exit;
                if _Blocage(['nrCloture'],True,'nrSaisieCreat') then Exit ;
                if lStAction='' then lStAction := 'ACTION=MODIFICATION' ;
{b fb 19/05/2006 FQ17680}
                AGLLanceFiche('CP','CPSECTION','','',lStAction + ';'+Compte+';'+Axe+';'+LesModif+';'+IntToStr(QuellePage));
{e fb 19/05/2006 FQ17680}
            end;
        // Consultation
        taConsult : AGLLanceFiche('CP','CPSECTION','','','ACTION=CONSULTATION;'+Compte+';'+Axe+';'+LesModif);
    end;
    if (Comment in [taModif,taModifEnSerie]) then _Bloqueur('nrBatch',False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 20/02/2003
Modifié le ... :   /  /
Description .. : point d'entrée pour la modification en série
Mots clefs ... :
*****************************************************************}
procedure FicheSection(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
begin
{b fb 19/05/2006 FQ17680}
  //  FicheSectionMZS(Axe,Compte,Comment,0,'');
  FicheSectionMZS(Axe,Compte,Comment,QuellePage,'');
{e fb 19/05/2006 FQ17680}
end;

{*****************************************************************}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 14/09/2006
Modifié le ... :   /  /    
Description .. : - LG - 15/09/2006 - FB 18798 - La valeur de QuelAxe 
Suite ........ : n'était pas modifie qd on forcai le changemetn d'axe
Mots clefs ... : 
*****************************************************************}
procedure TOM_SECTION.OnNewRecord;
begin
  Inherited;

  SetField('S_CONFIDENTIEL','0');
  SetField('S_SOLDEPROGRESSIF','X');
  SetField('S_SENS','M');

  // Nouvelles zones : Amélioration de l'analytique // SBO
  SetField('S_INDIRECTE',     '-' ) ;
  SetField('S_TRANCHEGENEDE', '' ) ;
  SetField('S_TRANCHEGENEA',  '' ) ;
  SetField('S_UO',            0 ) ;
  SetField('S_UOLIBELLE',     '' ) ;

  //YMO 29/03/2006 Initialisation
  SetField('S_FINCHANTIER',idate2099);
  if ((Mode in [taCreatOne,taCreat]) and (LeQuel <> '')) then
    SetField('S_SECTION',LeQuel);

  // MAJ n° Axe
  if (Mode in [taModif,taCreatOne,taCreat,taCreatEnSerie]) then
    begin
    if (QuelAxe <> '') then
      begin
      if (QuelAxe <> GetParamSocSecur('SO_CPPCLAXETVA', ''))
        then SetField('S_AXE',QuelAxe)
      else if GetParamSocSecur('SO_CPPCLAXETVA', '')='A1'
        then SetField('S_AXE','A2')
        else SetField('S_AXE','A1')
      end
    else if GetParamSocSecur('SO_CPPCLAXETVA', '')='A1'
      then SetField('S_AXE','A2')
      else SetField('S_AXE','A1') ;
    end;

  if Mode = taModif then
    Mode    := taCreat ;
  QuelAxe := GetField('S_AXE') ;
  LgCode  := VH^.Cpta[ AxeToFb( QuelAxe ) ].Lg ;
  SetControlProperty('S_SECTION', 'MAXLENGTH', LgCode );

  SetControlEnabled( 'S_AXE',     true  );
  SetControlEnabled( 'S_SECTION', true  );
  SetControlEnabled( 'BMENUZOOM', false );

  Pages.ActivePage := TTabSheet(GetControl('PGENERAL',true));
  AfficheInfos( True ) ;

  if ((Mode in [taCreatOne,taCreat]) and (LeQuel <> '')) // pour mise en place du focus si appel avec code pré-remplit SBO 30/05/2007
    then SetFocusControl('S_SECTION')
    else SetFocusControl('S_AXE') ;

end;

{*****************************************************************}
procedure TOM_SECTION.OnDeleteRecord ;
begin
    LastError := 1;
    if (PgiAskCancel( TraduireMemoire('Confirmez-vous la suppression de l''enregistrement ?'), Ecran.Caption) <> mrYes) then
      exit;
    LastError := 0;

    Inherited;
end;

{*****************************************************************}
procedure TOM_SECTION.OnUpdateRecord;
begin
    Inherited;

  if (GetControlText('S_AXE') = '') then // Vous devez renseigner un axe.
    begin
        Pages.ActivePage := TTabSheet(GetControl('PGENERAL',true));
        SetFocusControl('S_AXE');
        PgiInfo( TraduireMemoire('Vous devez renseigner un axe.'), Ecran.Caption);
        LastError := 1;
        exit;
    end;

  if (GetControlText('S_SECTION') = '') then // Vous devez renseigner un code.
    begin
        Pages.ActivePage := TTabSheet(GetControl('PGENERAL',true));
        SetFocusControl('S_SECTION');
        PgiInfo( TraduireMemoire('Vous devez renseigner un code.'), Ecran.Caption);
        LastError := 1;
        exit;
    end;
    if (GetControlText('S_LIBELLE') = '') then // Vous devez renseigner un libellé.
    begin
        Pages.ActivePage := TTabSheet(GetControl('PGENERAL',true));
        SetFocusControl('S_LIBELLE');
        PgiInfo( TraduireMemoire('Vous devez renseigner un libellé.'), Ecran.Caption);
        LastError := 1;
        exit;
    end;

    // GCO - 30/01/2006 - FQ 17216
    if GetParamSocSecur('SO_CPPCLSAISIETVA', False ) then
    begin
      if GetControlText('S_AXE') = GetParamSocSecur('SO_CPPCLAXETVA', '') then
      begin
        SetFocusControl('S_AXE');
        // GCO - 27/11/2006 - FQ 17785 (modifier à la place de créer...)
        PgiInfo( TraduireMemoire('Impossible de modifier une section de l''axe TVA.'), Ecran.Caption);
        LastError := -1;
        Exit;
      end;
    end;

    // === Ajout tests nouvelles zones :  DEV 5942

    // Test saisi de la tranche complete
(*    if (GetControlText('S_TRANCHEGENEDE') <> '') then
      begin
      if (GetControlText('S_TRANCHEGENEA')='') then
        begin
        Pages.ActivePage := TTabSheet(GetControl('PCOMPLEMENT',true));
        SetFocusControl('S_TRANCHEGENEA');
        PGIInfo( TraduireMemoire('La fourchette de comptes est incomplète.'), Ecran.Caption) ;
        LastError := 1;
        exit;
        end;
      end
    else
      if (GetControlText('S_TRANCHEGENEA')<>'') then
        begin
        Pages.ActivePage := TTabSheet(GetControl('PCOMPLEMENT',true));
        SetFocusControl('S_TRANCHEGENEDE');
        PGIInfo( TraduireMemoire('La fourchette de comptes est incomplète.'), Ecran.Caption) ;
        LastError := 1;
        exit;
        end;
*)
    if (not VerifCoherenceTL) then exit;

    Case (VerifCorrespondance(2+StrToInt(QuelAxe[2]),GetControlText('S_CORRESP1'),GetControlText('S_CORRESP2'))) of
        0 :;
        1 : begin // La section de correspondance n°1 n''existe pas. Vous devez le modifier.
            PGIInfo( TraduireMemoire('La section de correspondance n°1 n''existe pas. Vous devez le modifier.'), Ecran.Caption );
            LastError := 1; exit;
            end;
        2 : begin// La section de correspondance n°2 n''existe pas. Vous devez le modifier.
            PGIInfo( TraduireMemoire('La section de correspondance n°2 n''existe pas. Vous devez le modifier.'), Ecran.Caption );
            LastError := 1;
            exit;
            end;
    end;

    if (DS.state in [dsInsert]) then
    begin
        if (VerifiSiExiste) then
          begin
          LastError := 1;
          exit;
          end ;
        if (not IsChantier) then
        begin
            SetControlText('S_MAITREOEUVRE','');
            SetControlText('S_CHANTIER','');
        end;
    end;
    CodeStructure(GetControlText('S_AXE'),GetControlText('S_SECTION'));

  // ============================
  // ==== GESTION DU MULTISOC ===
  // ============================
  // Ajout de l'enregistrement local de la table cumuls pour gestion des totaux en multisoc
// En attente car pour le moment n'enregistre que sur la base courante...
//  if  DS.State in [DsInsert] then
//    CInsertCumulsMS( fbSect, GetField('S_SECTION'), GetField('S_AXE') ) ;

end;

{*****************************************************************}
procedure TOM_SECTION.OnAfterUpdateRecord;
begin
  Inherited;
  // GCO - 30/10/2007 - FQ 21755
{$IFDEF EAGLCLIENT}
  AvertirCacheServer('AXE');
{$ENDIF}
  AvertirTable('TTAXE');
end;

{*****************************************************************}
procedure TOM_SECTION.OnLoadRecord;
begin
  Inherited;
  
  { FQ 21210 BVE 01.08.07 }
  SetControlProperty('S_TRANCHEGENEDE','Plus','RB_CLASSERUB="TRA" AND RB_AXE="'+GetControlText('S_AXE')+'"');
  { END FQ 21210 }
  
  TCheckBox(GetControl('SCONFIDENTIEL',true)).Checked := (GetField('S_CONFIDENTIEL') = 1);

  // chargement des infos de cumuls MULTISOC
  ChargeCumulsMultiSoc ;
  AfficheInfos ;

  EstUnAxeChantier;
  SwapttRepart;

  if (ds.state = dsInsert) then begin
    SetControlEnabled('BSAICODSEC',True);
    SetControlEnabled('BCODSEC',True);
    end
  else begin
    SetControlEnabled('BSAICODSEC',false);
    SetControlEnabled('BCODSEC',false);
  end;
  SetControlEnabled('BMENUZOOM',true);

  if (((DS.State = dsInsert) and (GetControlText('S_SECTION') <> '')) or
      (DS.State in [dsEdit,dsBrowse]) or (Mode = taModif)) then
    SetControlEnabled('S_AXE',false);

  if Mode = taModifEnSerie then
    InitModifEnSerie(LesModif) ;
    
end;

{*****************************************************************}
procedure TOM_SECTION.OnChangeField(F: TField);
begin
  Inherited;

  SetControlEnabled('BMENUZOOM',(not(DS.State in [dsInsert])));
  SetControlEnabled('BIMPRIMER',(not(DS.State in [dsInsert])));

  if (GetControlText('S_ABREGE') = '') then
    SetControlText('S_ABREGE', Copy( GetControlText('S_LIBELLE'), 1, 17) );

  // GESTION NATURE GENE
  if F.FieldName='S_AFFAIREENCOURS' then
    begin
    SetControlEnabled('S_DEBCHANTIER', GetField('S_AFFAIREENCOURS')='X' );
    SetControlEnabled('S_FINCHANTIER', GetField('S_AFFAIREENCOURS')='X' );
    end
  else if F.FieldName = 'S_AXE' then
    begin
    if (Mode = taConsult) then
       SetControlProperty('S_SECTION', 'MAXLENGTH', 17 )
    else
      begin
      QuelAxe := GetField('S_AXE');
      LgCode  := VH^.Cpta[ AxeToFb( QuelAxe ) ].Lg;
      SetControlProperty('S_SECTION', 'MAXLENGTH', LgCode );
      end ;
    SwapttRepart;
    if (DS.State = dsInsert) then
      EstUnAxeChantier;
    end ;

end;

{*****************************************************************}
procedure TOM_SECTION.OnArgument(S: String);
var lStAxeTVA : string;
begin
    Inherited;
{ FQ 19463 BVE 10.04.07}
 if ctxPCL in V_PGI.PGIContexte then
    TGroupBox(GetControl('HGBCONTRAT')).Visible := false;
{ END FQ 19463}

    // Création des contrôles
    Ecran.HelpContext := 7175000;

    // Taille maxi des comptes généraux
    //SetControlProperty('S_TRANCHEGENEDE', 'MAXLENGTH', VH^.Cpta[ fbGene ].Lg );
    //SetControlProperty('S_TRANCHEGENEA',  'MAXLENGTH', VH^.Cpta[ fbGene ].Lg );

    // Evénements des contrôles
    TToolBarButton97(GetControl('BCOPIERSECT',true)).OnClick := OnClickBCopierSect;
    TToolBarButton97(GetControl('BIMPRIMER',true)).OnClick := OnClickBImprimer;
    THValComboBox(GetControl('S_AXE',true)).OnChange := OnChangeS_AXE;
    TToolBarButton97(GetControl('BCODSEC',true)).OnClick := OnClickBCodSec;
    TToolBarButton97(GetControl('BSAICODSEC',true)).OnClick := OnClickBSaiCodsec;
    TToolBarButton97(GetControl('BINSERT',true)).OnClick := OnClickBInsert;
    if (GetControl('BPARAMTRANCHE',False) <> nil) then
    begin
      TToolBarButton97(GetControl('BPARAMTRANCHE',true)).OnClick := OnClickBParamTranche;
      SetControlVisible('BPARAMTRANCHE',False);
      {$IFDEF COMPTA}
      SetControlVisible('BPARAMTRANCHE',True);
      {$ENDIF}
      {$IFDEF CCSTD}
      SetControlVisible('BPARAMTRANCHE',True);
      {$ENDIF}
    end;
    BMenuZoom := TToolBarButton97(GetControl('BMENUZOOM',true)); BMenuZoom.OnMouseEnter := OnEnterBMenuZoom;
    {JP 02/08/05 : FQ 15232 : Il y avait une inversion du branchement des évènements}
    TToolBarButton97(GetControl('BCUMP',true)).OnClick := OnClickBCumul;

    {$IFDEF EAGL}
    { FQ 20370 BVE 02.10.07 }
    {JP 14/10/05 : FQ 16862 : Le Mul des mouvements analytiques (Zoomana.PAS) n'étant pas
                   migré en eAgl, je désactive le menu}
    // TToolBarButton97(GetControl('BZECRIMVT')).Tag := 0;
    { END FQ 20370 }
    SetControlVisible('BIMPRIMER',true); {JP 18/01/06 : FQ 17201}
    {$ENDIF EAGL}

    {$IFNDEF PGIIMMO}
    {$IFNDEF GCGC}
    {JP 14/10/05 : FQ 16863 : Branchement de la balance section / général}
    ACritEdt := ClassCritEdt.Create;
    {$ENDIF PGIIMMO}
    {$ENDIF GCGC}

    TToolBarButton97(GetControl('BZECRIMVT',true)).OnClick := OnClickBZecrimvt;
{$IFDEF COMPTA}
    TToolBarButton97(GetControl('BBALSECGEN',true)).OnClick := OnClickBBalSecGen;
{$ELSE}
    SetControlEnabled ('BBALSECGEN',False);
{$ENDIF}

    TCheckBox(GetControl('SCONFIDENTIEL',true)).OnClick := OnClickConfidentiel;

    Ecran.OnKeyDown := FormKeyDown;

    Pages := TPageControl(GetControl('PAGES',true));
    PopZ := TPopUpMenu(GetControl('POPZ',true));

    S_SECTION := THEdit(GetControl('S_SECTION',true));
    S_SECTION.OnEnter := OnEnterS_SECTION;
    S_SECTION.OnExit := OnExitS_SECTION;

    // Récupère les arguments

    // ACTION=
    ReadTokenSt(S);
    if TFFiche(Ecran).TypeAction = taCreatEnSerie then
      TFFiche(Ecran).MonoFiche := False ;
    Mode := TFFiche(Ecran).TypeAction ;  // Mode

    // Compte
    Lequel := ReadTokenSt(S);

    // Axe
    QuelAxe := ReadTokenSt(S);

    // LesModif
    LesModif := ReadTokenSt(S);

    //Quelle Page //fb 19/05/2006 FQ17680
    KelPage := ReadTokenSt(S);

    TFFiche(Ecran).OnAfterFormShow := Self.OnAfterFormShow; //fb 31/05/2006 FQ17680
    if not (Mode in [taCreat,taCreatEnSerie,taCreatOne]) then
      TFFiche(Ecran).FLequel := QuelAxe+';'+Lequel; // SBO 28/05/2006 provoque un dyfsonctionnement sur l'initialisatin de l'axe et du code!!!!

    LibellesTableLibre(TTabSheet(GetControl('ZL',true)),'TS_TABLE','S_TABLE','S');

    MakeZoomOLE(Ecran.Handle) ;
    { FQ 20370 BVE 22.05.07 }
    // MakeZoomOLE passe la propriété visible de l'onglet INVISIBLES à false ce qui desactive le menu popup
    TTabSheet(GetControl('INVISIBLES')).Visible := true;
    { END FQ 20370 }
//    SetControlText('TTS_BUDSECT','');

    AfficheInfos(True) ;

    if (not OkConfidentiel) then exit;

    Case Mode of
        taModif : if (QuelAxe <> '') then
                    SetControlEnabled('S_AXE',false);
    end;

    if (mode <> taConsult) Then SetControlVisible('BCOPIERSECT',true);

{$IFDEF COMPTA}
  (*
  CA - 12/05/2004 - Pour éviter de tirer toute la comptabilité, le menu
  d'accès au cumuls des comptes mensuels est rendu invisible lorsque la
  fiche section est appelé depuis un module non complié avec la directive COMPTA
  *)
{$ELSE}
  SetControlProperty( 'BCUMP','Tag',0 );  // Le tag est différent de -(POPUP) donc le menu ne s'affiche pas.
{$ENDIF}

{$IFDEF EAGLCLIENT}
  THRichEditOle(GetControl('S_BLOCNOTE', True)).OnMouseDown := OnMouseDownS_BlocNote;
{$ELSE}
  THDBRichEditOle(GetControl('S_BLOCNOTE', True)).OnMouseDown := OnMouseDownS_BlocNote;
{$ENDIF}
  // Création section sur axe TVA impossible
  if ((GetParamSocSecur('SO_CPPCLSAISIETVA', False)) and
        (Mode in [taCreatOne,taCreat,taCreatEnSerie])) then
  begin
    lStAxeTva := GetParamSocSecur('SO_CPPCLAXETVA', '');
    if lStAxeTva = '' then
    begin // ParamSoc pas renseigné
      lStAxeTva := GetColonneSQL('AXE', 'X_AXE', 'X_LIBELLE = "TVA"');
      SetParamSoc('SO_CPPCLAXETVA', lStAxeTva);
    end;
    SetControlProperty('S_AXE','Plus',' AND X_AXE<>"'+GetParamSocSecur('SO_CPPCLAXETVA', '')+'"');
  end;

  // Empêcher le message sur fermeture dans modif en création // SBO 19/04/2007
  FAGLCloseQuery := TFFiche(Ecran).OnCloseQuery ;
  TFFiche(Ecran).OnCloseQuery := FormCloseQuery ;

end;

{*****************************************************************}
procedure TOM_SECTION.OnClose;
begin
  {$IFNDEF PGIIMMO}
  {$IFNDEF GCGC}
  {JP 14/10/05 : FQ 16863 : Branchement de la balance section / général}
  if Assigned(ACritEdt) then FreeAndNil(ACritEdt);
  {$ENDIF PGIIMMO}
  {$ENDIF GCGC}

  Inherited;

end;

{*****************************************************************}
procedure TOM_SECTION.OnCancelRecord;
begin
    Inherited;
end;

{*****************************************************************}
procedure TOM_SECTION.OnClickBCodSec(Sender: TObject);
{$IFNDEF CCADM}
{$IFNDEF PGIIMMO}
var
    Cod,Lib : String;
    i : Byte;
{$ENDIF}
{$ENDIF}
begin
{$IFNDEF CCADM}
{$IFNDEF PGIIMMO}
    Cod := '';
    Lib := '';

    FicheCodeSect(GetControlText('S_AXE'),Cod,Lib);

    if (Cod <> '') then
    begin
        // suppression de ce qu'il y a avant les ';'
        While (Length(Lib) > 35) do
        begin
            for i := 1 to Length(Lib) do if (Lib[i] = ';') then Delete(Lib,i-1,1);
        end;
        // remplacement des ';' par des ' '
        While (Pos(';',Lib) > 0) do Lib[Pos(';',Lib)] := ' ';

        //SG6 14.03.05 Remplacement des SetControlText en SetField ; FQ 15492
        SetField('S_SECTION', Cod);
        SetField('S_LIBELLE', Lib);
//        SetControlText('S_SECTION',Cod);
//        SetControlText('S_LIBELLE',Lib);
    end;

    if (S_SECTION.CanFocus) then S_SECTION.SetFocus;
    S_SECTION.SelStart := Length(S_SECTION.Text);
    S_SECTION.SelLength := 0;
{$ENDIF}
{$ENDIF}
end;

{*****************************************************************}
procedure TOM_SECTION.OnClickBCopierSect(Sender: TObject);
var
    tz : tZoomTable;
    QC : TQuery ;
    Cache : THCpteEdit;
    i : integer;

begin
    tz := AxeToTz(THValComboBox(GetControl('S_AXE',true)).Value);
    Cache := THCpteEdit.Create(Ecran);
    Cache.Visible := false;
    Cache.Parent := Ecran;
    Cache.ZoomTable := tz;
    Cache.okLocate := true;
    Cache.Text := Copy(GetControlText('S_SECTION'),1,4) ;

    if (GChercheCompte(Cache,FicheSection)) then
    begin
        if (mode = taModif) and (DS.State = dsBrowse) then DS.Edit;
        QC := OpenSQL('SELECT * FROM SECTION WHERE S_SECTION="'+Cache.Text+'"',true,-1,'',true);
        if (not QC.Eof) then
        begin
            // YMO 16/05/20006 FQ17889 Remplacement des SetControlText par SetField
            // YMO 03/2006 Ajout de champs à dupliquer
            SetField('S_CORRESP1',QC.FindField('S_CORRESP1').AsString);
            SetField('S_CORRESP2',QC.FindField('S_CORRESP2').AsString);
            // Options d'impression du grand livre
            SetField('S_SOLDEPROGRESSIF',QC.FindField('S_SOLDEPROGRESSIF').AsString);
            SetField('S_TOTAUXMENSUELS',QC.FindField('S_TOTAUXMENSUELS').AsString);
            SetField('S_SAUTPAGE',QC.FindField('S_SAUTPAGE').AsString);

            SetField('S_CLEREPARTITION',QC.FindField('S_CLEREPARTITION').AsString);
            TCheckBox(GetControl('SCONFIDENTIEL',true)).Checked := (QC.FindField('S_CONFIDENTIEL').AsInteger = 1);
            SetField('S_BLOCNOTE',QC.FindField('S_BLOCNOTE').AsString);

            For i:=0 to 9 do
              SetField('S_TABLE'+inttostr(i),QC.FindField('S_TABLE'+inttostr(i)).AsString);

           SetField('S_AFFAIREENCOURS',QC.FindField('S_AFFAIREENCOURS').AsString);
           SetField('S_DEBCHANTIER',QC.FindField('S_DEBCHANTIER').AsString);
           SetField('S_FINCHANTIER',QC.FindField('S_FINCHANTIER').AsString);
        end;
        Ferme(QC) ;
    end;

end;

{*****************************************************************}
procedure TOM_SECTION.OnClickBInsert(Sender: TObject);
begin
    if (ExJaiLeDroitConcept(TConcept(ccSecCreat),true)) then
      begin
      Lequel := '' ;
      TFFiche(Ecran).FLequel := '' ;
      TFFiche(Ecran).BInsertClick(Sender);
      end ;
end;

procedure TOM_SECTION.OnClickBSaiCodsec(Sender: TObject);
begin
{$IFNDEF PGIIMMO}
{$IFNDEF CMPGIS35}
    if (VH^.Cpta[AxeToFb(GetControlText('S_AXE'))].Structure) then ParamPlanAnal(GetControlText('S_AXE'));
    if (S_SECTION.CanFocus) then S_SECTION.SetFocus;

    S_SECTION.SelStart := Length(S_SECTION.Text);
    S_SECTION.SelLength := 0;
{$ENDIF}
{$ENDIF}
end;

procedure TOM_SECTION.OnChangeS_AXE(Sender: TObject);
begin
  SetControlProperty('S_TRANCHEGENEDE','Plus','RB_CLASSERUB="TRA" AND RB_AXE="'+GetControlText('S_AXE')+'"');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 03/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_SECTION.OnClickConfidentiel(Sender : TObject);
begin
    if (Sender = nil) then exit;
//    if (not bCharge) then
//    begin
        DS.Edit;
        if TCheckBox(Sender).Checked then SetField('S_CONFIDENTIEL','1')
        else SetField('S_CONFIDENTIEL','0') ;
//    end;
end;

{*****************************************************************}
procedure TOM_SECTION.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
    Vide : boolean ;
begin
    Vide := (Shift = []);

    if (Vide) then
    begin
        Case (Key) of
            VK_RETURN :
                begin
                    if (Ecran.ActiveControl is TCustomMemo) then exit;
                    NextControl(Ecran);
                end;
            else TFFiche(Ecran).FormKeyDown(Sender,Key,Shift);
        end;
    end;
end;

{ Procédures de zoom sur les états à partir de la fiche compte }

{*****************************************************************}
procedure TOM_SECTION.OnEnterBMenuZoom(Sender: TObject);
begin
    { FQ 20370 BVE 22.05.07 }
    TTabSheet(GetControl('INVISIBLES')).TabVisible := true;
    PopZoom97(BMenuZoom,POPZ) ;
    TTabSheet(GetControl('INVISIBLES')).TabVisible := false;
    { END FQ 20370 }
end;

{*****************************************************************}
{$IFDEF COMPTA}
procedure TOM_SECTION.OnClickBBalSecGen(Sender: TObject);
begin
  {JP 14/10/05 : FQ 16863 : Branchement de la balance section / général}
  {$IFNDEF PGIIMMO}
  {$IFNDEF GCGC}
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);

  ACritEdt.CritEdt.Bal.Axe := GetControlText('S_AXE');
  ACritEdt.CritEdt.Cpt1 := GetControlText('S_SECTION');
  ACritEdt.CritEdt.Cpt2 := ACritEdt.CritEdt.Cpt1;

  ACritEdt.CritEdt.Date1 := DebutAnnee(V_PGI.DateEntree);
  ACritEdt.CritEdt.Date2 := FinAnnee(V_PGI.DateEntree);
  ACritEdt.CritEdt.DateDeb := ACritEdt.CritEdt.Date1;
  ACritEdt.CritEdt.DateFin := ACritEdt.CritEdt.Date2;
  ACritEdt.CritEdt.Qualifpiece := '';
  ACritEdt.CritEdt.Exo.Code := '';//VH^.EnCours.Code;
  ACritEdt.CritEdt.NatureEtat := neBal;

  TheData := ACritEdt;

  CPLanceFiche_BalanceAnalGen;

  TheData := nil;
{$ENDIF}
{$ENDIF}
end;
{$ENDIF}

{*****************************************************************}
procedure TOM_SECTION.OnClickBCumul(Sender: TObject);
begin
{$IFDEF COMPTA}
    CumulCpteMensuel(LeFb,GetControlText('S_SECTION'),GetControlText('S_LIBELLE'),VH^.Entree);
{$ENDIF}
end;

{*****************************************************************}
procedure TOM_SECTION.OnClickBImprimer(Sender: TObject);
begin
{$IFDEF COMPTA}

  // Gestion du multisoc : soldes des comptes repris de la table cumuls via la vue TIERSMS
  if EstTablePartagee('SECTION') then
    V_PGI.EnableTableToView := True ;


  {JP 18/01/06 : FQ 17201 : En CWas, Les Booleans FicExport et DeuxPages à True (!) posaient un problème}
  LanceEtat( 'E', 'CST', 'ANA', true, False, False, nil,
             'S_SECTION="'+GetControlText('S_SECTION')+'" and S_AXE="'+GetControlText('S_AXE')+'"',
             Ecran.Caption, false, 0, '' ) ;

  // Gestion du multisoc : désactivation de la vue TIERSMS
  V_PGI.EnableTableToView := False ;

{$ENDIF}
end;

{*****************************************************************}
procedure TOM_SECTION.OnClickBZecrimvt(Sender: TObject);
{$IFDEF COMPTA}
var
  ACritEdt : ClassCritEdt;
{$ENDIF COMPTA}
begin
{$IFDEF COMPTA}
{ FQ 20370 BVE 02.10.07 Remplacement de ZoomAna par CPMulAna_TOF }
  ACritEdt := ClassCritEdt.Create;
  try                                 
    ACritEdt.CritEdt.LibreCodes1 := '' ;// Journal
    ACritEdt.CritEdt.sCpt1 := ''; // Compte
    ACritEdt.CritEdt.Cpt1 := GetControlText('S_SECTION'); // Section
    ACritEdt.CritEdt.Bal.Axe := GetControlText('S_AXE'); // Axe
    ACritEdt.CritEdt.Exo.Code := EXRF(VH^.Entree.Code); //Exercice
    if VH^.Precedent.Code <> '' then
       ACritEdt.CritEdt.Date1 := VH^.Precedent.Deb // Date Comptable debut
    else
       ACritEdt.CritEdt.Date1 := VH^.Encours.Deb ;// Date Comptable debut
    ACritEdt.CritEdt.Date2 := V_PGI.DateEntree ;// Date Comptable fin
    TheData := ACritEdt;
{$IFNDEF CMPGIS35}
    MultiCritereAnaZoom(taConsult,ACritEdt.CritEdt);
{$ENDIF}
  finally
    ACritEdt.Free;
    TheData := nil;
  end;          
  { END FQ 20370 }
{$ENDIF}
end;

{*****************************************************************}
procedure TOM_SECTION.OnEnterS_SECTION(Sender: TObject);
begin
    SetControlEnabled('BCODSEC',    DS.state in [dsInsert] );
    SetControlEnabled('BSAICODSEC', DS.state in [dsInsert] );
end;

{*****************************************************************}
procedure TOM_SECTION.OnExitS_SECTION(Sender: TObject);
var
    szName,section : String;
    len : integer;
begin
    section := GetControlText('S_SECTION');
    len     := Length(section);

    if (DS.state = dsInsert) then
      section := Trim(section);

    if (assigned(Ecran.ActiveControl))
      then szName := UpperCase(Ecran.ActiveControl.Name)
      else szName := '';

    if ((DS.state = dsInsert) and ((szName <> 'BCODSEC') or (szName <> 'BSAICODSEC'))) then
      begin
      SetControlEnabled('BCODSEC',false);
      SetControlEnabled('BSAICODSEC',false);
      end;

    if (section = '') then
      SetControlEnabled('S_AXE',true)
    else
      begin
      SetControlEnabled('S_AXE',false);
      if (len > LgCode) then
        section := Copy(section,1,LgCode);
      if (len < LgCode) then
        section := BourreLaDonc(section,AxeToFb(QuelAxe));
      end;

    // SetControlText('S_SECTION',section);
    if section <> '' then
      SetField('S_SECTION', section);

end;

{*****************************************************************}
procedure TOM_SECTION.AffecteLe(Champ, Valeur: string);
var
    C : TControl ;
begin
    C := GetControl(Champ,true);
{$IFDEF EAGLCLIENT}
    if (C is TCheckBox) or (C is THValComboBox) or (C is TEdit) then
{$ELSE}
    if (C is TDBCheckBox) or (C is THDBValComboBox) or (C is TDBEdit) then
{$ENDIF}
    begin
        SetControlText(Champ,Valeur);
        TEdit(C).Font.Color := clRed;
    end;
end;

{*****************************************************************}
procedure TOM_SECTION.EstUnAxeChantier;
begin
    IsChantier := VH^.Cpta[AxeToFb(QuelAxe)].Chantier;
end;


{*****************************************************************}
procedure TOM_SECTION.InitModifEnSerie(StModif: string);
var
    szName,St,Champ,Valeur : string;
    i : integer;
    Bouton : TToolBarButton97;
begin
    if (DS.State = dsBrowse) then DS.Edit;

    while (StModif <> '') do
    begin
        St := ReadTokenSt(StModif);
        i := Pos('=',St); if (i > 0) then Champ := Trim(Copy(St,1,i-1));
        i := Pos('"',St); if (i > 0) then St := Trim(Copy(St,i+1,Length(St)));
        i := Pos('"',St); if (i > 0) then Valeur := Trim(Copy(St,1,i-1));
        AffecteLe(Champ,Valeur);
    end;

    for i := 0 to TToolWindow97(GetControl('PBOUTON',true)).ControlCount-1 do
    begin
        if (TToolWindow97(GetControl('PBOUTON',true)).Controls[i] is TToolBarButton97) then
        begin
            Bouton := TToolBarButton97(TToolWindow97(GetControl('PBOUTON',true)).Controls[i]);
            szName := Uppercase(Bouton.Name);
            if ((szName <> 'BVALIDER') and (szName <> 'BFERME') and (szName <> 'BHELPBTN')) then Bouton.Enabled := false;
        end;
    end;
end;

{*****************************************************************}
function TOM_SECTION.OkConfidentiel: Boolean;
begin
  result := false;
  if ((V_PGI.Confidentiel = '0') and (GetCheckBoxState('SCONFIDENTIEL') = cbChecked)) then
    begin
    MessageAlerte( TraduireMemoire('L''accès à cette section est interdit.') );   //
    Ecran.Close; // Anciennement PostMessage(Ecran.Handle,WM_CLOSE,0,0) ;
    exit ;
    end;
  SetControlVisible('GBCONFIDENTIEL',(V_PGI.Confidentiel='1'));
  SetControlEnabled('GBCONFIDENTIEL',V_PGI.Superviseur);
  result := true ;
end;

{JP 13/01/06 : FQ 17293 : Remplacement de ZOOMTABLE par DATATYPE dans les SetControlProperty !!
{*****************************************************************}
procedure TOM_SECTION.SwapttRepart;
var s : Char; {FQ 16569}
begin
  if Length(QuelAxe) > 1 then s := QuelAxe[2]
                         else Exit;

  SetControlProperty( 'S_CLEREPARTITION', 'DATATYPE', 'TTCLEREPART'  + s       );
  SetControlProperty( 'S_CORRESP1',       'DATATYPE', 'TZCORRESPSEC' + s + '1' );
  SetControlProperty( 'S_CORRESP2',       'DATATYPE', 'TZCORRESPSEC' + s + '2' );
//  SetControlProperty( 'S_BUDSECT',        'DATATYPE', 'TZBUDSEC'     + s       );

  Case s of
        '1' : LeFb := fbAxe1;
        '2' : LeFb := fbAxe2;
        '3' : LeFb := fbAxe3;
        '4' : LeFb := fbAxe4;
        '5' : LeFb := fbAxe5;
    end;

  {JP 30/09/05 : FQ 16569 : On cache les zones si les paramsoc sont décochés}
  SetControlVisible('S_CORRESP2',  GetParamSocSecur('SO_CORSA' + s + '2', True, False));
  SetControlVisible('TS_CORRESP2', GetParamSocSecur('SO_CORSA' + s + '2', True, False));
  SetControlVisible('S_CORRESP1',  GetParamSocSecur('SO_CORSA' + s + '1', True, False));
  SetControlVisible('TS_CORRESP1', GetParamSocSecur('SO_CORSA' + s + '1', True, False));

end;

{*****************************************************************}
function TOM_SECTION.VerifCoherenceTL: Boolean;
var vStControl : string ;
    i          : Integer;
    Alerte     : Boolean;
begin
  result := true;

  if (not TTabSheet(GetControl('ZL',true)).TabVisible) then
    begin
    result := true;
    exit;
    end;

  Alerte := false;

  for i := 0 to 9 do
    begin
    vStControl := 'S_TABLE'+IntToStr(i) ;
    if not GetControlEnabled(vStControl) then
      Continue;
    if GetControlText(vStControl) = '' then
      Continue;
    if not LookupValueExist( GetControl( vStControl, True ) ) then
      Alerte := true;
    end;

    // ATTENTION. Il y a des codes de tables libres qui n'existent pas. Désirez vous continuer ?
    if (Alerte) then
      if (PgiAskCancel( TraduireMemoire('ATTENTION. Il y a des codes de tables libres qui n''existent pas. Désirez vous continuer ?'),Ecran.Caption) <> mrYes) then
        result := false;

end;

{*****************************************************************}
function TOM_SECTION.VerifiSiExiste: Boolean;
begin
    result := false;
    if (ExisteCpte(GetControlText('S_SECTION'),AxeToFb(QuelAxe))) then
    begin
        Pages.ActivePage := TTabSheet(GetControl('PGENERAL',true));
        SetFocusControl('S_SECTION');
        PGIInfo( TraduireMemoire('Le code que vous avez saisi existe déjà. Vous devez le modifier.'), Ecran.Caption ) ;
        result := true;
    end;
end;

{*****************************************************************}
procedure TOM_SECTION.ChargeCumulsMultiSoc;
var lstReq  : String ;
    lQCumul : TQuery ;
begin

  if DS.State = dsInsert then Exit ;
  if not EstTablePartagee( 'SECTION' ) then Exit ;

  if Assigned( FTobCumuls ) then FreeAndNil( FTobCumuls ) ;

  lStReq := 'SELECT CU_DEBIT1 as S_TOTALDEBIT ,  CU_CREDIT1 as S_TOTALCREDIT ,'
                 + 'CU_DEBIT2 as S_TOTDEBP ,CU_CREDIT2 as S_TOTCREP ,'
                 + 'CU_DEBIT3 as S_TOTDEBE ,CU_CREDIT3 as S_TOTCREE ,'
                 + 'CU_DEBIT4 as S_TOTDEBS ,CU_CREDIT4 as S_TOTCRES ,'
                 + 'CU_DEBITAN as S_TOTDEBANO ,CU_CREDITAN as S_TOTCREANO ,'
                 + 'CU_DEBIT5 as S_TOTDEBANON1,CU_CREDIT5 as S_TOTCREANON1'
             + ' FROM CUMULS WHERE CU_TYPE="' + fbToCumulType(fbSect) + '" '
                           + 'AND CU_COMPTE1="' + GetField('S_SECTION') + '"'
                           + 'AND CU_COMPTE2="' + GetField('S_AXE') + '"' ;

  lQCumul := OpenSQL( lStReq, True ,-1,'',true);
  if not lQCumul.Eof then
    begin
    FTobCumuls := TOB.Create('_CUMULS', nil, -1 ) ;
    FTobCumuls.SelectDB('', lQCumul ) ;
    end ;
  Ferme(lQCumul);

end;

function TOM_SECTION.GetChampsMS(vStChamp: String): double;
begin
  result := 0 ;
  if DS.State = dsInsert then Exit ;

  if EstTablePartagee( 'SECTION' ) then
    begin
    if not Assigned( FTobCumuls ) then Exit ;
    result := FTobCumuls.GetDouble(  vStChamp ) ;
    end
  else result := Valeur( GetField( vStChamp ) ) ;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/06/2005
Modifié le ... :   /  /    
Description .. : FQ 15897
Mots clefs ... :
*****************************************************************}
procedure TOM_SECTION.OnMouseDownS_BlocNote(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if DS.State = dsBrowse then
    DS.Edit;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOM_SECTION.AfficheInfos ( vBoInit : boolean ) ;
var lStVal : String ;
begin

  if vBoInit then
    begin
    lStVal := StrFMontant( 0, 15, V_PGI.OkdecV, '', True ) ;
    SetControlText('S_TOTDEBP1', lStVal ) ;
    SetControlText('S_TOTDEBE1', lStVal ) ;
    SetControlText('S_TOTDEBS1', lStVal ) ;
    SetControlText('S_TOTCREP1', lStVal ) ;
    SetControlText('S_TOTCREE1', lStVal ) ;
    SetControlText('S_TOTCRES1', lStVal ) ;
    SetControlText('SSOLCREP', lStVal ) ;
    SetControlText('SSOLCREE', lStVal ) ;
    SetControlText('SSOLCRES', lStVal ) ;
    // INFOS DE DERNIERS MVT
    if EstTablePartagee('SECTION') then
      SetControlVisible( 'HGBDERNMOUV', False )       // MODE MULTISOC : si la table est partagée, on cache les infos de dernier mouvement
    else
      begin
      SetControlText('S_DATEDERNMVT1',   '' );
      SetControlText('S_DEBITDERNMVT1',  lStVal ) ;
      SetControlText('S_CREDITDERNMVT1', lStVal ) ;
      end ;
    // INFO DE MODIFICATIONS
    SetControlText( 'SDATECREATION',   FormatDateTime('dd mmm yyyy',Now)       ) ;
    SetControlText( 'SDATEMODIF',      FormatDateTime('dd mmm yyyy',Now)       ) ;
    SetControlText( 'SDATEOUVERTURE',  FormatDateTime('dd mmm yyyy',iDate1900) ) ;
    SetControlText( 'SDATEFERMETURE',  FormatDateTime('dd mmm yyyy',iDate1900) ) ;
    end
  else
    begin
    // TOTAUX DU COMPTE
    // DEBITS
    lStVal := StrFMontant( GetChampsMS('S_TOTDEBP'), 15, V_PGI.OkdecV, '', True ) ;
    SetControlText('S_TOTDEBP1', lStVal ) ;
    lStVal := StrFMontant( GetChampsMS('S_TOTDEBE'), 15, V_PGI.OkdecV, '', True ) ;
    SetControlText('S_TOTDEBE1', lStVal ) ;
    lStVal := StrFMontant( GetChampsMS('S_TOTDEBS'), 15, V_PGI.OkdecV, '', True ) ;
    SetControlText('S_TOTDEBS1', lStVal ) ;

    // CREDITS
    lStVal := StrFMontant( GetChampsMS('S_TOTCREP'), 15, V_PGI.OkdecV, '', True ) ;
    SetControlText('S_TOTCREP1', lStVal ) ;
    lStVal := StrFMontant( GetChampsMS('S_TOTCREE'), 15, V_PGI.OkdecV, '', True ) ;
    SetControlText('S_TOTCREE1', lStVal ) ;
    lStVal := StrFMontant( GetChampsMS('S_TOTCRES'), 15, V_PGI.OkdecV, '', True ) ;
    SetControlText('S_TOTCRES1', lStVal ) ;

    // SOLDES
    AfficheLeSolde('SSOLCREP', GetChampsMS('S_TOTDEBP'), GetChampsMS('S_TOTCREP') );
    AfficheLeSolde('SSOLCREE', GetChampsMS('S_TOTDEBE'), GetChampsMS('S_TOTCREE') );
    AfficheLeSolde('SSOLCRES', GetChampsMS('S_TOTDEBS'), GetChampsMS('S_TOTCRES') );

    // MAJ des info de derniers mouvements
    if EstTablePartagee('SECTION') then
      begin
      // MODE MULTISOC : si la table est partagée, on cache les infos de dernier mouvement
      SetControlVisible( 'HGBDERNMOUV', False ) ;
      end
    else
      begin
      SetControlText('S_DATEDERNMVT1',   FormatDateTime('dd mmm yyyy',GetField('S_DATEDERNMVT') ) );
      SetControlText('S_DEBITDERNMVT1',  StrFMontant( GetField('S_DEBITDERNMVT'), 15, V_PGI.OkdecV, '', True ) ) ;
      SetControlText('S_CREDITDERNMVT1', StrFMontant( GetField('S_CREDITDERNMVT'), 15, V_PGI.OkdecV, '', True ) ) ;
      end ;

    // INFO DE MODIFICATIONS
    if (not VarIsNull(GetField('S_DATECREATION')))
      then SetControlText( 'SDATECREATION', FormatDateTime('dd mmm yyyy',GetField('S_DATECREATION') ) )
      else SetControlText( 'SDATECREATION', FormatDateTime('dd mmm yyyy',Now)                         );

    if (not VarIsNull(GetField('S_DATEMODIF')))
      then SetControlText( 'SDATEMODIF', FormatDateTime('dd mmm yyyy',GetField('S_DATEMODIF')) )
      else SetControlText( 'SDATEMODIF', FormatDateTime('dd mmm yyyy',Now)                     );

    if (not VarIsNull(GetField('S_DATEOUVERTURE')))
      then SetControlText( 'SDATEOUVERTURE', FormatDateTime('dd mmm yyyy',GetField('S_DATEOUVERTURE')) )
      else SetControlText( 'SDATEOUVERTURE', FormatDateTime('dd mmm yyyy',iDate1900)                   );

    if (not VarIsNull(GetField('S_DATEFERMETURE')))
      then SetControlText( 'SDATEFERMETURE', FormatDateTime('dd mmm yyyy',GetField('S_DATEFERMETURE')) )
      else SetControlText( 'SDATEFERMETURE', FormatDateTime('dd mmm yyyy',iDate1900)                   );

    end ;

end;

{b fb 31/05/2006 FQ17680}
procedure TOM_SECTION.OnAfterFormShow;
begin
  if KelPage<>'' then
    pages.activepageindex:=StrToInt(kelPage);
end;
{e fb 31/05/2006 FQ17680}

procedure TOM_SECTION.AfficheLeSolde( vStControl : string ; vDebit , vCredit : double ) ;
var lText  : string ;
    lSolde : double ;
begin

  lSolde := Arrondi( vDebit - vCredit, V_PGI.OkdecV ) ;
  if lSolde < 0 then
    lText := StrFMontant( -1*lSolde, 15, V_PGI.OkdecV, '', True ) + ' C'
  else if lSolde > 0 then
    lText := StrFMontant( lSolde, 15, V_PGI.OkdecV, '', True ) + ' D'
  else
    lText := StrFMontant( 0, 15, V_PGI.OkdecV, '', True ) ;

  SetControlText( vStControl, lText ) ;
end;

procedure TOM_SECTION.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin

  // Empêcher le message sur fermeture dans modif en création // SBO 19/04/2007
  if ( DS.State = dsInsert ) and
     ( GetControlText('S_SECTION') = '' ) then Exit ;

  if Assigned( FAGLCloseQuery ) then
    FAGLCloseQuery(Sender, Canclose ) ;
end;

procedure TOM_SECTION.OnClickBParamTranche(Sender: TObject);
begin
{$IFDEF COMPTA}
{$IFNDEF CMPGIS35}
  CPLanceFiche_ParamTrancheCompte (GetField('S_AXE'));
{$ENDIF}
{$ENDIF}
{$IFDEF CCSTD}
  CPLanceFiche_ParamTrancheCompte (GetField('S_AXE'));
{$ENDIF}
end;

Initialization
    registerclasses([TOM_SECTION]);
end.
