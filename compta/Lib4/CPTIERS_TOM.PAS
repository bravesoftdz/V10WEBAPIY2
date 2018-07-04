{***********UNITE*************************************************
Auteur  ...... : BPY
Créé le ...... : 20/02/2003
Modifié le ... : 12/06/2007
Description .. : GCO - 03/03/2004
Suite ........ : -> Uniformisation de l'appel à FicheTiers en 2/3 et CWAS
Suite ........ : Source TOM de la TABLE : TIERS (TIERS)
Suite ........ : CA - 05/08/2004 - FQ 13921
Suite ........ : 
Suite ........ : JP 18/01/06 : FQ 17201 : gestion de l'état en eAgl
Suite ........ : 
Suite ........ : JP 05/06/06 : FQ 17971 : Modification de la gestion du 
Suite ........ : bouton BCopier et du LookUp de duplication
Suite ........ : 
Suite ........ : CA 12/06/07 : FQ 18761 : T_FACTUREHT non modifiable
Mots clefs ... : TOM;TIERS
*****************************************************************}

Unit CPTIERS_TOM;       

//================================================================================
// Interface
//================================================================================
Interface

Uses
    StdCtrls,
    Controls,
    Classes,
    CRITEDT,
    AGLInit, // TheData
    {$IFDEF COMPTA}
{$IFNDEF CMPGIS35}
    CPBALAGEE_TOF,
    CPBALAUXIGEN_TOF, // CPLanceFiche_BalanceAuxiGen
    CPBALAUXI_TOF,
{$ENDIF}
    {$ENDIF}
    UTOB,
{$IFDEF EAGLCLIENT}
    eFiche,
    eFichList,
    maineagl,
    UtileAGL,
{$ELSE}
    db,
    dbctrls,
    edtREtat,  // LanceEtat
    {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
    hdb,
    AglIsoflex,
    Fiche,
    FichList,
    FE_MAIN,
    codepost,
    region,
{$ENDIF}
    Hqry,
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    HEnt1,
    HMsgBox,
    UTOM,
    utilPGI,
    ent1,
    HTB97,
    ParamSoc,
    FichComm,
    Buttons,
    Graphics,
    menus,
{$IFDEF COMPTA} // GCO - 24/05/2004 - Ajout IFDEF COMPTA
    ZECRIMVT_TOF,
    Cummens,
{$ENDIF}
    ULibWindows,
    UtilSais,
    choix,
{$IFDEF VER150}
    variants,
{$ENDIF}
    CPGENERAUX_TOM,
    CPCODEPOSTAL_TOF,
    HRichOle,         // THRichEditOle
    HRichEdt,         // THRichEditEdt
    CPREGION_TOF,
    LookUp  ,UentCommun              {FP 21/02/2006 Pour LookUpList}
    ;

//==================================================
// Externe
//==================================================

Procedure FicheTiers(Q : TQuery; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer );
Procedure FicheTiersMZS(Compte : String;Comment : TActionFiche;QuellePage : Integer;LesModif : String);
Procedure FicheTiersZoom(Q : TQuery; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer ; Range : string );

//==================================================
// Definition de class
//==================================================
Type
    TOM_TIERS = Class (TOM)
    public

      {$IFDEF EAGLCLIENT}
        T_Auxiliaire : THEdit;               
      {$ELSE}
        T_Auxiliaire : THDBEdit;
      {$ENDIF}

        // variable cencé etre accessible de l'exterieur ....
        LesModif,Lequel : string;
        FStRange : string ;

        procedure OnKeyPressT_Auxiliaire(Sender: TObject; var Key: Char);

    private
        // gestion des messagebox et menuzoom
        POPS : TPopUpMenu;
        MsgBox : THMsgBox;
        // sauvegarde du mode
        Mode : TActionFiche;
        // sauvegarde de la nature des compte et du pageCtrl
        Nature : THValComboBox;
        pages : TPageControl;
        // variable de svg
        MNat,ML,Dev,MultiDev, RTVA, TVAE : String ;
        AvecMvt,EnSaisieCpta : Boolean ;
        MultiDevise : Boolean ; // FQ 19238
        LgCode : Byte ;
        fbForceSaisieLibelle : boolean;
        FInPageIndex : integer ;
        IsUpdating : Boolean;
        FCritEdt : ClassCritEdt;
{$IFDEF EAGLCLIENT}
        ChangeLibelle : Boolean;  // FQ 20695
{$ENDIF}
        FTobCumuls              : TOB ;                    // Contient les infos de cumuls locales en mode multisoc
        FAGLCloseQuery          : TCloseQueryEvent ;

    public
        procedure OnNewRecord                ; override ;
        procedure OnDeleteRecord             ; override ;
        procedure OnUpdateRecord             ; override ;
        procedure OnAfterUpdateRecord        ; override ;
        procedure OnLoadRecord               ; override ;
        procedure OnChangeField(F : TField)  ; override ;
        procedure OnArgument(S : String)     ; override ;
        procedure OnClose                    ; override ;
        procedure OnCancelRecord             ; override ;
    Private

        procedure OnAfterFormShow;
        procedure FormCloseQuery             (Sender: TObject; var CanClose: Boolean);
        // evenement sur les controls
        procedure OnChangeNature(Sender : TObject);
        procedure OnChangeDevise(Sender: TObject);
        procedure OnChangeAuxiliaire(Sender: TObject);
        procedure OnChangeDebCred(Sender: TObject);
        //FP FQ15570 procedure OnValideDebCred(Sender: TObject);
        procedure OnChangeSCHEMAGEN(Sender : TObject);
        procedure OnExitSCHEMAGEN(Sender : TObject);
        procedure OnChangeACCELERATEUR(Sender : TObject);

        procedure OnChangingPages ( Sender : TObject; var AllowChange: Boolean);

        procedure OnDblClickCodePostalVille(Sender: TObject);
        procedure OnDblClickDivTerri(Sender: TObject);

        procedure OnClickContact(Sender: TObject);
        procedure OnClickRib(Sender: TObject);
        procedure OnClickLettrable(Sender: TObject);
        procedure OnClickImprimer(Sender: TObject);
        procedure OnClickMultiDevise(Sender: TObject);
        procedure OnClickIsPayeur(Sender: TObject);
        procedure OnClickConfidentiel(Sender: TObject);

        procedure OnEnterMenuZoom(Sender: TObject);
{$IFDEF COMPTA}
        procedure OnClickBBalAge(Sender: TObject);
        procedure OnClickBBalVen(Sender: TObject);        
{$ENDIF}
{$IFDEF COMPTA}
        procedure OnClickBBalAux(Sender: TObject);
        procedure OnClickBBalAuxGen(Sender: TObject);
{$ENDIF}
        procedure OnClickBCumul(Sender: TObject);
        procedure OnClickBJustif(Sender: TObject);
        procedure OnClickBZecrimvt(Sender: TObject);

        procedure OnExitAuxiliaire(Sender: TObject);
        procedure OnExitLibelle (Sender : TObject );
{$IFDEF EAGLCLIENT}
        procedure OnChangeLibelle (Sender : TObject );  // FQ 20695
{$ENDIF}
        procedure OnExitFacture(Sender: TObject);
        procedure OnExitCollectif(Sender: TObject);
        procedure OnExitPays(Sender: TObject);
        // nouvel enreg
        procedure OnClickBInsert(Sender: TObject);
        procedure OnElipsisClickCORRESP(Sender: TObject);  {FP 21/02/2006}

        // assignation de fonction
        procedure AssignFcts;
        // formatage et activation des zones
        procedure FormateLesMontants;
        procedure AfficheCumuls ;
        procedure AutoriseReleveFacture;
        procedure PourSalarie;
        procedure EnabledGB(GBname : string);
        procedure InitModif;
        procedure OnEstEnSaisieCompta;
        procedure ActivationControl;
        // test de confidentialité
        function OkConfidentiel: Boolean;
        // test sur les caracteristique du compte
        procedure LettrageChgAutorise;
        procedure NatureChgAutorise;
        procedure PayeurEna;
        function CodeValide: Boolean;
        function CodeCollectifValide: Boolean;
        function AccelerateurValide : boolean ;
        // verification de coherence
        function VerifCoherenceTL: Boolean;
        function LettrageChanger: Boolean;
        function TVAChanger: Boolean;
        // test de mouvement pour le compte
        function IsMouvemente: Boolean;
        // Test du multidevise (FQ 19238)
        function IsMultiDevise: Boolean;
        // init des messagebox et menuzoom
        procedure InitMsgBox;
        // Modification en serie
        procedure AffecteLe(Champ, Valeur: string);
        procedure InitModifEnSerie(StModif: string);
        // delete
        procedure DelRibContact;
        procedure CreateEnregGescomGRC;
        procedure CreationTIERSCOMPL ;

        //Gestion DAS 2
        procedure OnchangeDAS2(Sender : Tobject);
        procedure OnChangeRemuneration(Sender : Tobject) ;
        procedure OnChangeProfession(Sender : Tobject) ;
        procedure ForceModif;
{$IFNDEF CCADM}
      {$IFDEF CCSTD}
        // Pas besoin de la  DADS2 dans les standards
      {$ELSE}
        procedure MAJDads2; // BVE 12.09.07
      {$ENDIF}
{$ENDIF CCADM}


{$IFNDEF EAGLCLIENT}
        //ISOFLEX
        function GereIsoflex: Boolean;
        procedure AppelIsoflex(Sender: TObject);
{$ENDIF}
        procedure OnMouseDownT_BlocNote(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);

        // MULTISOC
        function  GetChampsMS( vStChamp : String ) : double ;   // Retourne le cumul demandé soit issu de la table commune, soit de la table locale en multisoc
        procedure ChargeCumulsMultiSoc ;                     // Renseigne FTobCumuls avec les valeurs locales en mode multisoc

        // COMPENSATION
        function VerifCorrespondance: Byte;                   {FP 21/02/2006}
        procedure BCopierClick(Sender: TObject);
    end;

//================================================================================
// Implementation
//================================================================================
Implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPProcGen,
  CPTypeCons,
  CPObjetGen,
  CPVersion,
  {$ENDIF MODENT1}

  HCompte, {JP 02/06/06 : FQ 17971 : THCpteEdit}
 CbpMCD,
 {$IFNDEF CCADM}
{$IFDEF CCSTD}
{$ELSE}
 {$IFNDEF PGIIMMO}
{$IFNDEF CMPGIS35}
  PgDADSCommun,
{$ENDIF}
 {$ENDIF}
{$ENDIF}
{$ENDIF CCADM}

  {$IFDEF eAGLCLIENT}
  MenuOLX
  {$ELSE}
  MenuOLG
  {$ENDIF eAGLCLIENT}
  ,UFonctionsCBP
  ;

//==================================================
// Definition des Constante
//==================================================
// les tag de pages ;)
const
    PGeneral     : integer = 0;
    PComptable   : integer = 1;
    PInformation : integer = 2;
    PCommercial  : integer = 3;
    TReglement   : integer = 4;
    TDAS2        : integer = 5;
    ZL           : integer = 6;
    // YMO 03/2006 Champs à copier à partir du compte source par le bouton dupliquer
    DuplicValues : array[1..38] of string = ('T_LETTRABLE','T_MULTIDEVISE','T_DEVISE','T_CORRESP1','T_CORRESP2',
                   'T_REGIMETVA','T_TVAENCAISSEMENT','T_SOUMISTPF','T_TOTAUXMENSUELS','T_SOLDEPROGRESSIF',
                   'T_SAUTPAGE','T_BLOCNOTE','T_TARIFTIERS','T_FACTUREHT','T_REMISE','T_ESCOMPTE','T_FACTURE',
                   'T_CREDITDEMANDE','T_CREDITACCORDE','T_CREDITPLAFOND','T_FACTURE','T_MODEREGLE','T_JOURPAIEMENT1',
                   'T_JOURPAIEMENT2','T_ISPAYEUR','T_AVOIRRBT','T_DEBRAYEPAYEUR','T_PAYEUR','T_RELEVEFACTURE',
                   'T_JOURRELEVE','T_FREQRELEVE','T_DATEDERNRELEVE','T_RELANCEREGLEMENT','T_RELANCETRAITE',
                   'T_LETTREPAIEMENT','T_MOTIFVIREMENT','T_RESIDENTETRANGER','T_NATUREECONOMIQUE');

//==================================================
// fonctions hors class
//==================================================

Procedure FicheTiersZoom(Q : TQuery; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer ; Range : string );
begin
 // Lancement de la fiche selon le contexte de creation modification ou consultation
  Case Comment of
    // creation
    taCreat, taCreatEnSerie, taCreatOne :
    begin
      if (not ExJaiLeDroitConcept(TConcept(ccAuxCreat),True)) then exit;
         AGLLanceFiche('CP','CPFICHETIERS',Axe,Compte,'ACTION=CREATION;'+IntToStr(QuellePage)+';ZOOM;'+Range);
    end;

    // modification
    taModif,taModifEnSerie :
    begin
      if (not ExJaiLeDroitConcept(TConcept(ccAuxModif),True)) then exit;
      // if (_Blocage(['nrCloture','nrBatch','nrSaisieModif','nrEnca','nrDeca'],True,'nrBatch')) then exit;
        AGLLanceFiche('CP','CPFICHETIERS',Axe,Compte,'ACTION=MODIFICATION;'+IntToStr(QuellePage));
    end;

    // Consultation
    taConsult : AGLLanceFiche('CP','CPFICHETIERS',Axe,Compte,'ACTION=CONSULTATION;'+IntToStr(QuellePage));
  end;

  if (Comment in [taModif,taModifEnSerie]) then _Bloqueur('nrBatch',False);
end;


{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 20/02/2003
Modifié le ... : 03/03/2004
Description .. : Point d'entrée
Mots clefs ... :
*****************************************************************}
procedure FicheTiers(Q : TQuery; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer );
var lStAction : String ;
begin
  // Lancement de la fiche selon le contexte de creation modification ou consultation
  lStAction := ActionToString( Comment ) ;
  Case Comment of
    // creation
    taCreat, taCreatEnSerie, taCreatOne :
    begin
      if (not ExJaiLeDroitConcept(TConcept(ccAuxCreat),True)) then exit;
      if lStAction='' then
        lStAction := 'ACTION=CREATION' ;
      AGLLanceFiche('CP','CPFICHETIERS',Axe,Compte,lStAction+';'+IntToStr(QuellePage));
    end;

    // modification
    taModif,taModifEnSerie :
    begin
      if (not ExJaiLeDroitConcept(TConcept(ccAuxModif),True)) then exit;
      if lStAction='' then
        lStAction := 'ACTION=MODIFICATION' ;
      AGLLanceFiche('CP','CPFICHETIERS',Axe,Compte,lStAction+';'+IntToStr(QuellePage));
    end;

    // Consultation
    taConsult : AGLLanceFiche('CP','CPFICHETIERS',Axe,Compte,'ACTION=CONSULTATION;'+IntToStr(QuellePage));
  end;

  if (Comment in [taModif,taModifEnSerie]) then _Bloqueur('nrBatch',False);
end;



{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 26/02/2003
Modifié le ... : 26/02/2003
Description .. : point d'entré pour la modif en serie
Mots clefs ... :
*****************************************************************}
Procedure FicheTiersMZS(Compte : String;Comment : TActionFiche;QuellePage : Integer;LesModif : String);
begin
    // Lancement de la fiche selon le contexte de creation modification ou consultation
    case Comment of
        taCreat,taCreatEnSerie,taCreatOne :
            begin
                if Not ExJaiLeDroitConcept(TConcept(ccAuxCreat),True) then Exit;
                AGLLanceFiche('CP','CPFICHETIERS','','','ACTION=CREATION;'+IntToStr(QuellePage)+';'+LesModif);
            end;
        taModif,taModifEnSerie :
            begin
                if Not ExJaiLeDroitConcept(TConcept(ccAuxModif),True) then Exit;
                AGLLanceFiche('CP','CPFICHETIERS','',Compte,'ACTION=MODIFICATION;'+IntToStr(QuellePage)+';'+LesModif);
            end;
    end;
end;

//==================================================
// Evenements par default de la TOM
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 20/02/2003
Modifié le ... : 23/01/2006
Description .. : Initialisation de la fiche
Suite ........ : - LG - 11/01/2005 - FB 14671 - recuperation des info de la 
Suite ........ : saisie en cours pour initialiser le code, la nature et le tiers 
Suite ........ : par defaut de l'auxi
Suite ........ : - LG - 23/01/2005 - replacement par des getparamsocsecur
Mots clefs ... : 
*****************************************************************}
procedure TOM_TIERS.OnNewRecord;
var i : integer;
    LL: THLabel;
begin
    Inherited;

    fbForceSaisieLibelle := False;
    if VH^.PaysLocalisation=codeISOES then
       SetControlProperty('T_SIRET','CharCase',ecUpperCase) ; //XVI 24/02/2005
    // mouvementé ?? un nouveau ??
    AvecMvt := False ;

    // active les control
    nature.enabled := true;
    SetControlEnabled('T_LETTRABLE',true);
    SetControlEnabled('T_AUXILIAIRE',true);


    if FStRange <> '' then
     begin
      // text de la nature
      nature.value := Copy(FStRange,16,3) ;
      SetField('T_NATUREAUXI',nature.value);
     end
      else
       begin
       // text de la nature
       nature.value := 'DIV';
       SetField('T_NATUREAUXI','DIV');
      end ;
      
    OnChangeNature(THValComboBox(GetControl('T_NATUREAUXI')));
    
    SetControlText('NATAUX',nature.text);

    // mise des valeur par default
    SetField('T_DEVISE',V_PGI.DevisePivot);
    // check par default
    SetField('T_SOLDEPROGRESSIF','X');
    SetField('T_LETTRABLE','X');
    SetField('T_CONFIDENTIEL','0');
    SetField('T_FACTUREHT','X');
//ajout me fiche 10416    if (ctxPCL in V_PGI.PGIContexte) then
    begin
        SetField('T_REGIMETVA',GetParamSocSecur('SO_REGIMEDEFAUT','FRA'));
        SetField('T_MODEREGLE',GetParamSocSecur('SO_GCMODEREGLEDEFAUT','001'));
        SetField('T_TVAENCAISSEMENT',GetParamSocSecur('SO_CODETVADEFAUT','TM'));
    end;

    // focus du control
    Pages.ActivePageIndex := PComptable;
    if (assigned(GetControl('T_AUXILIAIRE',true))) then TWinControl(GetControl('T_AUXILIAIRE',true)).Setfocus;
    // init
    LgCode:=VH^.Cpta[fbAux].Lg ;

    // set des libellé des tableau
    for i := 0 to 9 do
    begin
        LL := THLabel(GetControl('H_'+IntToStr(i),true));
        if (LL <> nil) then LL.Caption := '';
    end;

    if (CtxPcl in V_Pgi.PGIContexte) then
      ActivationControl;

    THLabel(GetControl('G_LIBELLE',true)).Caption := '' ;

    // Affichage des date
    SetControlText('TDATECREATION',FormatDateTime('dd mmm yyyy',Now));
    SetControlText('TDATEMODIF',FormatDateTime('dd mmm yyyy',Now));
    SetControlText('TDATEOUVERTURE',FormatDateTime('dd mmm yyyy',iDate1900));
    SetControlText('TDATEFERMETURE',FormatDateTime('dd mmm yyyy',iDate1900));

  SetControlEnabled('BCopier', True); {JP 05/06/06 : FQ17971}
    // ici l'AGL set les zone de l'ecran avec les variable du DataSet !
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 20/02/2003
Modifié le ... : 26/02/2003
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_TIERS.OnDeleteRecord;
begin
    Inherited;

    // error par default !!
    LastError := 1;

    // a effacé ??
    if MsgBox.Execute(1,'','')<>mrYes then Exit ;

    if (BourreEtLess(GetControlText('T_AUXILIAIRE'),fbAux) = '') then Exit ;
    // delete des CONTACT RIB et TARIF du compte
    ExecuteSQL('Delete from CONTACT Where C_AUXILIAIRE="'+GetControlText('T_AUXILIAIRE')+'"');
    ExecuteSQL('Delete from RIB Where R_AUXILIAIRE="'+GetControlText('T_AUXILIAIRE')+'"');
//    ExecuteSQL('Delete from TARIF Where F_AUXILIAIRE="'+GetControlText('T_AUXILIAIRE')+'"'); => Table tarif non encore créée le 20/02/97

    // si on arrive ici alors ... ca marche ;)
    LastError := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 20/02/2003
Modifié le ... : 26/02/2003
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_TIERS.OnUpdateRecord;
var ModPaie : String ;
    SIRET : String ; //XMG 19/01/04
begin
    Inherited;

    // est on en modif ??
    if (not (mode in [taCreat,taCreatEnSerie,taCreatOne,taModif,taModifEnSerie])) then
      begin
      LastError := 1;
      Exit;
      end ;

    // test de validité des code ??
    if (Not CodeValide)                      then
      begin
      LastError := 1;
      Exit;
      end ;
    if (Not CodeCollectifValide)             then
      begin
      LastError := 1;
      Exit;
      end ;

    // libele vide ??
    if (GetControlText('T_LIBELLE') = '') then
    begin
        Pages.ActivePageIndex := PGeneral;
        SetFocusControl('T_LIBELLE');
        MsgBox.execute(3,'','');
        LastError := 1;
        Exit;
    end;

    //YMO 04/07/2006 FQ18366 Vérification du tiers payeur
    if (Trim(GetControlText('T_PAYEUR'))<>'') and (Not ExisteSQL('SELECT 1 FROM TIERS WHERE T_AUXILIAIRE ="'+GetControlText('T_PAYEUR')+'"')) then
    begin
        Pages.ActivePageIndex := TReglement;
        SetFocusControl('T_PAYEUR');
        MsgBox.Execute(51,'','');
        LastError := 1;
        Exit;
    end;

    // etre son propre payeur ??
    if (GetControlText('T_PAYEUR') = GetControlText('T_AUXILIAIRE')) then
    begin
        Pages.ActivePageIndex := TReglement;
        SetFocusControl('T_PAYEUR');
        MsgBox.Execute(40,'','');
        LastError := 1;
        Exit;
    end;
    // mode de reglement ??
    if (GetControlText('T_MODEREGLE') = '') then
    begin
        Pages.ActivePageIndex := TReglement;
        SetFocusControl('T_MODEREGLE');
        MsgBox.execute(31,'','');
        LastError := 1;
        Exit;
    end;
    // regime TVA ??
    if ((GetControlText('T_REGIMETVA') = '') and (nature.value <> 'SAL')) then
    begin
        Pages.ActivePageIndex := PComptable;
        SetFocusControl('T_REGIMETVA');
        MsgBox.execute(32,'','');
        LastError := 1;
        Exit;
    end;
    if (not VerifCoherenceTL) then
      begin
      LastError := 1;
      Exit ;
      end ;

    // verif de ??
    Case VerifCorrespondance of   {FP 21/02/2006}
        0 : ;
        1 : begin Msgbox.Execute(37,'',''); LastError := 1; Exit; end;
        2 : begin Msgbox.Execute(38,'',''); LastError := 1; Exit; end;
    end;

  if (not AccelerateurValide) then
   begin
      LastError := 1;
      Exit ;
   end ;

   if VH^.PaysLocalisation=CodeISOES then
   begin
    //Pour l'espagne, on veirifie le SIRET, si le tiers est Espagnol.
    SIRET:=VerifNIF_ES(GetField('T_SIRET')) ;
    if (CodeISODuPays(Getfield('T_PAYS'))=CodeISOES) and (GetField('T_SIRET')<>SIRET) then
    begin
       if MsgBox.Execute(43,'','')= mrYes then
          SetField('T_SIRET',SIRET)
       else
       begin
          SetField('T_SIRET','') ;
          SetFocusCOntrol('T_SIRET') ;
          LastError:=1 ;
          Exit ;
       End ;
    End ;
   End ;//XVI 24/02/2005
    // on est en edition ??
    if (mode in [taModif,taModifEnSerie]) then
    begin
        if (LettrageChanger) then
        begin
            if ((Not _BlocCarFiche) and (MsgBox.Execute(23,'','') = mrYes)) then
            begin
                MsgBox.Execute(24,'','');
                ModPaie := Choisir(MsgBox.Mess[25],'MODEPAIE','MP_LIBELLE','MP_MODEPAIE','','');
                if (ModPaie <> '') then MajLettrageEcriture(ModPaie,GetControlText('T_AUXILIAIRE'),fbAux)
                else
                begin
                    MsgBox.Execute(26,'','');
                    SetControlText('T_LETTRABLE',ML);
                end;
            end
            else SetControlText('T_LETTRABLE',ML);
        end;

        {FQ19946 20.12.2007 YMO Message d'avertissement sur modif paramètres TVA sur un compte}
        if (TVAChanger) then MsgBox.Execute(52,'','');

    end;

// BPY le 11/07/2003
//    completion du T_TIERS ... ce champs existe t'il ??
      if GetField('T_TIERS')='' then
        SetField('T_TIERS',BourreEtLess(GetControlText('T_AUXILIAIRE'),fbAux));

    // selon la nature
    if (not ((nature.value = 'CLI') or (nature.value = 'AUD'))) then
    begin
        SetControlText('T_RELANCEREGLEMENT','');
        SetControlText('T_RELANCETRAITE','');
    end;
    if (nature.Value <> 'CLI') then
    begin
        SetControlText('T_RELEVEFACTURE','-');
        SetControlText('T_JOURRELEVE','0');
        SetControlText('T_FREQRELEVE','');
    end;

// BPY le 11/07/2003
    if (mode in [taCreat,taCreatEnSerie,taCreatOne]) then
    begin
      SetField('T_CODEIMPORT',GetControlText('T_AUXILIAIRE'));
{$IFDEF COMPTA}
  // fiche 20549
    MAJHistoparam ('TIERS',  GetControlText('T_AUXILIAIRE'));
{$ENDIF}

    end;
  // ============================
  // ==== GESTION DU MULTISOC ===
  // ============================
  // Ajout de l'enregistrement local de la table cumuls pour gestion des totaux en multisoc
// En attente car pour le moment n'enregistre que sur la base courante...
//  if  DS.State in [DsInsert] then
//    CInsertCumulsMS( fbAux, GetField('T_AUXILIAIRE'), '' ) ;

   { FQ 20474 BVE 31.05.07 }
   SetField('T_RELANCEREGLEMENT', GetControlText('T_RELANCEREGLEMENT'));
   SetField('T_RELANCETRAITE',    GetControlText('T_RELANCETRAITE'));
   SetField('T_MODEREGLE',        GetControlText('T_MODEREGLE'));
   { END FQ 20474 }
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 20/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_TIERS.OnAfterUpdateRecord;
begin
  if mode in  [taCreat,taCreatEnSerie,taCreatOne] then CreateEnregGescomGRC;
    CreationTIERSCOMPL ;
  Inherited;
  IsUpdating := True;
{$IFNDEF CCADM}
{$IFDEF CCSTD}
  // Pas besoin de la  DADS2 dans les standards
{$ELSE}
  { FQ 20733 BVE 12.09.07 }
  if ( mode in [taModif,taModifEnSerie]) and
  ( ExisteSQL('SELECT 1 FROM DADS2HONORAIRES WHERE PDH_HONORAIRE = "' + GetField('T_AUXILIAIRE') + '"')) then
  begin
     MAJDads2;
  end;
  { END FQ 20733 }
{$ENDIF}
{$ENDIF CCADM}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 12/09/2007
Modifié le ... : 12/09/2007
Description .. : Permet de mettre à jour l'enregistrement contenu dans la 
Suite ........ : table DADSHONORAIRES à partir des données qui 
Suite ........ : viennent d'être saisies
Suite ........ : FQ 20733
Mots clefs ... : 
*****************************************************************}
{$IFNDEF CCADM}
{$IFDEF CCSTD}
        // Pas besoin de la  DADS2 dans les standards
{$ELSE}
procedure TOM_TIERS.MAJDads2;
 {$IFNDEF PGIIMMO}
{$IFNDEF CMPGIS35}
var
  TobAuxi, TobDads2 : TOB;
  Q  : TQuery;
  SQL: String;
  Controle : TControle;
 {$ENDIF}
{$ENDIF}
begin
 {$IFNDEF PGIIMMO}
{$IFNDEF CMPGIS35}
  try
     TobDads2 := TOB.Create('DADS2HONORAIRES',nil,-1);
     TobAuxi  := TOB.Create('TIERS',nil,-1);

     // Chargement des données
     TobAuxi.GetEcran(ecran);
     SQL := 'SELECT * FROM DADS2HONORAIRES WHERE PDH_HONORAIRE = "' + GetField('T_AUXILIAIRE') + '"';
     Q := OpenSQL(SQL,false);
     while not Q.Eof do
     begin
        TobDads2.SelectDB('',Q);
        TiersToDads2(TobAuxi,TobDads2,Controle);
        TobDads2.UpdateDB();
        Q.Next;
     end;
  finally
     if assigned(TobAuxi) then FreeAndNil(TobAuxi);
     if assigned(TobDads2) then FreeAndNil(TobDads2);
     if assigned(Q) then Ferme(Q);
  end;
 {$ENDIF}
{$ENDIF}
end;
{$ENDIF}
{$ENDIF CCADM}

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 20/02/2003
Modifié le ... : 10/07/2007
Description .. : Chargement du record
Suite ........ : - LG  - 16/08/2004 - FB 13677 - correction des
Suite ........ : deplacements avec la souris
Suite ........ : - LG - 18/08/2044 - utilisation des fct setcontrol...
Suite ........ : - LG - 11/05/2007 - FB 20144 - on recupere le valeurs 
Suite ........ : eventuelle du LeQuel pour preremplir la fiche
Suite ........ : - LG - 10/07/2007 - FB 20979 - les modif en serie ne fct 
Suite ........ : plus
Mots clefs ... : 
*****************************************************************}
procedure TOM_TIERS.OnLoadRecord;
var
    lQuery : TQuery ;
begin
    Inherited;
    IsUpdating := False;
    if VH^.PaysLocalisation=codeISOES then
       SetControlProperty('T_SIRET','CharCase',ecNormal) ; //XVI 24/02/2005

    SetControlChecked('TCONFIDENTIEL', (GetField('T_CONFIDENTIEL') = '1') ) ;

    // chargement des infos de cumuls MULTISOC
    ChargeCumulsMultiSoc ;

    // Affichge des cumuls
    AfficheCumuls ;

    // Affichage des date
     if (not VarIsNull(GetField('T_DATECREATION'))) then
      SetControlText('TDATECREATION',FormatDateTime('dd mmm yyyy',GetField('T_DATECREATION')))
    else SetControlText('TDATECREATION',FormatDateTime('dd mmm yyyy',Now));
    if (not VarIsNull(GetField('T_DATEMODIF'))) then
      SetControlText('TDATEMODIF',FormatDateTime('dd mmm yyyy',GetField('T_DATEMODIF')))
    else SetControlText('TDATEMODIF',FormatDateTime('dd mmm yyyy',Now));
    if (not VarIsNull(GetField('T_DATEOUVERTURE'))) then
      SetControlText('TDATEOUVERTURE',FormatDateTime('dd mmm yyyy',GetField('T_DATEOUVERTURE')))
    else SetControlText('TDATEOUVERTURE',FormatDateTime('dd mmm yyyy',iDate1900));
    if (not VarIsNull(GetField('T_DATEFERMETURE'))) then
      SetControlText('TDATEFERMETURE',FormatDateTime('dd mmm yyyy',GetField('T_DATEFERMETURE')))
    else SetControlText('TDATEFERMETURE',FormatDateTime('dd mmm yyyy',iDate1900));

    // Recuperation des mvt !!
    AvecMvt  := IsMouvemente;
    { FQ 19238 BVE 02.07.07 }
    MultiDevise := IsMultiDevise;
    { END FQ 19238 }

    TCheckBox(GetControl('YTC_ACCELERATEUR',true)).OnClick       := nil ;
    THEdit(GetControl('YTC_SCHEMAGEN',true)).OnExit              := nil ;
    THEdit(GetControl('YTC_SCHEMAGEN',true)).OnChange            := nil ;
    TCheckBox(GetControl('YTC_DAS2',true)).OnClick               := nil;
    THValComboBox(GetControl('YTC_REMUNERATION',true)).Onchange  := nil ;
    THEdit(GetControl('YTC_PROFESSION',true)).Onchange           := nil ;
    THEdit(GetControl('T_AUXILIAIRE',true)).OnChange             := nil ;
    THValComboBox(GetControl('T_DEVISE',true)).OnChange          := nil ;

    SetControlText('YTC_SCHEMAGEN', '' ) ;
    SetControlCaption('G_LIBELLE', '' ) ;
    SetControlChecked('YTC_ACCELERATEUR', false ) ;
    SetControlChecked('YTC_DAS2', false ) ;
    SetControlText('YTC_PROFESSION', '' ) ;
    SetControlProperty('YTC_REMUNERATION','Value', '' ) ;
    SetControlProperty('YTC_INDEMNITE','Value', '' ) ;
    SetControlProperty('YTC_AVANTAGE','Value', '' ) ;
    SetControlProperty('YTC_RETENUESOURCE','Value', '' ) ;
    SetControlEnabled('G_DAS2',False);
    SetControlEnabled('YTC_PROFESSION',False);
    SetControlEnabled('YTC_REMUNERATION',False);
    SetControlEnabled('YTC_RETENUESOURCE',False);

    lQuery := OpenSql('select YTC_SCHEMAGEN,YTC_ACCELERATEUR, YTC_DAS2, YTC_PROFESSION, YTC_REMUNERATION, YTC_INDEMNITE, YTC_AVANTAGE, YTC_RETENUESOURCE from TIERSCOMPL where YTC_AUXILIAIRE="'+ BourreEtLess(GetControlText('T_AUXILIAIRE'), fbAux)+'" ' , true ) ;
                              // ajout me fiche 16217
    if (not lQuery.EOF) and  not (Mode in [taCreat, taCreatEnSerie, taCreatOne])then
    begin
      SetControlText('YTC_SCHEMAGEN', lQuery.findfield('YTC_SCHEMAGEN').AsString ) ;
      SetControlChecked('YTC_ACCELERATEUR', lQuery.findfield('YTC_ACCELERATEUR').AsString = 'X' ) ;
      //DAS2
      if (GetControlText('T_NATUREAUXI')='FOU') then
      begin
        SetControlChecked('YTC_DAS2', lQuery.findfield('YTC_DAS2').AsString = 'X' ) ;
        SetControlText('YTC_PROFESSION', lQuery.findfield('YTC_PROFESSION').AsString ) ;
        SetControlProperty('YTC_REMUNERATION','Value',lQuery.findfield('YTC_REMUNERATION').AsString ) ;
        SetControlProperty('YTC_INDEMNITE','Value',lQuery.findfield('YTC_INDEMNITE').AsString ) ;
        SetControlProperty('YTC_AVANTAGE','Value',lQuery.findfield('YTC_AVANTAGE').AsString ) ;
        SetControlProperty('YTC_RETENUESOURCE','Value',lQuery.findfield('YTC_RETENUESOURCE').AsString ) ;
        SetControlEnabled('G_DAS2',lQuery.findfield('YTC_DAS2').AsString = 'X');
        SetControlEnabled('YTC_PROFESSION',lQuery.findfield('YTC_DAS2').AsString = 'X');
        SetControlEnabled('YTC_REMUNERATION',lQuery.findfield('YTC_DAS2').AsString = 'X');
        SetControlEnabled('YTC_RETENUESOURCE',lQuery.findfield('YTC_DAS2').AsString = 'X');
      end ;
    end; //lQuery
    Ferme(lQuery) ;

    if VH^.Lequel <> '' then
     InitModifEnSerie(VH^.Lequel)
      else
     // modif en serie
       if lesModif <> '' then //
       //if (Mode = taModifEnSerie) then
        InitModifEnSerie(LesModif);

    THEdit(GetControl('YTC_SCHEMAGEN',true)).OnExit              := OnExitSCHEMAGEN ;
    TCheckBox(GetControl('YTC_ACCELERATEUR',true)).OnClick       := OnChangeACCELERATEUR ;
    THEdit(GetControl('YTC_SCHEMAGEN',true)).OnChange            := OnChangeSCHEMAGEN ;
    TCheckBox(GetControl('YTC_DAS2',true)).OnClick               := OnChangeDAS2 ;
    THValComboBox(GetControl('YTC_REMUNERATION',true)).Onchange  := OnChangeRemuneration ;
    THEdit(GetControl('YTC_PROFESSION',true)).Onchange           := OnChangeProfession ;
    THEdit(GetControl('T_AUXILIAIRE',true)).OnChange             := OnChangeAuxiliaire;
    THValComboBox(GetControl('T_DEVISE',true)).OnChange          := OnChangeDevise;

    OnExitSCHEMAGEN(nil) ;

    SetControlEnabled('YTC_SCHEMAGEN', GetCheckBoxState('YTC_ACCELERATEUR') = cbChecked ) ;
//    Pages.pages[TDAS2].TabVisible := (GetControlText('T_NATUREAUXI')='FOU') and (VH^.PaysLocalisation<>CodeISOES); //XVI 24/02/2005
// CA - 26/03/2008 - Onglet DAS2 non visible en localisation différente de France
    Pages.pages[TDAS2].TabVisible := (GetControlText('T_NATUREAUXI')='FOU') and (VH^.PaysLocalisation=CodeISOFR);

    // active certain controls selon les données
    PourSalarie;
    InitModif;
    OnEstEnSaisieCompta;
    AutoriseReleveFacture;
    NatureChgAutorise;
    LettrageChgAutorise;
    PayeurEna;
    
    if FInPageIndex = 2 then
    begin
      SetFocusControl('T_BLOCNOTE');
    end;

    // YMO Bouton dupliquer acessible uniquement à la création
    SetControlEnabled('BCopier', (DS.State = dsInsert)); {JP 05/06/06 : FQ17971}

    // ici l'AGL set les zone de l'ecran avec les variable du DataSet !
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 20/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_TIERS.OnChangeField(F: TField);
begin
    Inherited;
    // changement de nature
    if (ctxPCL in V_PGI.PGIContexte) then
    begin
        // FQ 16383 - Prise en compte mode consultation
        nature.Enabled := (not AvecMvt) and (Mode<>taConsult);
        if (nature.Enabled) then nature.Color := ClWindow
        else nature.Color := ClBtnFace;
    end;

    // en saisie ??
    if (EnSaisieCpta) then SetControlEnabled('BInsert',false);

    // set de l'abregé
    if (GetControlText('T_ABREGE') = '') then SetControlText('T_ABREGE',Copy(GetControlText('T_LIBELLE'),1,17));
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 20/02/2003
Modifié le ... : 11/01/2005
Description .. : Initialisation de la fiche
Suite ........ : - LG - 11/01/2005 - FB 14671 - recuperation des info de la 
Suite ........ : saisie en cours pour initialiser le code, la nature et le tiers
Suite ........ : par defaut de l'auxi
Mots clefs ... : 
*****************************************************************}
procedure TOM_TIERS.OnArgument(S: String);
var
    {$IFDEF EAGLCLIENT}
    lEdit     : THEdit ;
    lCheckBox : TCheckBox ;
    lCombo    : THValComboBox;
    {$ELSE}
    lEdit     : THDBEdit ;
    lCheckBox : TDBCheckBox ;
    lCombo    : THDBValComboBox;
    BIsoFlex : TToolBarButton97;
    {$ENDIF}
    Brib,Bctt,imprimer : TToolBarButton97;
    button : TbitBtn;
    button97 : TToolBarButton97;
begin
    Inherited;

    FCritEdt := ClassCritEdt.Create;

    IsUpdating := False;
    TFFiche(Ecran).OnAfterFormShow := Self.OnAfterFormShow;
    TFFiche(Ecran).HelpContext := 7142000; // 14371

    // Empêcher le message sur fermeture dans modif en création // SBO 19/04/2007
    FAGLCloseQuery := TFFiche(Ecran).OnCloseQuery ;
    TFFiche(Ecran).OnCloseQuery := FormCloseQuery ;

    // Recuperation du combo de nature  et du page ctrl et set d'une fct si existe et tue tt si existe pas !
    Nature := THValComboBox(GetControl('T_NATUREAUXI',true));
    pages  := TPageControl(GetControl('PAGES',true));

    SetControlText('YTC_SCHEMAGEN', '' ) ;
    SetControlCaption('G_LIBELLE', '' ) ;
    SetControlChecked('YTC_ACCELERATEUR', false ) ;
    //Gestion de la DAS2
    SetControlChecked('YTC_DAS2' , false ) ;
    SetControlChecked('G_DAS2' , false ) ;

    TToolBarButton97(GetControl('BINSERT',true)).OnClick := OnClickBInsert;
    TToolBarButton97(GetControl('BCOPIER',True)).OnClick := BCopierClick;

    // creation des messagebox et menuzoom
    MsgBox := THMsgBox.create(FMenuG);
    InitMsgBox;
    POPS := TPopUpMenu(GetControl('POPS',true));

    // lecture du mode et sauvegarde
    ReadTokenSt(S);                           // lecture du mode pour AGL
    if TFFiche(Ecran).TypeAction = taCreatEnSerie then
      TFFiche(Ecran).MonoFiche := False ;
    Mode := TFFiche(Ecran).TypeAction ;

 //   if (S <> '') then Pages.ActivePageIndex := strtoint(ReadTokenSt(S)); // lecture de la page si il y a lieu
    if (S <> '') then FInPageIndex := strtoint(ReadTokenSt(S)) ; // lecture de la page si il y a lieu

    if (S <> '') then LesModif := S ;                        // recup des modif pour la modificationen serie
    if Copy(LesModif,1,4) = 'ZOOM' then
     begin
      LesModif := '' ;
      if (S <> '') then ReadTokenSt(S) ;
      if (S <> '') then FStRange := ReadTokenSt(S) ;
    end ;

    // set de EnSaisieCpta
    if (Mode = taModif) then EnSaisieCpta := SaisieLancer;

    // set du LgCode
    if (Mode <> taConsult) then LgCode := VH^.Cpta[fbAux].Lg;

    // recup du checkbox confidentiel
    TCheckBox(GetControl('TCONFIDENTIEL',true)).OnClick := OnClickConfidentiel;

    // fonction de zoom sur le cope postalle/ville
    {$IFDEF EAGLCLIENT}
    lEdit     := THEdit(GetControl('T_CODEPOSTAL',true));
    {$ELSE}
    lEdit     := THDBEdit(GetControl('T_CODEPOSTAL',true));
    {$ENDIF}
    if assigned(lEdit) then
      lEdit.OnDblClick := OnDblClickCodePostalVille;

    {$IFDEF EAGLCLIENT}
    lEdit     := THEdit(GetControl('T_VILLE',true));
    {$ELSE}
    lEdit     := THDBEdit(GetControl('T_VILLE',true));
    {$ENDIF}
    if assigned(lEdit) then
      lEdit.OnDblClick := OnDblClickCodePostalVille;


    {$IFDEF EAGLCLIENT}
    lEdit     := THEdit(GetControl('T_FACTURE',true));
    {$ELSE}
    lEdit     := THDBEdit(GetControl('T_FACTURE',true));
    {$ENDIF}
    if assigned(lEdit) then
      begin
      lEdit.OnExit := OnExitFacture;
      // Ajout condition sur tablette TZTTOUS
      lEdit.plus := 'T_NATUREAUXI<>"NCP" AND T_NATUREAUXI<>"CON" AND T_NATUREAUXI<>"PRO" AND T_NATUREAUXI<>"SUS" Order by T_AUXILIAIRE';
      end ;

    // ajout d'un control sur l'exit de tier facturé et de collectif
    {$IFDEF EAGLCLIENT}
    lEdit     := THEdit(GetControl('T_COLLECTIF',true));
    {$ELSE}
    lEdit     := THDBEdit(GetControl('T_COLLECTIF',true));
    {$ENDIF}
    if assigned(lEdit) then
      lEdit.OnExit := OnExitCollectif;

    // fonction sur le DblClick de la division territorial
    {$IFDEF EAGLCLIENT}
    lEdit     := THEdit(GetControl('T_DIVTERRIT',true));
    {$ELSE}
    lEdit     := THDBEdit(GetControl('T_DIVTERRIT',true));
    {$ENDIF}
    if assigned(lEdit) then
      lEdit.OnDblClick := OnDblClickDivTerri;

    // fontion sur le click de lettrable
    {$IFDEF EAGLCLIENT}
    lCheckBox     := TCheckBox(GetControl('T_LETTRABLE',true));
    {$ELSE}
    lCheckBox     := TDBCheckBox(GetControl('T_LETTRABLE',true));
    {$ENDIF}
    if assigned(lCheckBox) then
      lCheckBox.OnClick := OnClickLettrable;

    // fontion sur le click de ispayeur
    {$IFDEF EAGLCLIENT}
    lCheckBox     := TCheckBox(GetControl('T_ISPAYEUR',true));
    {$ELSE}
    lCheckBox     := TDBCheckBox(GetControl('T_ISPAYEUR',true));
    {$ENDIF}
    if assigned(lCheckBox) then
      lCheckBox.OnClick := OnClickIsPayeur;

    // fontion sur le click de multidevise
    {$IFDEF EAGLCLIENT}
    lCheckBox     := TCheckBox(GetControl('T_MULTIDEVISE',true));
    {$ELSE}
    lCheckBox     := TDBCheckBox(GetControl('T_MULTIDEVISE',true));
    {$ENDIF}
    if assigned(lCheckBox) then
      lCheckBox.OnClick := OnClickMultiDevise;

    // ajout des fonction sur les nouveaux boutons
    Brib := TToolBarButton97(GetControl('BRIB',true));
    if (assigned(Brib)) then Brib.OnClick := OnClickRib;
    Bctt := TToolBarButton97(GetControl('BCONTACT',true));
    if (assigned(Bctt)) then Bctt.OnClick := OnClickContact;
    // ajout des fonction sur le bouton imprimer
    imprimer := TToolBarButton97(GetControl('BImprimer',true));
    if (assigned(imprimer)) then imprimer.OnClick := OnClickImprimer;

    // ajout de controls sur l'exit et le change d'auxiliaire
{$IFDEF EAGLCLIENT}
    T_Auxiliaire := THEdit(GetControl('T_AUXILIAIRE', True));
{$ELSE}
    T_Auxiliaire := THDBEdit(GetControl('T_AUXILIAIRE', True));
{$ENDIF}
    T_Auxiliaire.MaxLength  := VH^.Cpta[fbAux].Lg;
    T_Auxiliaire.OnExit     := OnExitAuxiliaire;
    T_Auxiliaire.OnKeyPress := OnKeyPressT_Auxiliaire;

    // ajout des valeur "aucune" au combo des relance
    {$IFDEF EAGLCLIENT}
    lCombo     := THValComboBox(GetControl('T_RELANCETRAITE',true));
    {$ELSE}
    lCombo     := THDBValComboBox(GetControl('T_RELANCETRAITE',true));
    {$ENDIF}
    lCombo.vide := True ;
    lCombo.VideString := TraduireMemoire('Aucune');
    lCombo.Reload;

    // ajout des valeur "aucune" au combo des relance
    {$IFDEF EAGLCLIENT}
    lCombo     := THValComboBox(GetControl('T_RELANCEREGLEMENT',true));
    {$ELSE}
    lCombo     := THDBValComboBox(GetControl('T_RELANCEREGLEMENT',true));
    {$ENDIF}
    lCombo.vide := True ;
    lCombo.VideString := TraduireMemoire('Aucune');

    // ajout de fonctions onchange sur les combo pays
    {$IFDEF EAGLCLIENT}
    lCombo     := THValComboBox(GetControl('T_PAYS',true));
    {$ELSE}
    lCombo     := THDBValComboBox(GetControl('T_PAYS',true));
    {$ENDIF}
    if assigned(lCombo) then
      lCombo.OnExit := OnExitPays;

    THcritMaskEdit(GetControl('T_LIBELLE')).OnExit := OnExitLibelle;
{$IFDEF EAGLCLIENT}
    { FQ 20695 BVE 19.06.07 }
    ChangeLibelle := false;
    THEdit(GetControl('T_LIBELLE' )).OnChange := OnChangeLibelle;
    THEdit(GetControl('T_LIBELLE_')).OnChange := OnChangeLibelle;
    { END FQ 20695 }
    THRichEditOle(GetControl('T_BLOCNOTE', True)).OnMouseDown := OnMouseDownT_BlocNote;
    SetControlVisible('BIMPRIMER', True); {JP 18/01/06 : FQ 17201}
{$ELSE}
    THDBRichEditOle(GetControl('T_BLOCNOTE', True)).OnMouseDown := OnMouseDownT_BlocNote;
{$ENDIF}


    // si confidentiel .....
    if (not OkConfidentiel) then Exit ;

    // set des visibles des navcontrol si existant !
    SetControlVisible('BFirst',(not V_PGI.MonoFiche));
    SetControlVisible('BPrev',(not V_PGI.MonoFiche));
    SetControlVisible('BNext',(not V_PGI.MonoFiche));
    SetControlVisible('BLast',(not V_PGI.MonoFiche));

    // set des fonction de menuzoom
    button97 := TToolBarButton97(GetControl('BMenuZoom',true));  if (assigned(button97)) then button97.OnMouseEnter := OnEnterMenuZoom;
    // set des fonction d'impression Zoom
{$IFDEF COMPTA}
    button := TBitBtn(GetControl('BBalAge',true));    if (assigned(button)) then button.OnClick := OnClickBBalAge;
    button := TBitBtn(GetControl('BBalAuxGen',true)); if (assigned(button)) then button.OnClick := OnClickBBalAuxGen;
    button := TBitBtn(GetControl('BBalVen',true));    if (assigned(button)) then button.OnClick := OnClickBBalVen;
    button := TBitBtn(GetControl('BBalAux',true));    if (assigned(button)) then button.OnClick := OnClickBBalAux;
{$ELSE}
  SetControlEnabled ('BBalAge',False);
  SetControlEnabled ('BBalAuxGen',False);
  SetControlEnabled ('BBalVen',False);
{$ENDIF}
    button := TBitBtn(GetControl('BCumul',true));     if (assigned(button)) then button.OnClick := OnClickBCumul;
    button := TBitBtn(GetControl('BJustif',true));    if (assigned(button)) then button.OnClick := OnClickBJustif;
    button := TBitBtn(GetControl('BZecrimvt',true));  if (assigned(button)) then button.OnClick := OnClickBZecrimvt;

    {JP 30/09/05 : FQ 16569 : On cache les zones si les paramsoc sont décochés}
    SetControlVisible('T_CORRESP2',  GetParamSocSecur('SO_CORSAU2', True, True));
    SetControlVisible('TT_CORRESP2', GetParamSocSecur('SO_CORSAU2', True, True));
    SetControlVisible('T_CORRESP1',  GetParamSocSecur('SO_CORSAU1', True, True));
    SetControlVisible('TT_CORRESP1', GetParamSocSecur('SO_CORSAU1', True, True));

    // test de la serie et desactivation des control correspondant !
    if (EstSerie(S5)) then      // si on est en S5
    begin
        {FP 25/04/2006 FQ17909 Il faut autoriser les 2 plans de correspondance
        SetControlVisible('T_CORRESP2',False);
        SetControlVisible('TT_CORRESP2',False);}
        SetControlVisible('T_LANGUE',False);
        SetControlVisible('TT_LANGUE',False);
    end;
 { YMO 08/06/2006 FQ18217 Affichage des zones onglet règlement
 else if EstSerie(S3) then   // si on est en S3
    begin
        {FP 25/04/2006 FQ17909 Il faut autoriser les 2 plans de correspondance
        SetControlVisible('T_CORRESP2',False);
        SetControlVisible('TT_CORRESP2',False);
        SetControlVisible('T_LANGUE',False);
        SetControlVisible('TT_LANGUE',False);
        {FP 25/04/2006 FQ17909 Il faut autoriser les 2 plans de correspondance
        SetControlVisible('T_CORRESP1',False);
        SetControlVisible('TT_CORRESP1',False);
        SetControlVisible('T_CONSO',False);
        SetControlVisible('TT_CONSO',False);
        SetControlVisible('T_ISPAYEUR',False);
        SetControlVisible('T_AVOIRRBT',False);
        SetControlVisible('T_DEBRAYEPAYEUR',VH^.PaysLocalisation=CodeISOES); //XVI 24/02/2005
        SetControlVisible('T_PAYEUR',VH^.PaysLocalisation=CodeISOES);
        SetControlVisible('TT_PAYEUR',VH^.PaysLocalisation=CodeISOES);
        SetControlVisible('T_SOUMISTPF',VH^.PaysLocalisation=CodeISOES);
        SetControlVisible('TCONFIDENTIEL',False);
        SetControlVisible('T_JOURPAIEMENT1',VH^.PaysLocalisation=CodeISOES);
        SetControlVisible('TT_JOURPAIEMENT1',VH^.PaysLocalisation=CodeISOES);
        SetControlVisible('T_JOURPAIEMENT2',VH^.PaysLocalisation=CodeISOES);
        SetControlVisible('TT_JOURPAIEMENT2',VH^.PaysLocalisation=CodeISOES);

        SetControlEnabled('bBalven.Enabled',VH^.PaysLocalisation=CodeISOES);
        SetControlEnabled('bBalAuxGen.Enabled',VH^.PaysLocalisation=CodeISOES);
    end;
    }

    // traduction des tables libres !
    LibellesTableLibre(Pages.Pages[ZL],'TT_TABLE','T_TABLE','T') ;

    // set des fonction sur les control double
    AssignFcts;

    // formatage des zones
    FormateLesMontants;
{$IFNDEF EAGLCLIENT}
    if V_PGI.ModePcl <> '1' then
    begin
      if GereIsoflex then
      begin
        BIsoFlex := TToolBarButton97(GetControl('BISOFLEx'));

        if BISOFLEX <> nil then Bisoflex.onClick := AppelIsoflex;
      end;
    end;
{$ENDIF}

  {b FP 21/02/2006}
  SetControlProperty('T_CORRESP1', 'DataType', 'CPCPTECORRESPAUX1');
  SetControlProperty('T_CORRESP2', 'DataType', 'CPCPTECORRESPAUX2');

  if TCompensation.IsPlanCorresp(1) then
    begin
    {$IFDEF EAGLCLIENT}THEdit{$ELSE}THDBEdit{$ENDIF}
    (GetControl('T_CORRESP1',true)).OnElipsisClick := OnElipsisClickCORRESP;
    end
  else if TCompensation.IsPlanCorresp(2) then
    begin
    {$IFDEF EAGLCLIENT}THEdit{$ELSE}THDBEdit{$ENDIF}
    (GetControl('T_CORRESP2',true)).OnElipsisClick := OnElipsisClickCORRESP;
    end;
  {e FP 21/02/2006}
  { CA - 12/06/2007 - FQ 18761 }
  SetControlEnabled('T_FACTUREHT',False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 20/02/2003
Modifié le ... : 23/11/2006
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_TIERS.OnClose;
var Auxiliaire : string;
begin
  // supprimé les rib et contacte cree !
  Auxiliaire := Trim(GetField('T_AUXILIAIRE'));
  if Auxiliaire <> '' then  
  begin
    if (not (ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="' + auxiliaire + '"'))) then
      DelRibContact;
  end;

  FreeAndNil(MsgBox);
  FreeAndNil(FCritEdt);
  if Assigned( FTobCumuls ) then FreeAndNil( FTobCumuls ) ; // SBO Ajout MultiSoc

  Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 20/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_TIERS.OnCancelRecord;
begin
    Inherited;
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/06/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_TIERS.OnClickBInsert(Sender: TObject);
begin
    if (ExJaiLeDroitConcept(TConcept(ccAuxCreat),true)) then
      TFFiche(Ecran).BInsertClick(Sender);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 20/02/2003
Modifié le ... : 18/12/2006
Description .. : Changement de la nature du compte .....
Suite ........ : CA - 18/12/2006 - On dégage la variable payeur qui ne sert
Suite ........ :  à rien.
Mots clefs ... :
*****************************************************************}
procedure TOM_TIERS.OnChangeNature(Sender: TObject);
begin
  // Positionnement du collectif
  if IsUpdating then Exit; // FFF 185
  if AvecMvt then Exit ;
  if (GetControlText('T_NATUREAUXI')='CLI') then SetField('T_COLLECTIF',VH^.DefautCli)
  else if (GetControlText('T_NATUREAUXI')='FOU') then SetField('T_COLLECTIF',VH^.DefautFou)
  else if (GetControlText('T_NATUREAUXI')='SAL') then SetField('T_COLLECTIF',VH^.DefautSal)
  else if (GetControlText('T_NATUREAUXI')='AUD') then SetField('T_COLLECTIF',VH^.DefautDDivers)
  else if (GetControlText('T_NATUREAUXI')='AUC') then SetField('T_COLLECTIF',VH^.DefautCDivers)
  else if (GetControlText('T_NATUREAUXI')='DIV') then SetField('T_COLLECTIF',VH^.DefautDivers);

  //DAS2 Visible que si "FOU" et pas Espagne...
//  Pages.pages[5].TabVisible := (GetControlText('T_NATUREAUXI')='FOU') and (VH^.PaysLocalisation<>CodeISOES) ;
// CA - 26/03/2008 - Onglet DAS2 non visible en localisation différente de France
    Pages.pages[TDAS2].TabVisible := (GetControlText('T_NATUREAUXI')='FOU') and (VH^.PaysLocalisation=CodeISOFR);

  // FQ 16689 : ne pas cocher automatiquement DAS 2
  // SetControlChecked('YTC_DAS2', (GetControlText('T_NATUREAUXI')='FOU') and (VH^.PaysLocalisation<>CodeISOES)) ;


  // set de la nature sur la premiere page !
  SetControlText('NATAUX',nature.Text);

  // set du zoom table sur les tier payeur
  if (nature.Value = '') then Exit ;

  // client et debiteur divers
  if ((nature.Value = 'CLI') or (nature.Value = 'AUD')) then
    SetControlProperty('T_PAYEUR','Plus','T_NATUREAUXI="AUD" OR T_NATUREAUXI="CLI" OR T_NATUREAUXI="DIV"')
  // fournisseur et crediteur divers
  else if ((nature.Value = 'FOU') or (nature.Value = 'AUC')) then
    SetControlProperty('T_PAYEUR','Plus','T_NATUREAUXI="AUC" OR T_NATUREAUXI="FOU" OR T_NATUREAUXI="DIV"')
  // divers
  else if (nature.Value = 'DIV') then
    SetControlProperty('T_PAYEUR','Plus','')
  // salarié ... pas de choix ;)
  else if (nature.Value = 'SAL') then
    SetControlProperty('T_PAYEUR','Plus','');

  AutoriseReleveFacture;
  PourSalarie;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.OnChangeDevise(Sender: TObject);
var
    devise : THValComboBox;
begin
    devise := THValComboBox(GetControl('T_DEVISE',true));
    if (not assigned(devise)) then exit;
    if (GetCheckBoxState('T_MULTIDEVISE') = cbChecked) then exit;
    if ((Dev <> devise.value) and (AvecMvt) and (Dev <> '')) then
     devise.value := Dev ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.OnChangeAuxiliaire(Sender: TObject);
begin
    ActivationControl;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/02/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_TIERS.OnChangingPages(Sender: TObject; var AllowChange: Boolean);
begin
  if (not (CtxPcl in V_Pgi.PGIContexte)) then Exit;

  AllowChange := (Trim(GetControlText('T_Auxiliaire')) <> '');

  if (not AllowChange) then
  begin
    if Trim(GetControlText('T_Auxiliaire')) = '' then
      SetFocusControl('T_Auxiliaire')
    else
      SetFocusControl('T_NatureAuxi');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/02/2003
Modifié le ... : 24/02/2003
Description .. : verification des code postal
Mots clefs ... :
*****************************************************************}
procedure TOM_TIERS.OnDblClickCodePostalVille(Sender: TObject);
begin
  VerifCodePostal(nil,THEdit(GetControl('T_CODEPOSTAL',true)),THEdit(GetControl('T_VILLE',true)),(THEdit(Sender).name = 'T_CODEPOSTAL'));
  {JP 21/07/05 : FQ 15603 : en fait seuls les contrôles sont mis à jours dans VerifCodePostal, mais
                 pas les champs => l'enregistrement ne tenait pas compte des modifications et
                 remettait les valeurs antérieures}
  SetField('T_CODEPOSTAL', GetControlText('T_CODEPOSTAL'));
  SetField('T_VILLE', GetControlText('T_VILLE'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.OnDblClickDivTerri(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
    PaysRegion(THValComboBox(GetControl('T_PAYS',true)),THEdit(GetControl('T_DIVTERRIT',true)),True) ;
{$ELSE}
    PaysRegion(THDBValComboBox(GetControl('T_PAYS',true)),THDBEdit(GetControl('T_DIVTERRIT',true)),True) ;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/02/2003
Modifié le ... : 24/02/2003
Description .. : Click sur le bouton rib
Mots clefs ... :
*****************************************************************}
procedure TOM_TIERS.OnClickRib(Sender: TObject);
Var
    Ok : Boolean;
    LeMode : TActionFiche;
begin
    // init du mode ...
    LeMode := Mode;

    if (DS.State = dsInsert) then
    begin
      SetControlText('T_AUXILIAIRE',BourreLaDonc(GetControlText('T_AUXILIAIRE'),fbAux)) ;
      if (Not CodeValide) then Exit ;
    end;

    // on test si il existe deja un BIB
    Ok := ExisteSQL('Select R_AUXILIAIRE from RIB where R_AUXILIAIRE="' + GetControlText('T_AUXILIAIRE') + '"');
    if ((LeMode = taConsult) And (Not Ok)) then
    begin
        Msgbox.execute(19,'','');
        Exit;
    end;

    // modification du mode d'utilisation de la fiche RIB
    if ((LeMode <> taConsult) And (Not Ok)) then if (Msgbox.execute(20,'','') <> mrYes) then Exit
                                                 else LeMode := taCreatOne;
    if (Ok And (LeMode in [taCreat..taCreatOne])) then LeMode := taModif;

    // lancement de la fiche
//    FicheRIB_AGL(GetControlText('T_AUXILIAIRE'),LeMode,False,'',TRUE); Lek 310506 FQ17466 passe à FromSaisie pour message RIB
    FicheRIB_AGL(GetControlText('T_AUXILIAIRE'),LeMode,True,'',TRUE);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/02/2003
Modifié le ... : 24/02/2003
Description .. : Click sur le bouton contact
Mots clefs ... :
*****************************************************************}
procedure TOM_TIERS.OnClickContact(Sender: TObject);
var
    Ok : Boolean;
    LeMode : TActionFiche;
begin
    // init du mode ...
    LeMode := Mode;

    if (DS.State = dsInsert) then
    begin
      SetControlText('T_AUXILIAIRE',BourreLaDonc(GetControlText('T_AUXILIAIRE'),fbAux)) ;
      if (Not CodeValide) then Exit;
    end;

    // on test si il existe deja un contact
    Ok := ExisteSQL('Select C_AUXILIAIRE from CONTACT where C_AUXILIAIRE="'+GetControlText('T_AUXILIAIRE')+'"');
    if ((LeMode=taConsult) And (Not Ok)) then
    begin
        Msgbox.Execute(21,'','');
        Exit;
    end;

    // modification du mode d'utilisation de la fiche contact
    if ((LeMode <> taConsult) And (Not Ok)) then if (Msgbox.Execute(22,'','') <> mrYes) then Exit
                                                 else LeMode := taCreat;
    if (Ok And (LeMode in [taCreat..taCreatOne])) then LeMode := taModif;

    // lancement de la fiche
    FicheContact_AGL(GetControlText('T_AUXILIAIRE'),LeMode);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/02/2003
Modifié le ... : 24/02/2003
Description .. : changement des controls actif si on change le lettrable
Mots clefs ... :
*****************************************************************}
procedure TOM_TIERS.OnClickLettrable(Sender: TObject);
begin
    AutoriseReleveFacture;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.OnClickImprimer(Sender: TObject);
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}

  // Gestion du multisoc : soldes des comptes repris de la table cumuls via la vue TIERSMS
  if EstTablePartagee('TIERS') then
    V_PGI.EnableTableToView := True ;

  {JP 18/01/06 : FQ 17201 : En CWas, Les Booleans FicExport et DeuxPages à True (!) posaient un problème}
  LanceEtat( 'E', 'CST', 'AUX', true, False, False, nil,
             'T_AUXILIAIRE="'+GetControlText('T_AUXILIAIRE')+'"',
             Ecran.Caption, false, 0, '' ) ;

  // Gestion du multisoc : désactivation de la vue TIERSMS
  V_PGI.EnableTableToView := False ;

{$ENDIF}
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.OnClickMultiDevise(Sender: TObject);
var
    devise : THValComboBox;
begin
    // get du control !
    devise := THValComboBox(GetControl('T_DEVISE',true));
    if (not assigned(devise)) then exit;

    // test
    if ((MultiDev = 'X') And (GetCheckBoxState('T_MULTIDEVISE') = cbUnChecked)) then
    begin
        if (AvecMvt and MultiDevise) then
        begin
            SetControlChecked('T_MULTIDEVISE',(MultiDev = 'X'));
            devise.value := '' ;
        end
        else
        begin
            if (Dev = '') then devise.value := V_PGI.DevisePivot
            else devise.value := Dev;
        end;
    end
//    else if (GetCheckBoxState('T_MULTIDEVISE') = cbChecked) then devise.value := ''
    // CA - 02/11/2006 - On permet l'inidcation d'une devise par défaut en multidevises.
  //  else if (GetCheckBoxState('T_MULTIDEVISE') = cbChecked) then SetField ('T_DEVISE', V_PGI.devisepivot)
    else if ((MultiDev = '-') And (GetCheckBoxState('T_MULTIDEVISE') = cbUnChecked)) then
    begin
        if (Dev = '') then devise.value := V_PGI.DevisePivot
        else devise.value := Dev ;
    end;

    // enable du control
    // CA - 02/11/2006 - Suite demande GPAO, on permet de saisir une devise par défaut en multi devises
    //if (GetCheckBoxState('T_MULTIDEVISE') = cbUnChecked) then devise.Enabled := True
    //else devise.Enabled := False ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.OnClickIsPayeur(Sender: TObject);
var
    Qpayeur : TQuery;
begin
    // est ce que l'on decoche
    if ((TCheckBox(Sender).Checked = false) and (GetField('T_ISPAYEUR') = 'X')) then
    begin
        // y a t'il des tiers associer
        Qpayeur := OpenSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_PAYEUR="' + GetControlText('T_AUXILIAIRE') + '"',true);
        if (not (QPayeur.EOF)) then
        begin
            // creation de laliste
            while (not (QPayeur.EOF)) do
            begin
                if (MsgBox.mess[47] = '') then MsgBox.mess[47] := Qpayeur.Fields[0].AsString
                else MsgBox.mess[47] := MsgBox.mess[47] + ', ' + Qpayeur.Fields[0].AsString;
                Qpayeur.next;
            end;
            // creation du message
            MsgBox.mess[47] := MsgBox.mess[45] + MsgBox.Mess[47];
            MsgBox.mess[47] := MsgBox.mess[47] + MsgBox.mess[46];
            // message
            MsgBox.Execute(47,'','');
{ FQ 18617 BVE 11.04.07 }
            MsgBox.mess[47] := '';
{ END FQ 18617 BVE 11.04.07 }
            TCheckBox(Sender).Checked := true;
            Ferme(Qpayeur);
            exit;
        end;
        Ferme(Qpayeur);
    end;

    PayeurEna;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 03/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.OnClickConfidentiel(Sender : TObject);
begin
    if (Sender = nil) then exit;
//    if (not bCharge) then
//    begin
        if (DS.State = dsBrowse) then DS.Edit; // SBO 18/08/2004 FQ 14217 ne forcer le mode édition que si nécessaire
        if TCheckBox(Sender).Checked then SetField('T_CONFIDENTIEL','1')
        else SetField('T_CONFIDENTIEL','0') ;
//    end;
end;

{ Procédures de zoom sur les états à partir de la fiche compte }

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.OnEnterMenuZoom(Sender: TObject);
begin
    // attention marche pas en EAGL .... a cause d'un ctrl de merde qui a pas de parent ;)
    PopZoom97(TToolBarButton97(GetControl('BMENUZOOM',true)),POPS);
end;


{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
Procedure TOM_TIERS.OnClickBBalAux(Sender: TObject);
Var
    ACritEdt : ClassCritEdt;
    D1,D2 : TdateTime;
    Etab : string;
begin
  ACritEdt := ClassCritEdt.Create;
  try
    Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
    D1 := VH^.Encours.Deb;
    D2 := VH^.Encours.Fin;
    if (VH^.Entree.Code = VH^.Suivant.Code) then
    begin
        D1 := VH^.Suivant.Deb;
        D2 := VH^.Suivant.Fin;
    end;
    ACritEdt.CritEdt.Date1:=D1 ;
    ACritEdt.CritEdt.DateDeb:=ACritEdt.CritEdt.Date1 ;
    ACritEdt.CritEdt.Date2:=D2 ;
    ACritEdt.CritEdt.DateDeb:=ACritEdt.CritEdt.Date2 ;
    ACritEdt.CritEdt.Cpt1  := GetControlText('T_AUXILIAIRE');
    ACritEdt.CritEdt.Cpt2  := GetControlText('T_AUXILIAIRE');
    ACritEdt.CritEdt.SCpt1 := '';
    ACritEdt.CritEdt.SCpt2 := '';
    Etab:=EtabForce ; if Etab<>'' then ACritEdt.CritEdt.Etab:=Etab ;
    TheData := ACritEdt;
{$IFNDEF CMPGIS35}
    CPLanceFiche_BalanceAuxiliaire ( ) ;
{$ENDIF}
  finally
    ACritEdt.Free;
    TheData := nil;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.OnClickBBalAge(Sender: TObject);
Var
    ACritEdt : ClassCritEdt;
    D1 : TdateTime;
    Etab : string;
begin
  ACritEdt := ClassCritEdt.Create;
  try
    Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
    D1 := VH^.Encours.Deb;
    if (VH^.Entree.Code = VH^.Suivant.Code) then D1 := VH^.Suivant.Deb;
    ACritEdt.CritEdt.Date1:=D1 ;
    ACritEdt.CritEdt.DateDeb:=ACritEdt.CritEdt.Date1 ;
    ACritEdt.CritEdt.Cpt1  := GetControlText('T_AUXILIAIRE');
    ACritEdt.CritEdt.Cpt2  := GetControlText('T_AUXILIAIRE');
    ACritEdt.CritEdt.SCpt1 := '';
    ACritEdt.CritEdt.SCpt2 := '';
    Etab:=EtabForce ; if Etab<>'' then ACritEdt.CritEdt.Etab:=Etab ;
    TheData := ACritEdt;
{$IFNDEF CMPGIS35}
    CPLanceFiche_BalanceAgee ( ) ;
{$ENDIF}
  finally
    ACritEdt.Free;
    TheData := nil;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.OnClickBBalVen(Sender: TObject);
Var
    ACritEdt : ClassCritEdt;
    D1 : TdateTime;
    Etab : string;
begin
  ACritEdt := ClassCritEdt.Create;
  try
    Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
    D1 := VH^.Encours.Deb;
    if (VH^.Entree.Code = VH^.Suivant.Code) then D1 := VH^.Suivant.Deb;
    ACritEdt.CritEdt.Date1:=D1 ;
    ACritEdt.CritEdt.DateDeb:=ACritEdt.CritEdt.Date1 ;
    ACritEdt.CritEdt.Cpt1  := GetControlText('T_AUXILIAIRE');
    ACritEdt.CritEdt.Cpt2  := GetControlText('T_AUXILIAIRE');
    ACritEdt.CritEdt.SCpt1 := '';
    ACritEdt.CritEdt.SCpt2 := '';
    Etab:=EtabForce ; if Etab<>'' then ACritEdt.CritEdt.Etab:=Etab ;
    TheData := ACritEdt;
{$IFNDEF CMPGIS35}
    CPLanceFiche_BalanceVentilee ( ) ;
{$ENDIF}
  finally
    ACritEdt.Free;
    TheData := nil;
  end;
end;
{$ENDIF}


{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... : 19/11/2004
Description .. : - BPY le 19/11/2004 - Fiche n° 12165 - correction de l'etat
Suite ........ : pour fonctionnement en CWAS. + uniformisation CWAS 2/3
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
Procedure TOM_TIERS.OnClickBBalAuxGen(Sender: TObject);
var
    Etab : String;
begin
    Fillchar(FCritEdt.CritEdt, SizeOf(FCritEdt.CritEdt), #0);

    // Exercice
    if (VH^.Entree.Code = VH^.Suivant.Code) then
    begin
        FCritEdt.CritEdt.Exo.Code := VH^.Suivant.Code;
        FCritEdt.CritEdt.Date1 := VH^.Suivant.Deb;
        FCritEdt.CritEdt.Date2 := VH^.Suivant.Fin;
    end
    else
    begin
        FCritEdt.CritEdt.Exo.Code := VH^.Encours.Code;
        FCritEdt.CritEdt.Date1 := VH^.Encours.Deb;
        FCritEdt.CritEdt.Date2 := VH^.Encours.Fin;
    end;
    FCritEdt.CritEdt.Qualifpiece := 'N';

    FCritEdt.CritEdt.Cpt1  := GetControlText('T_AUXILIAIRE');
    FCritEdt.CritEdt.Cpt2  := GetControlText('T_AUXILIAIRE');
    Etab := EtabForce;
    if (Etab <> '') then FCritEdt.CritEdt.Etab := Etab;

    TheData := FCritEdt;
{$IFNDEF CMPGIS35}
    CPLanceFiche_BalanceAuxiGen;
{$ENDIF}
    TheData := nil;
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 26/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.OnClickBJustif(Sender: TObject);
begin
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 26/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.OnClickBCumul(Sender: TObject);
begin
{$IFDEF COMPTA}
    CumulCpteMensuel(fbAux,GetControlText('T_AUXILIAIRE'),GetControlText('T_LIBELLE'),VH^.Entree) ;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 26/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.OnClickBZecrimvt(Sender: TObject);
begin
{$IFNDEF PGIIMMO}
{$IFDEF COMPTA}
    ZoomEcritureMvt(GetControlText('T_AUXILIAIRE'),fbAux,'MULMMVTS') ;
{$ENDIF}
{$ENDIF}
end;

{ Fin }

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.OnExitAuxiliaire(Sender: TObject);
var codeauxi, libelle : string;
    moderegl : THValComboBox;
begin

  // GCO - 05/06/2006 - FQ 18155
  // Bizarrement, SetFocus ne fonctionne pas, si "onglet sur +sieurs lignes"
  // dans les paramètres DEVELOPPEUR
  if Trim(GetControlText('T_AUXILIAIRE')) = '' then
  begin
    SetFocusControl('T_AUXILIAIRE');
    Exit;
  end;

  // recup des text
  Codeauxi := BourreEtLEss(GetControlText('T_AUXILIAIRE'), fbAux);
  Libelle  := GetControlText('T_LIBELLE');
  Moderegl := THValComboBox(GetControl('T_MODEREGLE',true));

  if ((codeauxi <> '') and (libelle = '') and (ctxPCL in V_PGI.PGIContexte)) then
  begin
      // libelle := AGLLanceFiche('CP','LIBELLECPTE','','','') ;
      Pages.ActivePageIndex := PGeneral;
      fbForceSaisieLibelle := True;
      SetFocusControl('T_LIBELLE');
      if (moderegl.Text = '') then moderegl.ItemIndex := 0;
  end;

  // reset des text
  SetControlText('T_AUXILIAIRE',codeauxi);
  SetControlText('T_LIBELLE',libelle);
  // il faut aussi mettre à jour le champs  - CA - 05/08/2004 - FQ 13921
  SetField ('T_AUXILIAIRE', codeauxi);
  SetField('T_LIBELLE',libelle);
      // ajout me fiche 20412
  SetField('T_FACTURE', Codeauxi);

end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.OnExitFacture(Sender: TObject);
var
    QLoc : TQuery;
    facture : string;
    SQL     : string;
begin
    if (Mode <> taConsult) then
    begin
        // recup du text
        facture := THEdit(Sender).text;
        // si le tiers payeur n'est pas nul on le complete ....
        if (facture = '') then Exit;
        if (Length(facture) < LgCode) then facture := BourreLaDonc(facture,fbAux);
        // reset du text
        THEdit(Sender).text := facture;
        {JP 12/02/04 : correction d'une violation et de d'erreur de requête}
        SQL := 'SELECT T_FERME FROM TIERS WHERE T_AUXILIAIRE = "' + facture + '" AND ' +
               (THEdit(GetControl('T_FACTURE',true)).plus);
        // on fait la requete pour avoir le tiers facturé corespondant
        QLoc := OpenSql(SQL, True);
        // on verifie son existance
        if (Qloc.RecordCount = 0) then
        begin
            MsgBox.Execute(35,'','');
            THEdit(Sender).SetFocus;
        end;
        // on verifie je sait pas quoi ...
        if QLoc.Fields[0].AsString='X' then
        begin
            MsgBox.Execute(34,'','');
            THEdit(Sender).SetFocus;
        end;
        Ferme(QLoc);
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.OnExitCollectif(Sender: TObject);
//var
//    QLoc : TQuery;
//    collectif : string;
begin
  // GCO - 07/10/2004 - FQ 14637
  if (TFFiche(Ecran).TypeAction <> taConsult) then
  begin
    if Length(GetControlText('T_COLLECTIF')) <> VH^.Cpta[fbGene].Lg then
      SetField('T_COLLECTIF', BourreEtLess( GetControlText('T_COLLECTIF'), fbGene));
  // FIN GCO

    { en fait il semblerais qu'on se foute completement du retour de ExisteH ... donc ....
     // GCO - 07/10/2004 - Auteur de cette modif svp ?

        // recup du text
        collectif := THEdit(Sender).text;
        // si le tiers payeur n'est pas nul on le complete ....
        if (collectif = '') then Exit;
        if (Length(collectif) < VH^.Cpta[fbAux].Lg) then collectif := BourreLaDonc(collectif,fbAux);
        // reset du text
        THEdit(Sender).text := collectif;

        // on fait la requete pour avoir le tiers facturé corespondant
        QLoc := OpenSql(''+facture+'"'+(THValComboBox(GetControl('T_COLLECTIF',true)).plus),True);
        if (Qloc.RecordCount = 0) then
        begin
            MsgBox.Execute(35,'','');
            THEdit(Sender).SetFocus;
        end;
        // on verifie je sait pas quoi ...
        if QLoc.Fields[0].AsString='X' then
        begin
            MsgBox.Execute(34,'','');
            THEdit(Sender).SetFocus;
        end;
        Ferme(QLoc);
    }
  end;
    // if (GetControlText('T_LIBELLE') = '') and (ctxPCL in V_PGI.PGIContexte) then
    // SetControlText('T_LIBELLE',AGLLanceFiche('CP','LIBELLECPTE','','',''));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/02/2003
Modifié le ... : 24/02/2003
Description .. : Format
Mots clefs ... :
*****************************************************************}
procedure TOM_TIERS.OnChangeDebCred(Sender: TObject);
var
    szName : String;
begin
    if (Sender = nil) then exit;
    szName := TControl(Sender).Name + '1';
    if THEdit(Sender).Text='' then
      THNumEdit(GetControl(szName,True)).Value := 0
    else
      THNumEdit(GetControl(szName,True)).Value := StrToFloat(THEdit(Sender).Text );
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/02/2003
Modifié le ... : 24/02/2003
Description .. : Format
Mots clefs ... :
*****************************************************************}
{FP FQ15570
procedure TOM_TIERS.OnValideDebCred(Sender: TObject);
var
    szName : String;
begin
    if (Sender = nil) then exit;

    szName := Copy(TControl(Sender).Name,1,length(TControl(Sender).Name)-1);
    THEdit(GetControl(szName,True)).text := FloatToStr(THNumEdit(Sender).value);
end;}

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/02/2003
Modifié le ... : 24/02/2003
Description .. : Format
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.FormateLesMontants;
var
    mask : string;
begin
    // format des zones de l'exercices (ecran comptable)
    ChangeMask(THNumEdit(GetControl('TSOLCREP',true)),V_PGI.OkdecV,'');
    ChangeMask(THNumEdit(GetControl('TSOLCREE',true)),V_PGI.OkdecV,'');
    ChangeMask(THNumEdit(GetControl('TSOLCRES',true)),V_PGI.OkdecV,'');
    ChangeMask(THNumEdit(GetControl('T_TOTCREP1',true)),V_PGI.OkdecV,'');
    ChangeMask(THNumEdit(GetControl('T_TOTCREE1',true)),V_PGI.OkdecV,'');
    ChangeMask(THNumEdit(GetControl('T_TOTCRES1',true)),V_PGI.OkdecV,'');
    ChangeMask(THNumEdit(GetControl('T_TOTDEBP1',true)),V_PGI.OkdecV,'');
    ChangeMask(THNumEdit(GetControl('T_TOTDEBE1',true)),V_PGI.OkdecV,'');
    ChangeMask(THNumEdit(GetControl('T_TOTDEBS1',true)),V_PGI.OkdecV,'');
    // format des zones de derniere mouvement (ecran comptable)
    ChangeMask(THNumEdit(GetControl('T_DEBITDERNMVT1',true)),V_PGI.OkdecV,'');
    ChangeMask(THNumEdit(GetControl('T_CREDITDERNMVT1',true)),V_PGI.OkdecV,'');
    {FP FQ15570: suppression des composants de type THNumEdit et utilisation de la property DisplayFormat
                 à la place
    // format des zones de credit (ecran commercial)
    ChangeMask(THNumEdit(GetControl('T_CREDITPLAFOND1',true)),V_PGI.OkdecV,'');
    ChangeMask(THNumEdit(GetControl('T_CREDITDEMANDE1',true)),V_PGI.OkdecV,'');
    ChangeMask(THNumEdit(GetControl('T_CREDITACCORDE1',true)),V_PGI.OkdecV,'');}

    mask := StrfMask(V_PGI.OkdecV,'',True);

    {b FP FQ15570}
    SetControlProperty('T_CREDITPLAFOND','DISPLAYFORMAT',mask);
    SetControlProperty('T_CREDITDEMANDE','DISPLAYFORMAT',mask);
    SetControlProperty('T_CREDITACCORDE','DISPLAYFORMAT',mask);
    {e FP FQ15570}
    SetControlProperty('T_TOTALDEBIT','EDITMASK',mask);
    SetControlProperty('T_TOTALCREDIT','EDITMASK',mask);
    SetControlProperty('T_TOTDEBANO','EDITMASK',mask);
    SetControlProperty('T_TOTCREANO','EDITMASK',mask);
    SetControlProperty('T_TOTDEBANON1','EDITMASK',mask);
    SetControlProperty('T_TOTCREANON1','EDITMASK',mask);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_TIERS.AssignFcts;
var
    valeur : THEdit;
    //FP FQ15570 saisie : THNumEdit;
begin
    // zones de derniere mouvement (ecran comptable)
    if EstTablePartagee('TIERS') then
      begin
      // MODE MULTISOC : si la table est partagée, on cache les infos de dernier mouvement
      SetControlVisible( 'HGBDERNMOUV', False ) ;
      end
    else
      begin
      valeur := THEdit(GetControl('T_DEBITDERNMVT',true));      if (assigned(valeur)) then valeur.OnChange := OnChangeDebCred;
      valeur := THEdit(GetControl('T_CREDITDERNMVT',true));     if (assigned(valeur)) then valeur.OnChange := OnChangeDebCred;
      end ;
{FP FQ15570
    // zones de credit (ecran commercial)
    valeur := THEdit(GetControl('T_CREDITDEMANDE',true));     if (assigned(valeur)) then valeur.OnChange := OnChangeDebCred;
    valeur := THEdit(GetControl('T_CREDITACCORDE',true));     if (assigned(valeur)) then valeur.OnChange := OnChangeDebCred;
    valeur := THEdit(GetControl('T_CREDITPLAFOND',true));     if (assigned(valeur)) then valeur.OnChange := OnChangeDebCred;
    // set de recup des valeur sur les zone de chiffrage visible
    saisie := THNumEdit(GetControl('T_CREDITDEMANDE1',true)); if (assigned(saisie)) then saisie.OnExit := OnValideDebCred;
    saisie := THNumEdit(GetControl('T_CREDITACCORDE1',true)); if (assigned(saisie)) then saisie.OnExit := OnValideDebCred;
    saisie := THNumEdit(GetControl('T_CREDITPLAFOND1',true)); if (assigned(saisie)) then saisie.OnExit := OnValideDebCred;
}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/02/2003
Modifié le ... : 24/02/2003
Description .. : Enable/disable les groupe box et leur contenu
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.EnabledGB(GBname : string);
var
    l : Integer ;
    GB : TGroupBox;
begin
    GB := TGroupBox(GetControl(GBname,true));
    if (assigned(GB)) then for l := 0 to GB.ControlCount-1 do GB.Controls[l].Enabled := GB.Enabled;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/02/2003
Modifié le ... : 24/02/2003
Description .. : enable et disable certain controle selon certain critere de
Suite ........ : nature de compte !
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.AutoriseReleveFacture;
begin
    // caché les controls ki vos bien selon la nature !
    if ((nature.Value='CLI') Or (nature.Value='AUD')) then
    begin
        if (GetCheckBoxState('T_LETTRABLE') = cbChecked) then
        begin
            if (Not GetControlEnabled('T_RELANCEREGLEMENT')) then
            begin
                SetControlEnabled('T_RELANCEREGLEMENT',True);
                SetControlEnabled('T_RELANCETRAITE',True);
                SetControlEnabled('TT_RELEVEFACTURE',True);
            end ;
        end
        else
        begin
            SetControlEnabled('T_RELANCEREGLEMENT',False);
            SetControlEnabled('T_RELANCETRAITE',False);
            SetControlEnabled('TT_RELEVEFACTURE',False);
            SetControlChecked('T_RELEVEFACTURE',False);
            SetControlText('T_JOURRELEVE','0');
            SetControlCaption('T_FREQRELEVE','');
        end ;
        // Assurance crédit
        SetControlEnabled('TT_ASSURANCE',True);
    end
    else
    begin
        SetControlEnabled('T_RELANCEREGLEMENT',False);
        SetControlEnabled('T_RELANCETRAITE',False);
        SetControlEnabled('TT_RELEVEFACTURE',False);
        // Assurance crédit
        SetControlEnabled('TT_ASSURANCE',False);
    end ;
    EnabledGB('TT_ASSURANCE');
    EnabledGB('TT_RELEVEFACTURE');
    SetControlEnabled('TT_RELANCEREGLEMENT',GetControlEnabled('T_RELANCEREGLEMENT'));
    SetControlEnabled('TT_RELANCETRAITE',GetControlEnabled('T_RELANCETRAITE'));
    SetControlEnabled('SOLCPTA',FALSE) ; //XMG 19/01/04
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/02/2003
Modifié le ... : 24/02/2003
Description .. : desactivationn de certain control dans le cas des salarié
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.PourSalarie ;
begin
    if (Nature.Value = 'SAL') then
    begin
        // cahngement de page
        if (Pages.ActivePageIndex = PCommercial) then Pages.ActivePageIndex := PComptable;
        // hide de la page
        Pages.pages[3].TabVisible := False ;
        // hide des controls
        SetControlVisible('T_COUTHORAIRE',True);
        SetControlVisible('TT_COUTHORAIRE',True);
        SetControlVisible('T_SOUMISTPF',False);
        SetControlVisible('HLRegimeTVA',False);
        SetControlVisible('T_REGIMETVA',False);
        SetControlVisible('TT_TVAENCAISSEMENT',False);
        SetControlVisible('T_TVAENCAISSEMENT',False);
        SetControlVisible('TT_NIF',False);
        SetControlVisible('T_NIF',False);
        SetControlVisible('TT_SIRET',False);
        SetControlVisible('T_SIRET',False);
        SetControlVisible('TT_APE',False);
        SetControlVisible('T_APE',False);
        // desactivation des controls
        SetControlEnabled('T_PAYEUR',False);
        SetControlEnabled('TT_PAYEUR',False);
        SetControlEnabled('T_DEBRAYEPAYEUR',False);
        SetControlEnabled('BContact',False);
        SetControlEnabled('TT_TRESORERIE',False);
    end
    else
    begin
        // unhide de la page
        Pages.pages[3].TabVisible := True ;
        // unhide des controls
        SetControlVisible('T_COUTHORAIRE',False);
        SetControlVisible('TT_COUTHORAIRE',False);
        If (Not EstSerie(S3)) Then SetControlVisible('T_SOUMISTPF',True);
        SetControlVisible('HLRegimeTVA',True);
        SetControlVisible('T_REGIMETVA',True);
        SetControlVisible('TT_TVAENCAISSEMENT',VH^.PaysLocalisation<>CODEISOES) ; //XVI 24/02/2055
        SetControlVisible('T_TVAENCAISSEMENT',VH^.PaysLocalisation<>CODEISOES) ;
        SetControlVisible('TT_NIF',True);
        SetControlVisible('T_NIF',True);
        SetControlVisible('TT_SIRET',True);
        SetControlVisible('T_SIRET',True);
        SetControlVisible('TT_APE',True);
        SetControlVisible('T_APE',True);
        // activation des controls
        SetControlEnabled('TT_TRESORERIE',True);
        SetControlEnabled('T_PAYEUR',True);
        SetControlEnabled('TT_PAYEUR',True);
        SetControlEnabled('T_DEBRAYEPAYEUR',True);
        SetControlEnabled('BContact',True);
    end;
    if VH^.PaysLocalisation=CODEISOES then
    begin
      SetControlVisible('TT_NIF',FALSE);
      SetControlVisible('T_NIF',FALSE);
      SetControlVisible('TT_SIRET',TRUE);
      SetControlVisible('T_SIRET',TRUE);
    End ;  //XVI 24/02/2005
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.InitModif;
begin
    MNat := nature.text;
    ML := GetControlText('T_LETTRABLE');
    Dev := GetControlText('T_DEVISE');
    MultiDev := GetControlText('T_MULTIDEVISE');
    RTVA:= GetControlText('T_REGIMETVA');
    TVAE:= GetControlText('T_TVAENCAISSEMENT');
    // CA - 02/11/2006 - Suite demande GPAO, on permet de saisir une devise par défaut en multi devises
  //  if (MultiDev = 'X') then SetControlEnabled('T_DEVISE',False)
  //  else SetControlEnabled('T_DEVISE',True);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.OnEstEnSaisieCompta;
begin
    // FQ 16383 - Prise en compte mode consultation
    if ((Not EnSaisieCpta) and (Mode <> taConsult)) then Exit;
    nature.enabled := False;
    SetControlEnabled('T_LETTRABLE',False);
    SetControlEnabled('BInsert',False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Function TOM_TIERS.OkConfidentiel : Boolean;
begin
    Result:=False ;
    if V_PGI.Confidentiel='0' then
    begin
        if (GetCheckBoxState('TConfidentiel') = cbChecked) then
        begin
            MessageAlerte(MsgBox.Mess[30]);
            close;
            Exit;
        end;
    end;
    SetControlVisible('TCONFIDENTIEL',(V_PGI.Confidentiel = '1'));
    SetControlEnabled('TCONFIDENTIEL',V_PGI.Superviseur);
    Result := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.NatureChgAutorise ;
begin
    // FQ 16383 - Prise en compte mode consultation
    if (Mode <> taConsult) then
      nature.enabled := True;
    if (EnSaisieCpta) then Exit;
    if (not AvecMvt)  then Exit;
    nature.enabled := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.LettrageChgAutorise ;
begin
    // FQ 16383 - Prise en compte mode consultation
    if (Mode <> taConsult) then SetControlEnabled('T_LETTRABLE',True);
    if (EnSaisieCpta) then Exit;
    if (not AvecMvt)  then Exit;
    SetcontrolEnabled('T_LETTRABLE',(GetCheckBoxState('T_LETTRABLE') = cbUnchecked));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.PayeurEna ;
begin
    if (GetCheckBoxState('T_ISPAYEUR') = cbChecked) then
    begin
        SetcontrolEnabled('T_PAYEUR',False);
        SetcontrolEnabled('TT_PAYEUR',False);
        SetcontrolEnabled('T_DEBRAYEPAYEUR',False);
        SetcontrolEnabled('T_AVOIRRBT',False);

        SetControlChecked('T_AVOIRRBT',False);
    end
    else
    begin
        SetcontrolEnabled('T_PAYEUR',True);
        SetcontrolEnabled('TT_PAYEUR',True);
        SetcontrolEnabled('T_AVOIRRBT',True);
        SetcontrolEnabled('T_DEBRAYEPAYEUR',True);
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Function TOM_TIERS.IsMouvemente : Boolean;
begin
    if ((Mode = taConsult) or (DS.State = dsInsert)) then
    begin
        Result := False;
        Exit;
    end;
    {JP 18/08/05 : FQ 15548 : L'affichage de la fiche est longue si la table est grosse.
                   Je ne suis pas sûr que cela résolve le problème, mais cela ne peut pas
                   faire de mal de supprimer le Exist
    Result := ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="'+GetControlText('T_AUXILIAIRE')+'" '+'AND EXISTS(SELECT E_AUXILIAIRE FROM ECRITURE WHERE E_AUXILIAIRE="'+GetControlText('T_AUXILIAIRE')+'" )') ;}
    Result := ExisteSQL('SELECT E_AUXILIAIRE FROM ECRITURE WHERE E_AUXILIAIRE = "' + GetControlText('T_AUXILIAIRE') + '"') ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 02/07/2007
Modifié le ... :   /  /    
Description .. : Permet de savoir si les mouvements du compte sont sur 
Suite ........ : plusieurs devise ou non.
Suite ........ : FQ 19238
Mots clefs ... : 
*****************************************************************}
Function TOM_TIERS.IsMultiDevise : Boolean;
var
  SQL : string;
  Q   : TQuery;
begin
  result := false;
  SQL := 'SELECT COUNT(DISTINCT(E_DEVISE)) FROM ECRITURE WHERE E_AUXILIAIRE = "' + GetControlText('T_AUXILIAIRE') + '"';
  Q := OpenSQL(SQL,true);
  try
     if not Q.Eof then
     begin
        if Q.Fields[0].AsInteger > 1 then
           Result := true
        else
           Result := false;
     end;
  finally
     Ferme(Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Function TOM_TIERS.CodeValide : Boolean;
var
    StAux : String;
begin
    Result := false ;

    if (DS.state = dsinsert) then
    begin
        // recup du num de compte
        StAux := GetControlText('T_AUXILIAIRE');
        // si vide
        if (StAux = '') then
        begin
            Pages.ActivePageIndex := PComptable;
            SetFocusControl('T_AUXILIAIRE');
            MsgBox.Execute(2,'','');
            Exit;
        end;
        // si ???
        if (ExisteCarInter(StAux)) then
        begin
            Pages.ActivePageIndex := PComptable;
            SetFocusControl('T_AUXILIAIRE');
            MsgBox.Execute(41,'','');
            Exit;
        end;
        // test de la non existance du compte chez les generaux
        if Presence('TIERS','T_AUXILIAIRE',StAux) then
        begin
            Pages.ActivePageIndex := PComptable;
            SetFocusControl('T_AUXILIAIRE');
            MsgBox.Execute(4,'','');
            Exit;
        end;
// BPY le 11/07/2003
//        if Presence('TIERS','T_TIERS',BourreEtLess(stAux,fbAux)) then
//        begin
//          Pages.ActivePageIndex:=PComptable;
//          SetFocusControl('T_AUXILIAIRE');
//          MsgBox.Execute(42,'','');
//          Exit;
//        end;
    end;

    Result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 26/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Function TOM_TIERS.CodeCollectifValide : Boolean ;
var
    Q : TQuery ;
    NatAux,NatGen : String3;
begin
    Result := False;

    // si vide
    if (GetControlText('T_COLLECTIF') = '') then
    begin
        Pages.ActivePageIndex := PComptable;
        SetFocusControl('T_COLLECTIF');
        MsgBox.execute(16,'','');
        Exit;
    end;
    // test de l'existance du compte chez les compte generaux
    if (not Presence('GENERAUX','G_GENERAL',GetControlText('T_COLLECTIF'))) then
    begin
        Pages.ActivePageIndex := PComptable;
        SetFocusControl('T_COLLECTIF');
        MsgBox.execute(17,'','');
        Exit;
    end
    else
    begin
        // test la coherence du couple de compte nature auxiliaire / nature generaux
        NatAux := nature.value;
        Q := OpenSql('Select G_NATUREGENE From GENERAUX Where G_GENERAL="'+GetControlText('T_COLLECTIF')+'"',True) ;
        NatGen := Q.Fields[0].AsString ;
        Ferme(Q) ;

        if not NATUREGENAUXOK(NatGen, NatAux) then  // FQ 12769
        begin
            Pages.ActivePageIndex := PComptable;
            SetFocusControl('T_COLLECTIF');
            MsgBox.execute(18,'','');
            Exit;
        end;
    end;

    Result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.InitModifEnSerie(StModif : string);
var
    St,Champ, Valeur : string;
    i : integer;
    B : TBitBtn;
begin
    // change le mode ...
    if (DS.State = dsBrowse) then DS.Edit;
    // recupere les parametre de changement et affecte
    While (StModif <> '') do
    begin
        St := ReadTokenSt(StModif);
        i := Pos('=',St); if i>0 then Champ := Trim(Copy(St,1,i-1));
        i := Pos('"',St); if i>0 then St := Trim(Copy(St,i+1,Length(St)));
        i := Pos('"',St); if i>0 then Valeur := Trim(Copy(St,1,i-1));
        AffecteLe(Champ,Valeur);
    end;
    // disable tt les bouton
    For i:=0 to TFFiche(ecran).ControlCount-1 do
    begin
        if (TFFiche(ecran).Controls[i] is TBitBtn) then
        begin
            B := TBitBtn(TFFiche(ecran).Controls[i]);
            if ((B.Name <> 'BValider') and (B.Name <> 'BFerme') and (B.Name <> 'BAide')) then B.Enabled := false;
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.AffecteLe(Champ, Valeur : string) ;
var
    C : TControl;
begin
    C := GetControl(Champ);
    if (C=nil) then exit;

{$IFDEF EAGLCLIENT}
    if ((C is TCheckBox) or (C is THValComboBox) or (C is THEdit)) then
{$ELSE}
    if ((C is TDBCheckBox) or (C is THDBValComboBox) or (C is THDBEdit) or (C is THValComboBox)) then // 12399
{$ENDIF}
    begin
        //SetControlText(Champ,Valeur);
{b fb 17/11/2005 FQ14934}
        SetField ( Champ, Valeur );
        if (Champ='YTC_ACCELERATEUR') then
          SetControlChecked(Champ, Valeur = 'X' )
        else if (Champ='YTC_SCHEMAGEN') then
          SetControlText(Champ,Valeur)
        else
          SetField ( Champ, Valeur );
{e fb 17/11/2005 FQ14934}
        THEdit(C).Font.Color := clRed;
    end
{$IFDEF EAGLCLIENT}
    else if (C is THSpinEdit) then
{$ELSE}
    else if (C is THDBSpinEdit) then
{$ENDIF}
    begin
        SetControlText(Champ,Valeur);
        THSpinEdit(C).Font.Color := clRed;
    end
     else
      if ( C is THEdit ) then
       begin
         SetControlText(Champ,Valeur);
         THEdit(C).Font.Color := clRed;
       end
        else
         if ( C is TCheckBox ) then
          begin
           SetControlChecked(Champ, Valeur = 'X' ) ;
           THEdit(C).Font.Color := clRed;
          end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.DelRibContact ;
var
    auxiliaire : string;
begin
    // get bour et reset du compte auxiliair
    auxiliaire := GetControlText('T_AUXILIAIRE');
    if (BourreEtLess(auxiliaire,fbAux) = '') then Exit ;
    // delete ...
    ExecuteSQL('Delete from CONTACT Where C_AUXILIAIRE="' + auxiliaire + '"') ;
    ExecuteSQL('Delete from RIB Where R_AUXILIAIRE="' + auxiliaire + '"') ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOM_TIERS.ActivationControl;
var
    lStAuxiliaire : string;
begin
{$IFNDEF PGIIMMO}
    if (not (CtxPcl in V_Pgi.PGIContexte)) then Exit;

    lStAuxiliaire := Trim(GetControlText('T_Auxiliaire'));

    SetControlEnabled('T_Collectif',(lStAuxiliaire <> ''));
    SetControlProperty('T_Collectif','Color',IIF(GetControlEnabled('T_Collectif'),ClWindow,ClBtnFace));

    SetControlEnabled('T_Lettrable',(lStAuxiliaire <> ''));
    SetControlEnabled('T_MultiDevise',(lStAuxiliaire <> ''));
    SetControlEnabled('TConfidentiel',(GetControlVisible('TConfidentiel')) and (lStAuxiliaire <> ''));

    SetControlEnabled('T_Devise',(lStAuxiliaire <> ''));
    SetControlProperty('T_Devise','Color',IIF(GetControlEnabled('T_Devise'),ClWindow,ClBtnFace));

    SetControlEnabled('T_SoumisTpf',(lStAuxiliaire <> ''));

    SetControlEnabled('T_RegimeTva',(lStAuxiliaire <> ''));
    SetControlProperty('T_RegimeTva','Color',IIF(GetControlEnabled('T_RegimeTva'),ClWindow,ClBtnFace));

    SetControlEnabled('T_TvaEncaissement',(lStAuxiliaire <> ''));
    SetControlProperty('T_TvaEncaissement','Color',IIF(GetControlEnabled('T_TvaEncaissement'),ClWindow,ClBtnFace));

    SetControlEnabled('T_Corresp1',(lStAuxiliaire <> ''));
    SetControlProperty('T_Corresp1','Color',IIF(GetControlEnabled('T_Corresp1.Enabled'),ClWindow,ClBtnFace));

    SetControlEnabled('T_Corresp2',(lStAuxiliaire <> ''));
    SetControlProperty('T_Corresp2','Color',IIF(GetControlEnabled('T_Corresp2'),ClWindow,ClBtnFace));
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 26/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Function TOM_TIERS.VerifCoherenceTL : Boolean;
var
    i : Integer;
    C : THEdit;
    Alerte : Boolean;
begin
    Result:=True ;

    if (not Pages.Pages[ZL].TabVisible) then Exit;

    Alerte := False ;

    for i:=0 to 9 do
    begin
        C := THEdit(GetControl('T_TABLE'+IntToStr(i),true));
        if (Not C.Enabled) then Continue;
        if (C.Text = '')   then Continue;
        if (not ExisteSql('SELECT NT_NATURE, NT_LIBELLE From NATCPTE WHERE NT_NATURE Like "'+C.Text+'" AND NT_TYPECPTE="T0'+IntToStr(i)+'"')) then Alerte := True;
    END ;
    if (Alerte) then if (MsgBox.Execute(39,'','') <> mrYes) then Result := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 26/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Function TOM_TIERS.LettrageChanger : Boolean ;
Var
    AvecMvtTemp,AeteModifier : Boolean ;
BEGIN
    Result := False;
    AvecMvtTemp := False;

    if ((ML='-') and (GetCheckBoxState('T_LETTRABLE') = cbUnChecked) or ((ML='X') and (GetCheckBoxState('T_LETTRABLE') = cbChecked))) then AeteModifier := False
    else AeteModifier := True ;

    if ((AeteModifier) And (Not AvecMvt)) then AvecMvtTemp := IsMouvemente;

    if (AvecMvtTemp) then
    begin
        if ((ML='X') and (GetCheckBoxState('T_LETTRABLE') = cbUnChecked) ) then
        begin
            MsgBox.Execute(36,'','') ;
            AvecMvt := True ;
            SetControlText('T_LETTRABLE',ML);
            Exit;
        end;
        AvecMvt := True ;
    end;

    if (Not AvecMvt) then Exit ;
    if (DS.State = dsInsert) then Exit ;
    Result := ((ML='-') and (GetCheckBoxState('T_LETTRABLE') = cbChecked));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 20/12/2007
Modifié le ... :   /  /
Description .. : Test de modif des param TVA
Mots clefs ... : FQ19946
*****************************************************************}
Function TOM_TIERS.TVAChanger : Boolean ;
Var
    AeteModifie : Boolean ;
BEGIN
    Result := False;
    AeteModifie := False;

    if (RTVA) <> (GetControlText('T_REGIMETVA')) then
      AeteModifie := True;

    if (TVAE) <> (GetControlText('T_TVAENCAISSEMENT')) then
      AeteModifie := True;

    if (DS.State = dsInsert) then Exit ;
    Result := AeteModifie;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/02/2003
Modifié le ... : 25/02/2003
Description .. : init des message de la MessageBox
Mots clefs ... :
*****************************************************************}
procedure TOM_TIERS.InitMsgBox;
begin
    MsgBox.Mess.Add('0;Comptes auxiliaires;Voulez-vous enregistrer les modifications?;Q;YNC;Y;C;');
    MsgBox.Mess.Add('1;Comptes auxiliaires;Confirmez-vous la suppression de l'#39'enregistrement?;Q;YNC;N;C;');
    MsgBox.Mess.Add('2;Comptes auxiliaires;Vous devez renseigner un code.;W;O;O;O;');
    MsgBox.Mess.Add('3;Comptes auxiliaires;Vous devez renseigner un libellé.;W;O;O;O;');
    MsgBox.Mess.Add('4;Comptes auxiliaires;Le code que vous avez saisi existe déjà. Vous devez le modifier.;W;O;O;O;');
    MsgBox.Mess.Add('5;Comptes auxiliaires;Vous ne pouvez pas supprimer ce compte auxiliaire : il est un modèle.;W;O;O;O;');
    MsgBox.Mess.Add('6;Comptes auxiliaires;Vous ne pouvez pas supprimer ce compte auxiliaire : il est référencé dans un catalogue.;W;O;O;O;');
    MsgBox.Mess.Add('7;Comptes auxiliaires;Vous ne pouvez pas supprimer ce compte auxiliaire : il est référencé pour des achats ou des ventes.;W;O;O;O;');
    MsgBox.Mess.Add('8;Comptes auxiliaires;Vous ne pouvez pas supprimer ce compte auxiliaire : il est dans une ligne de piece interne.;W;O;O;O;');
    MsgBox.Mess.Add('9;Comptes auxiliaires;Vous ne pouvez pas supprimer ce compte auxiliaire : il est référencé dans une piece de gestion commerciale.;W;O;O;O;');
    MsgBox.Mess.Add('10;Comptes auxiliaires;Vous ne pouvez pas supprimer ce compte auxiliaire : il est dans une piece interne.;W;O;O;O;');
    MsgBox.Mess.Add('11;Comptes auxiliaires;Vous ne pouvez pas supprimer ce compte auxiliaire : il est référencé dans un RIB.;W;O;O;O;');
    MsgBox.Mess.Add('12;Comptes auxiliaires;Vous ne pouvez pas supprimer ce compte auxiliaire : il est auxiliaire d'#39'attente pour la société.;W;O;O;O;');
    MsgBox.Mess.Add('13;Comptes auxiliaires;Vous ne pouvez pas supprimer ce compte auxiliaire : il est référencé dans une section analytique.;W;O;O;O;');
    MsgBox.Mess.Add('14;Comptes auxiliaires;Vous ne pouvez pas supprimer ce compte auxiliaire : il est un client, représentant, fournisseur ou acheteur.;W;O;O;O;');
    MsgBox.Mess.Add('15;Comptes auxiliaires;Vous ne pouvez pas supprimer ce compte auxiliaire : il pointe vers un auxiliaire interne.;W;O;O;O;');
    MsgBox.Mess.Add('16;Comptes auxiliaires;Vous devez renseigner le compte collectif.;W;O;O;O;');
    MsgBox.Mess.Add('17;Comptes auxiliaires;Le compte collectif que vous avez renseigné n'#39'est pas un compte général.;W;O;O;O;');
    MsgBox.Mess.Add('18;Comptes auxiliaires;La nature du compte collectif n'#39'est pas en correspondance avec la nature du compte auxiliaire.;W;O;O;O;');
    MsgBox.Mess.Add('19;Comptes auxiliaires;Ce compte auxiliaire n'#39'a pas de RIB associé.;E;O;O;O;');
    MsgBox.Mess.Add('20;Comptes auxiliaires;Ce compte auxiliaire n'#39'a pas de RIB associé. Voulez-vous le créer ?;Q;YNC;Y;C;');
    MsgBox.Mess.Add('21;Comptes auxiliaires;Ce compte auxiliaire n'#39'a pas de contact associé.;E;O;O;O;');
    MsgBox.Mess.Add('22;Comptes auxiliaires;Ce compte auxiliaire n'#39'a pas de contact associé. Voulez-vous le créer ?;Q;YNC;Y;C;');
    MsgBox.Mess.Add('23;Comptes auxiliaires;La caractéristique "lettrable" a été modifiée. Désirez-vous continuer?;Q;YNC;N;C;');
    MsgBox.Mess.Add('24;Comptes auxiliaires;La modification de la caractéristique "lettrable" va entraîner des modifications d'#39'écritures. Vous devez sélectionner un mode de paiement.;W;O;O;O;');
    MsgBox.Mess.Add('Choix d'#39'un mode de paiement');
    MsgBox.Mess.Add('26;Comptes auxiliaires;Le traitement ne peut pas être effectué : le mode de paiement n'#39'est pas renseigné.;W;O;O;O;');
    MsgBox.Mess.Add('L'#39'enregistrement est inaccessible');
    MsgBox.Mess.Add('28;Comptes auxiliaires;Vous ne pouvez pas supprimer ce compte auxiliaire : il est mouvementé.;W;O;O;O;');
    MsgBox.Mess.Add('29;Comptes auxiliaires;Vous ne pouvez pas supprimer ce compte auxiliaire : il est référencé dans un guide.;W;O;O;O;');
    MsgBox.Mess.Add('L'#39'accès à ce tiers est interdit.');
    MsgBox.Mess.Add('31;Comptes auxiliaires;Vous devez renseigner un mode de règlement.;W;O;O;O;');
    MsgBox.Mess.Add('32;Comptes auxiliaires;Vous devez renseigner un régime fiscal.;W;O;O;O;');
    MsgBox.Mess.Add('Aucune relance');
    MsgBox.Mess.Add('34;Comptes auxiliaires;Attention : le tiers choisi est fermé.;W;O;O;O;');
    MsgBox.Mess.Add('35;Comptes auxiliaires;Le code que vous avez saisi n'#39'existe pas. Vous devez le modifier.;W;O;O;O;');
    MsgBox.Mess.Add('36;Comptes auxiliaires;Vous ne pouvez pas modifier la caractéristique "lettrable" : un utilisateur vient de saisir une pièce avec cet auxiliaire.;W;O;O;O;');
    MsgBox.Mess.Add('37;Comptes auxiliaires;Le compte de correspondance n°1 n'#39'existe pas. Vous devez le modifier.;W;O;O;O;');
    MsgBox.Mess.Add('38;Comptes auxiliaires;Le compte de correspondance n°2 n'#39'existe pas. Vous devez le modifier.;W;O;O;O;');
    MsgBox.Mess.Add('39;Comptes auxiliaires;ATTENTION. Il y a des codes de tables libres qui n'#39'existent pas. Désirez vous continuer?;Q;YNC;N;C;');
    MsgBox.Mess.Add('40;Comptes auxiliaires;Le tiers payeur ne peut pas être lui-même.;W;O;O;O;');
    MsgBox.Mess.Add('41;Comptes auxiliaires;Le code auxiliaire comporte des caractères interdits.;W;O;O;O;');
    MsgBox.Mess.Add('42;Comptes auxiliaires;Le code que vous avez saisi existe déjà en gestion interne. Vous devez le modifier.;W;O;O;O;');
    MsgBox.Mess.Add('43;Comptes auxiliaires;La clé NIF/CIF est erronnée. Souhaitez-vous la recalculer ?;Q;YN;Y;C;') ; //XMG 19/01/04
    MsgBox.Mess.Add('44;Comptes auxiliaires;Le compte d''accélération que vous avez renseigné n'#39'est pas un compte général.;W;O;O;O;');
    MsgBox.Mess.Add('45;Comptes auxiliaires;Le compte est référencé en tant que payeur d''autre(s) tiers '#13'(');
    MsgBox.Mess.Add(').;W;O;O;O;');
    MsgBox.Mess.Add('');

    MsgBox.Mess.Add('48;Comptes auxiliaires;Opération impossible : Vous n''avez pas renseigné le code identifiant du compte auxiliaire.;W;O;O;O;');
    MsgBox.Mess.Add('49;Comptes auxiliaires;Confirmez-vous la duplication du compte sélectionné ?;Q;YN;N;N;');
    MsgBox.Mess.Add('50;Choix d''un compte auxiliaire à dupliquer');
    MsgBox.Mess.Add('51;Tiers payeur;Ce compte n''existe pas.;W;O;O;O;');
    MsgBox.Mess.Add('52;Comptes auxiliaires;Attention : La modification du paramètrage de TVA n''est pas répercutée sur les écritures existantes.;W;O;O;O;');
end;

// Création d'un enregistrement vierge dans la table PROSPECTS et TIERSCOMPL si on a
// respectivement le contexte ctxGRC et ctxGescom
procedure TOM_TIERS.CreateEnregGescomGRC;
var
  T : Tob;
begin
  // Tables PROSPECTS
  if (ctxGRC in V_PGI.PGIContexte) AND (GetField('T_NATUREAUXI') = 'CLI') then begin
    T := Tob.Create('PROSPECTS', nil, -1);
    T.InitValeurs;
    T.PutValue('RPR_AUXILIAIRE', GetControlText('T_AUXILIAIRE'));
    T.InsertDB(nil);
    T.Free;
  end;

  // Table TIERSCOMPL
//  if (ctxGescom in V_PGI.PGIContexte) then begin
//    CreationTIERSCOMPL ;
//  end;
end;

procedure TOM_TIERS.CreationTIERSCOMPL ;
var
  T : Tob;
  szSQL : String;

  /////////////////////////////////////////////////////////////////////////////
  procedure UpdateChamps(szListeChamps : String);
  var
    szChamp, szValue : String;
    C : TControl;
    Typ : TOpeType;
  begin
    while szListeChamps <> '' do begin
      szChamp := ReadTokenST(szListeChamps);

      C := GetControl(szChamp);
      if (C = nil) then continue;

      Typ :=  otString;
      if C is THEdit then Typ := THEdit(C).OpeType;

      szValue := GetControlText(szChamp);
      case Typ of
        otString: T.PutValue(szChamp, szValue);
        otReel  : if (szValue<>'') then T.PutValue(szChamp, StrToFloat(szValue));
        otDate  : if (szValue<>'') and IsValidDate(szValue) then T.PutValue(szChamp, StrToDate(szValue));
      end;
    end;
  end;

  /////////////////////////////////////////////////////////////////////////////
  procedure UpdateCombos(szListeChamps : String);
  var
    szChamp : String;
  begin
    while szListeChamps <> '' do begin
      szChamp := ReadTokenST(szListeChamps);
      if (GetControl(szChamp) = nil) then continue;
      T.PutValue(szChamp, IIF(GetControlEnabled(szChamp),GetControlText(szChamp),''));
    end;
  end;

  /////////////////////////////////////////////////////////////////////////////
  procedure UpdateCheckbox(szListeChamps : String);
  var
    szChamp : String;
  begin
    while szListeChamps <> '' do begin
      szChamp := ReadTokenST(szListeChamps);
      if (GetControl(szChamp) = nil) then continue;
      if GetCheckBoxState(szChamp) = cbChecked then T.PutValue(szChamp, 'X')
                                               else T.PutValue(szChamp, '-');
    end;
  end;

  /////////////////////////////////////////////////////////////////////////////
  function GetUpdate(szListeChamps : String) : String;
  var
    szChamp, szValue : String;
    C : TControl;
    bDoSQL : Boolean;
    Typ : TOpeType;
  begin
    Result := ', ';
    while szListeChamps <> '' do begin
      szChamp := ReadTokenST(szListeChamps);

      C := GetControl(szChamp);
      if (C = nil) then continue;

      bDoSQL := False;
      Typ :=  otString;
      if C is TCheckBox then bDoSQL := (TCheckBox(C).Font.Color = clRed) else
      if C is THValCombobox then begin
        bDoSQL := (THValCombobox(C).Font.Color = clRed);
        end else
      if C is THEdit then begin
        bDoSQL := (THEdit(C).Font.Color = clRed);
        Typ := THEdit(C).OpeType;
      end;
      if not bDoSQL then continue;

      Result := Result + szChamp+'=';
      szValue := T.GetValue(szChamp);
      case Typ of
        otString: Result := Result + '"'+szValue+'", ';
        otReel  : Result := Result + szValue+', ';
        otDate  : Result := Result + '"'+USDateTime(StrToDate(szValue))+'", ';
      end;
    end;
    // Supprime la dernière virgule
    if (Result<>'') then Delete(Result, Length(Result)-1, 1);
  end;

begin
// BPY le 25/08/2004 => fiche n° 14269 : pb lors de l'enreg des DASS2
   T := Tob.Create('TIERSCOMPL', nil, -1);
   T.InitValeurs;
   T.PutValue('YTC_AUXILIAIRE', GetControlText('T_AUXILIAIRE'));
   T.PutValue('YTC_TIERS', GetField('T_TIERS'));

   // 12399
   UpdateCheckbox('YTC_ACCELERATEUR;YTC_DAS2');
   UpdateChamps('YTC_SCHEMAGEN;'+
                'YTC_PROFESSION;'+
                'YTC_TABLELIBRETIERS1;YTC_TABLELIBRETIERS2;YTC_TABLELIBRETIERS3;YTC_TABLELIBRETIERS4;YTC_TABLELIBRETIERS5;YTC_TABLELIBRETIERS6;YTC_TABLELIBRETIERS7;YTC_TABLELIBRETIERS8;YTC_TABLELIBRETIERS9;YTC_TABLELIBRETIERSA;'+
                'YTC_VALLIBRE1;YTC_VALLIBRE2;YTC_VALLIBRE3;'+
                'YTC_DATELIBRE1;YTC_DATELIBRE2;YTC_DATELIBRE3;'+
                'YTC_TABLELIBREFOU1;YTC_TABLELIBREFOU2;YTC_TABLELIBREFOU3;'+
                'YTC_VALLIBREFOU1;YTC_VALLIBREFOU2;YTC_VALLIBREFOU3;'+
                'YTC_DATELIBREFOU1;YTC_DATELIBREFOU2;YTC_DATELIBREFOU3;'+
                'YTC_TEXTELIBRE1;YTC_TEXTELIBRE2;YTC_TEXTELIBRE3;'+
                'YTC_BOOLLIBRE1;YTC_BOOLLIBRE2;YTC_BOOLLIBRE3');
   UpdateCombos('YTC_REMUNERATION;YTC_INDEMNITE;YTC_AVANTAGE;YTC_FAMREG;YTC_RETENUESOURCE');

   if ExisteSql('SELECT YTC_AUXILIAIRE FROM TIERSCOMPL WHERE YTC_AUXILIAIRE="'+GetControlText('T_AUXILIAIRE')+'" ' ) then begin
     szSQL := GetUpdate('YTC_FAMREG;'+
                        'YTC_TABLELIBRETIERS1;YTC_TABLELIBRETIERS2;YTC_TABLELIBRETIERS3;YTC_TABLELIBRETIERS4;YTC_TABLELIBRETIERS5;YTC_TABLELIBRETIERS6;YTC_TABLELIBRETIERS7;YTC_TABLELIBRETIERS8;YTC_TABLELIBRETIERS9;YTC_TABLELIBRETIERSA;'+
                        'YTC_VALLIBRE1;YTC_VALLIBRE2;YTC_VALLIBRE3;'+
                        'YTC_DATELIBRE1;YTC_DATELIBRE2;YTC_DATELIBRE3;'+
                        'YTC_TABLELIBREFOU1;YTC_TABLELIBREFOU2;YTC_TABLELIBREFOU3;'+
                        'YTC_VALLIBREFOU1;YTC_VALLIBREFOU2;YTC_VALLIBREFOU3;'+
                        'YTC_DATELIBREFOU1;YTC_DATELIBREFOU2;YTC_DATELIBREFOU3;'+
                        'YTC_TEXTELIBRE1;YTC_TEXTELIBRE2;YTC_TEXTELIBRE3;'+
                        'YTC_BOOLLIBRE1;YTC_BOOLLIBRE2;YTC_BOOLLIBRE3');

     ExecuteSQL('UPDATE TIERSCOMPL SET YTC_SCHEMAGEN="' + GetControlText('YTC_SCHEMAGEN') +'", ' +
                'YTC_ACCELERATEUR="'+ T.getValue('YTC_ACCELERATEUR') +'", ' +
                'YTC_DAS2="' + T.GetValue('YTC_DAS2') +'", ' +
                'YTC_PROFESSION="' + T.GetValue('YTC_PROFESSION') +'", ' +
                'YTC_REMUNERATION="' + T.GetValue('YTC_REMUNERATION') +'", ' +
                'YTC_INDEMNITE="' + T.GetValue('YTC_INDEMNITE') +'", ' +
                'YTC_RETENUESOURCE="' + T.GetValue('YTC_RETENUESOURCE') +'", ' +
                'YTC_AVANTAGE="' + T.GetValue('YTC_AVANTAGE') +'" '+ szSQL+
                'WHERE YTC_AUXILIAIRE="'+GetControlText('T_AUXILIAIRE')+'" ' );
     end
   else
     T.InsertDB(nil);
   T.Free;
end;

//================================================================================
// Initialization
//================================================================================
procedure TOM_TIERS.OnExitPays(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
    PaysRegion(THValComboBox(GetControl('T_PAYS',true)),THEdit(GetControl('T_DIVTERRIT',true)),True) ;
{$ELSE}
    PaysRegion(THDBValComboBox(GetControl('T_PAYS',true)),THDBEdit(GetControl('T_DIVTERRIT',true)),True) ;
{$ENDIF}
end;

procedure TOM_TIERS.OnchangeDAS2(Sender: TObject);
BEGIN
    if VH^.PaysLocalisation<>CodeISOES then
    begin
       ForceModif;

       if (not (GetCheckBoxState('YTC_DAS2') = cbChecked)) then
       begin
          SetControlText('YTC_PROFESSION','');
          SetControlProperty('YTC_REMUNERATION','Value','');
          SetControlProperty('YTC_INDEMNITE','Value','');
          SetControlProperty('YTC_AVANTAGE','Value','');
          SetControlProperty('YTC_RETENUESOURCE','Value','');
          SetControlEnabled('YTC_PROFESSION',False);
          SetControlEnabled('YTC_REMUNERATION',False);
          SetControlEnabled('YTC_RETENUESOURCE',False);
          SetControlEnabled('G_DAS2',false);
       end
       else
       begin
         SetControlEnabled('G_DAS2',true);
         SetControlEnabled('YTC_PROFESSION',True);
         SetControlEnabled('YTC_REMUNERATION',True);
       end;
    End ; //XMG 24/02/2005
end;

procedure TOM_TIERS.OnChangeRemuneration(Sender : Tobject) ;
begin
    ForceModif;

    if (THValComboBox(GetControl('YTC_REMUNERATION',true)).Value = 'IND') then
    begin
        SetControlEnabled('YTC_INDEMNITE',true);
        SetControlEnabled('YTC_AVANTAGE',false);
        SetControlProperty('YTC_AVANTAGE','value','');
        SetControlEnabled('YTC_RETENUESOURCE',False);
        SetControlProperty('YTC_RETENUESOURCE','value','');
    end
    else if (THValComboBox(GetControl('YTC_REMUNERATION',true)).Value = 'AVA') then
    begin
        SetControlEnabled('YTC_INDEMNITE',false);
        SetControlProperty('YTC_INDEMNITE','value','');
        SetControlEnabled('YTC_AVANTAGE',true);
        SetControlEnabled('YTC_RETENUESOURCE',False);
        SetControlProperty('YTC_RETENUESOURCE','value','');
    end
    else if (THValComboBox(GetControl('YTC_REMUNERATION',true)).Value = 'RET') then
    begin
        SetControlEnabled('YTC_INDEMNITE',false);
        SetControlProperty('YTC_INDEMNITE','value','');
        SetControlEnabled('YTC_AVANTAGE',true);
        SetControlEnabled('YTC_RETENUESOURCE',True);
        SetControlProperty('YTC_RETENUESOURCE','value','');
    end
    else begin
        SetControlEnabled('YTC_INDEMNITE',false);
        SetControlProperty('YTC_INDEMNITE','value','');
        SetControlEnabled('YTC_AVANTAGE',false);
        SetControlProperty('YTC_AVANTAGE','value','');
        SetControlEnabled('YTC_RETENUESOURCE',False);
        SetControlProperty('YTC_RETENUESOURCE','value','');          
    end;
end;

procedure TOM_TIERS.OnChangeProfession(Sender : Tobject) ;
BEGIN
  ForceModif ;
END;

procedure TOM_TIERS.OnChangeACCELERATEUR(Sender: TObject);
begin
 if csDestroying in Ecran.ComponentState then Exit ;
 ForceModif ;
 SetControlEnabled('YTC_SCHEMAGEN', GetCheckBoxState('YTC_ACCELERATEUR') = cbChecked ) ;
 if GetCheckBoxState('YTC_ACCELERATEUR') = cbUnChecked then
  SetControlText('YTC_SCHEMAGEN', '' ) ;
end;

procedure TOM_TIERS.OnChangeSCHEMAGEN(Sender: TObject);
begin
 if csDestroying in Ecran.ComponentState then Exit ;
 ForceModif ;
end;


{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 01/09/2004
Modifié le ... :   /  /    
Description .. : - LG - 01/09/2004 - FB 14313 - le compte est obligatoire si 
Suite ........ : on coche l'accelerateur
Mots clefs ... :
*****************************************************************}
function TOM_TIERS.AccelerateurValide : boolean;
begin
if (GetCheckBoxState('YTC_ACCELERATEUR') = cbUnChecked) then // or ( GetControlText('YTC_SCHEMAGEN') = '') then
 begin
  result := true ;
  exit ;
 end ;
result := false ;
if (not Presence('GENERAUX','G_GENERAL',GetControlText('YTC_SCHEMAGEN'))) then
    begin
        Pages.ActivePageIndex := 6 ;
        SetFocusControl('YTC_SCHEMAGEN');
        MsgBox.execute(44,'','');
        Exit;
    end ;
 result := true ;
end;

procedure TOM_TIERS.OnExitSCHEMAGEN(Sender: TObject);
var
  Lib, comptegene : string;
begin
 if csDestroying in Ecran.ComponentState then Exit ;
 if GetControlText('YTC_SCHEMAGEN') = '' then exit ;
 Comptegene :=BourreOuTronque(GetControlText('YTC_SCHEMAGEN'), fbGene);
 SetControlText('YTC_SCHEMAGEN', Comptegene ) ;
 if Sender = nil then exit; // pour appel dans onload
 Lib := RechDom('TZGENERAL', Comptegene, False, 'G_FERME <> "X"');
 {JP 26/08/05 : FQ 16105 : On s'assure que le compte général n'est pas fermé}
 if (Lib = '') then
 begin        // ajout me fiche 16726
   MessageAlerte('Le compte : '+ GetControlText('YTC_SCHEMAGEN') + ' est fermé');
   SetControlText('YTC_SCHEMAGEN', '');
   SetFocusControl('YTC_SCHEMAGEN');
   THEdit(GetControl('YTC_SCHEMAGEN')).ElipsisClick(GetControl('YTC_SCHEMAGEN'));
 end
 else
   SetControlCaption('G_LIBELLE', Lib);
end;

procedure TOM_TIERS.OnExitLibelle(Sender: TObject);
begin
  { En mode PCL, dès que le libellé a été saisi, on ne force plus le focus sur le libellé }
  if fbForceSaisieLibelle then
  begin
    fbForceSaisieLibelle := False;
    Pages.ActivePageIndex := PComptable;
    SetFocusControl('T_NATUREAUXI');
  end;
end;

{$IFDEF EAGLCLIENT}
{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/06/2007
Modifié le ... :   /  /
Description .. : Procedure permettant de gerer les deux zones LIBELLE
Suite ........ : (Onglet Coordonnées et DAS2)
Suite ........ : FQ 20695
Mots clefs ... :
*****************************************************************}
procedure TOM_TIERS.OnChangeLibelle (Sender : TObject );
var
  Src,Dst : string;
begin
   if ChangeLibelle then Exit; // Evite l'appel recursif.
   ChangeLibelle := true;
   if THEdit(Sender).Name = 'T_LIBELLE' then
   begin
      Src := 'T_LIBELLE';
      Dst := 'T_LIBELLE_';
   end
   else
   begin
      Src := 'T_LIBELLE_';
      Dst := 'T_LIBELLE';
   end;
   SetControlText(Dst, GetControlText(Src));
   ChangeLibelle := false;
end;
{$ENDIF}

procedure TOM_TIERS.OnAfterFormShow;
begin
  if assigned(pages) then pages.OnChanging := OnChangingPages;
  if assigned(Nature) then Nature.OnChange := OnChangeNature;

  // GCO - 06/12/2004 - Pourquoi ne l'affecter que pour la ModifEnSerie
  if FInPageIndex = 0 then FInPageIndex := 1 ;
  Pages.ActivePageIndex := FInPageIndex ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 18/08/2004
Modifié le ... :   /  /
Description .. : - SBO 18/08/2004 FQ 14217 ne forcer le mode édition que
Suite ........ : si nécessaire
Mots clefs ... :
*****************************************************************}
procedure TOM_TIERS.ForceModif ;
begin
 if (DS.State = dsBrowse) then // SBO 18/08/2004 FQ 14217 ne forcer le mode édition que si nécessaire
   begin
   DS.Edit ;
   SetField('T_AUXILIAIRE',GetControlText('T_AUXILIAIRE'));
   end ;
end ;



{***********A.G.L.***********************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 03/09/2004
Modifié le ... : 03/09/2004
Description .. : Ajout des modifications afin d'accéder à SGEP
Mots clefs ... :
*****************************************************************}

{$IFNDEF EAGLCLIENT}
procedure TOM_TIERS.AppelIsoflex(Sender: TObject);
begin
  AglIsoflexViewDoc(NomHalley, Ecran.Name, 'TIERS', 'T_AUXILIAIRE', '', GetField('T_AUXILIAIRE'), '');
end;

function TOM_TIERS.GereIsoflex : boolean ;
begin
  if AglIsoflexPresent then
  begin
    SetControlVisible('BISOFLEX', TRUE);
    result := TRUE;
  end
  else
    result := FALSE;
end;
//FIN?
{$ENDIF}

procedure TOM_TIERS.ChargeCumulsMultiSoc;
var lstReq  : String ;
    lQCumul : TQuery ;
begin

  if DS.State = dsInsert then Exit ;
  if not EstTablePartagee( 'TIERS' ) then Exit ;

  if Assigned( FTobCumuls ) then FreeAndNil( FTobCumuls ) ;

  lStReq := 'SELECT CU_DEBIT1 as T_TOTALDEBIT ,  CU_CREDIT1 as T_TOTALCREDIT ,'
                 + 'CU_DEBIT2 as T_TOTDEBP ,CU_CREDIT2 as T_TOTCREP ,'
                 + 'CU_DEBIT3 as T_TOTDEBE ,CU_CREDIT3 as T_TOTCREE ,'
                 + 'CU_DEBIT4 as T_TOTDEBS ,CU_CREDIT4 as T_TOTCRES ,'
                 + 'CU_DEBITAN as T_TOTDEBANO ,CU_CREDITAN as T_TOTCREANO ,'
                 + 'CU_DEBIT5 as T_TOTDEBANON1,CU_CREDIT5 as T_TOTCREANON1'
             + ' FROM CUMULS WHERE CU_TYPE="' + fbToCumulType(fbAux) + '" '
                           + 'AND CU_COMPTE1="' + GetField('T_AUXILIAIRE') + '"' ;

  lQCumul := OpenSQL( lStReq, True ) ;
  if not lQCumul.Eof then
    begin
    FTobCumuls := TOB.Create('_CUMULS', nil, -1 ) ;
    FTobCumuls.SelectDB('', lQCumul ) ;
    end ;
  Ferme(lQCumul);

end;

function TOM_TIERS.GetChampsMS(vStChamp: String): double;
begin
  result := 0 ;
  if DS.State = dsInsert then Exit ;

  if EstTablePartagee( 'TIERS' ) then
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
procedure TOM_TIERS.OnMouseDownT_BlocNote(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if DS.State = dsBrowse then
    DS.Edit;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOM_TIERS.AfficheCumuls;
var lStVal : String ;
begin

  // TOTAUX DU COMPTE
  // Maj des champs tampons pour affichage des montants d'infos
  lStVal := StrFMontant( GetChampsMS('T_TOTDEBP'), 15, V_PGI.OkdecV, '', True ) ;
  SetControlText('T_TOTDEBP1', lStVal ) ;
  lStVal := StrFMontant( GetChampsMS('T_TOTCREP'), 15, V_PGI.OkdecV, '', True ) ;
  SetControlText('T_TOTCREP1', lStVal ) ;
  lStVal := StrFMontant( GetChampsMS('T_TOTDEBE'), 15, V_PGI.OkdecV, '', True ) ;
  SetControlText('T_TOTDEBE1', lStVal ) ;
  lStVal := StrFMontant( GetChampsMS('T_TOTCREE'), 15, V_PGI.OkdecV, '', True ) ;
  SetControlText('T_TOTCREE1', lStVal ) ;
  lStVal := StrFMontant( GetChampsMS('T_TOTDEBS'), 15, V_PGI.OkdecV, '', True ) ;
  SetControlText('T_TOTDEBS1', lStVal ) ;
  lStVal := StrFMontant( GetChampsMS('T_TOTCRES'), 15, V_PGI.OkdecV, '', True ) ;
  SetControlText('T_TOTCRES1', lStVal ) ;

  // Calcul des soldes
  AfficheLeSolde(THNumEdit(GetControl('TSOLCREP',true)), GetChampsMS('T_TOTDEBP'),    GetChampsMS('T_TOTCREP'));
  AfficheLeSolde(THNumEdit(GetControl('TSOLCREE',true)), GetChampsMS('T_TOTDEBE'),    GetChampsMS('T_TOTCREE'));
  AfficheLeSolde(THNumEdit(GetControl('TSOLCRES',true)), GetChampsMS('T_TOTDEBS'),    GetChampsMS('T_TOTCRES'));
  AfficheLeSolde(THNumEdit(GetControl('SOLCPTA',true)),  GetChampsMS('T_TOTALDEBIT'), GetChampsMS('T_TOTALCREDIT')) ;

end;

{b FP 21/02/2006}
procedure TOM_TIERS.OnElipsisClickCORRESP(Sender: TObject);
var
  i, j:             Integer;
  OldSt:            String;
  Plan:             String;
  Nature:           String;
  OppositeNature:   String;
  sauvLookUpLocate: Boolean; {FP 25/04/2006 FQ17910}
	Mcd : IMCDServiceCOM;
	Field     : IFieldCOMEx ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  Plan    := (Sender as TComponent).Name[Length((Sender as TComponent).Name)];
  Nature  := '';
  if TCompensation.IsNatureClient(GetField('T_NATUREAUXI')) then
    begin
    Nature         := '1';      {Compte client}
    OppositeNature := '2';      {Compte fournisseur}
    end
  else if TCompensation.IsNatureFournisseur(GetField('T_NATUREAUXI')) then
    begin
    Nature         := '2';      {Compte fournisseur}
    OppositeNature := '1';      {Compte client}
    end;

  if Nature = '' then   {Ce n'est ni un compte client ni un compte fournisseur}
  else
    begin
    {Modifie le libellé de la colonne CR_LIBRETEXTE1 ou CR_LIBRETEXTE2 en 'Compte client' ou 'Compte fournisseur'}
    Field := Mcd.GetField('CR_LIBRETEXTE'+OppositeNature) as IfieldComEx;
    if Assigned(Field) then
      BEGIN
      OldSt:=Field.Libelle ;
      if OppositeNature = '1' then
        Field.Libelle:=TraduireMemoire('Compte client')       {FP 25/04/2006 FQ17909}
      else
        Field.Libelle:=TraduireMemoire('Compte fournisseur'); {FP 25/04/2006 FQ17909}
      END ;

    {b FP 25/04/2006 FQ17910: Permet de visualiser le compte client ou fournisseur correspondant}
    sauvLookUpLocate   := V_PGI.LookUpLocate;
    V_PGI.LookUpLocate := True;
    try
      LookupList((Sender as TControl),TraduireMemoire('Plan de correspondance'),'CORRESP',
        'CR_CORRESP','CR_LIBELLE,CR_LIBRETEXTE'+OppositeNature,
        'CR_TYPE="AU'+Plan+'" AND '+TCompensation.GetSQLCorrespCompte(GetField('T_NATUREAUXI'), GetField('T_AUXILIAIRE')),
        'CR_CORRESP',true, 0);
    finally
      V_PGI.LookUpLocate := sauvLookUpLocate;
      end;
    {e FP 25/04/2006}

    if Assigned(Field) then Field.Libelle := OldSt;
    end;
end;

function TOM_TIERS.VerifCorrespondance: Byte;
  function VerifCompensation(NumPlan: Integer): Byte;
  var
    NomPlan: String;
  begin
    Result := 0;
    NomPlan := TCompensation.GetNomPlanCorrespondance(GetField('T_NATUREAUXI'), GetField('T_AUXILIAIRE'));
    if NomPlan <> GetControlText('T_CORRESP'+IntToStr(NumPlan)) then
      Result := NumPlan;
  end;
begin
  if TCompensation.IsCompensation then
    begin
    if TCompensation.IsPlanCorresp(1) then
      begin
      Result := VerifCompensation(1);
      if Result = 0 then
        Result := Ent1.VerifCorrespondance(2,'',GetControlText('T_CORRESP2'));
      end
    else
      begin
      Result := VerifCompensation(2);
      if Result = 0 then
        Result := Ent1.VerifCorrespondance(2,GetControlText('T_CORRESP1'),'');
      end;
    end
  else
    Result := Ent1.VerifCorrespondance(2,GetControlText('T_CORRESP1'),GetControlText('T_CORRESP2'));
end;
{e FP 21/02/2006}

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 23/03/2006
Modifié le ... :   /  /
Description .. : Bouton copier ; duplication des champs du compte choisi
Mots clefs ... :
*****************************************************************}
procedure TOM_TIERS.BCopierClick(Sender: TObject);
// Copie la rubrique courante et place les données en mode nouveau
Var St : String ;
    Q : TQuery ;
    i : integer ;
    lowv, highv : integer;
    Cache : THCpteEdit; {JP 02/06/06 : FQ 17971}
    tz : tZoomTable;
begin
  // Test code saisi ?
   If GetControlText('T_AUXILIAIRE')='' Then
   BEGIN
     MsgBox.Execute(48,'','') ;
     Exit ;
   END ;

  tz := tzTiers;
  Cache := THCpteEdit.Create(Ecran);
  Cache.Visible := false;
  Cache.Parent := Ecran;
  Cache.ZoomTable := tz;
  Cache.okLocate := true;
  Cache.Text := GetControlText('T_AUXILIAIRE') ;

  {JP 02/06/06 : FQ 17971 : changement du LookUp}  
  if (GChercheCompte(Cache, FicheTiers)) then begin
    st := Cache.Text;
  // Choix de la rubrique a dupliquer
  //  St := Choisir(MsgBox.Mess[50],'TIERS','T_AUXILIAIRE || " / " || T_LIBELLE','T_AUXILIAIRE','T_AUXILIAIRE<>"'+'ymo'+'"','T_AUXILIAIRE');
    if St='' then exit;
    // Confirmation
    if MsgBox.Execute(49,'','')<>mrYes then Exit ;
    SetControlEnabled('BCopier',False);

    // Changement d'état si nécessaire
    if (not (DS.State in [dsEdit,dsInsert])) then DS.Edit;

    lowv:=Low(DuplicValues);
    highv:=High(DuplicValues);
          // Copie
    Q := OpenSQL('Select * From TIERS Where T_AUXILIAIRE="' + St + '"', TRUE) ;
    if not Q.Eof then
    begin

      // YMO 16/05/20006 FQ17889 Remplacement des SetControlText par SetField
      for i:=Lowv to Highv do
        SetField(DuplicValues[i], Q.FindField(DuplicValues[i]).AsString) ;

      // onglet tables libres
      for i:=0 to 9 do
        SetField('T_TABLE'+inttostr(i), Q.FindField('T_TABLE'+inttostr(i)).AsString) ;

      SetControlChecked('TCONFIDENTIEL', (Q.FindField('T_CONFIDENTIEL').AsString = '1') ) ;

    end;
    Ferme(Q) ;
  end;
  SetControlEnabled('T_AUXILIAIRE',DS.State in [dsInsert]);

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/07/2006
Modifié le ... :   /  /
Description .. : FQ 18539
Mots clefs ... :
*****************************************************************}
procedure TOM_TIERS.OnKeyPressT_Auxiliaire(Sender: TObject; var Key: Char);
begin
  if EstSpecif('51205') then Exit;

  if not (key in ['a'..'z', 'A'..'Z', '0'..'9', #8, GetInfoCpta(fbAux).cb]) then
    Key := #0;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOM_TIERS.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin

  // Empêcher le message sur fermeture dans modif en création // SBO 19/04/2007
  if ( DS.State = dsInsert ) and
     ( GetControlText('T_AUXILIAIRE') = '' ) then Exit ;

  if Assigned( FAGLCloseQuery ) then
    FAGLCloseQuery(Sender, Canclose ) ;

end;

Initialization
    registerclasses([TOM_TIERS]);
end.
