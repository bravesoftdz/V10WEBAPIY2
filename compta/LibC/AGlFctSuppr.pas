{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 20/02/2003
Modifié le ... :   /  /
Description .. : Passage en eAGL
Mots clefs ... :
Suite......... : mbo - 17/10/2006 - en suppression de compte : ajout test sur comptes de subvention d'invest.
Suite......... : mbo - 03/11/2006 - en suppression de compte : le compte de reprise de subvention n'est plus stocké
                                    (on le reconstitue à partir du compte de subvention)
Suite......... : BTY - 06/11/07 Suppression de compte : bloquer si présent dans IMMOREGFR des passages au régime réel des biens agricoles
Suite......... : => pas de test ds EstDansBase (EstDansMultiSoc) qui sert si tables partagées Entreprise (module immobilisations est non partageable)
Suite......... : BTY - 07/11/07 FQ 21792 Suppression de compte : ajout test compte de charge dans IMMO
*****************************************************************}

unit AGlFctSuppr;     // Programme de suppression générale

//================================================================================
// Interface
//================================================================================
interface

uses
    Forms,
    Classes,          // TList,
    Controls,         // mrYes
    Hctrls,           // ExecuteSQL, ExisteSQL, Presence, OpenSQL, Ferme
    MajTable,
{$IFDEF EAGLCLIENT}
    eMul,             // TFMul
{$ELSE}
    {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
    HDB,              // THDBGrid, NbSelected, AllSelected
    Mul,              // TFMul
{$ENDIF}
{$IFDEF MODENT1}
    CPTypeCons,
{$ENDIF MODENT1}
    SysUtils,         // UpperCase
    ParamSoc,         // GetParamSoc
    CPSection_TOM,    // EstDansAxe, EstCorresp, EstDansVentil
    CPGeneraux_TOM,   // EstAnalytiquePure, EstCpteAxe, EstCpteCorresp, EstCpteDevise, EstCpteJournal, EstCpteModePaie, EstCpteSociete, EstCpteTVA, EstEcrGuide, EstCpteCorresp
    SaisUtil,         // EstTVATPF
    Ent1,             // TFichierBase, VH, fbGene, fbAux, fbJal, fbSect
    HStatus,          // MoveCur, InitMove, FiniMove
    RapSuppr,         // DelInfo, RapportDeSuppression
    M3FP,             // RegisterAglProc
    uTob,             // Tob
    ed_tools,         // VideListe
    UtilPGI,          // EstMultiSoc, EstTablePartagee, fbToTable, GetBasesMS
    HMsgBox,          // MessageAlerte, PGIAsk, HShowMessage
    CPJournal_TOM     // EstDansSociete, EstDansEcriture, EstDansSouche
    ;

//==================================================
// Externe
//==================================================
procedure CPSupprimeListeEnreg(vTobFListe : Tob ; vType : TFichierBase);
procedure AGLSupprimeListCpte(parms: array of variant; nb: integer );

{$IFDEF EAGLCLIENT}
procedure SupprimeListeEnreg(L : THGrid; Q :TQuery; St,Titre:string);
{$ELSE}
procedure SupprimeListeEnreg(L : THDBGrid; Q :TQuery; St,Titre:string);
{$ENDIF}


//==================================================
// Definition de class
//==================================================
Type
    TEvtSup = Class
        Code    : string;
        Codax   : string;
        GCCode  : string;

        {JP 27/12/05 : FQ 11919 : Demande de message plus détaillé pour le comptes non detruits}
        MessageErreur : string;

        FType   : TFichierBase ;
        FDossier   : String ;

        procedure DegageGene;
        procedure Degagesection;
        procedure DegageJournal;
        procedure DegageTiers;

        Function  EstCompteSupprimable                      : Integer ;

        // Evt mixtes
        Function  EstMouvemente        : Boolean ;   // ok
        Function  EstDansAxe           : Boolean ;   // ok
        Function  EstCpteCorresp       : Boolean ;
        Function  EstDansSociete       : Boolean ; // ok
        Function  EstEcrGuide          : Boolean ; // ok
        Function  EstDansPiece         : Boolean ; // ok
        Function  EstDansAmortissement : boolean;  // ok
        Function  EstDansGCD           : boolean;  //
        Function  EstDansAgricole      : boolean;  // BTY 06/11/07 Agricole


        // Evt Généraux
        function  EstCpteJournal       : Boolean ; // ok
        function  EstCpteCollRef       : Boolean ; // ok
        function  EstCpteModepaie      : Boolean ; // ok
        function  EstCpteTva           : Boolean ; // ok
        function  EstCpteDevise        : Boolean ; // ok
        function  EstCodeCptaGC        : Boolean; {JP 27/12/05 : FQ 11919}
        function  EstTableauVar        : Boolean; //Tableau des variations
        function  EstNoteTravail       : Boolean; //note de travail

        // Evt Tiers
        Function  EstDansSection       : Boolean ;  // ok
        Function  EstUnPayeur          : Boolean ;  // ok
        Function  EstDansUtilisat      : Boolean ;  // ok
        Function  EstDansActivite      : Boolean ;  // ok
        Function  EstDansRessource     : Boolean ;  // ok
        Function  EstDansAffaire       : Boolean ;  // ok
        Function  EstDansActions       : Boolean ;  // ok
        Function  EstDansPersp         : Boolean ;  // ok
        Function  EstDansCata          : Boolean ;  // ok
        Function  EstDansPaie          : Boolean ;  // ok
        Function  EstDansMvtPaie       : Boolean ;  // ok

        // Evt Journaux
        Function  EstDansSouche          : Boolean ; // ok
        function  EstDansParamPiece      : Boolean ; // ok
        function  EstDansParamPieceCompl : Boolean ; // ok
        function  EstDansGuide           : Boolean ;


        // Evt Section
        function  EstVentilType          : Integer ; // ok
        function  EstSectionAxeTva       : Boolean ;

        // Ajout Test pour le multisociété
        Function  EstDansMultiSoc                           : Boolean ;
        Function  EstDansBase            ( vBase : String ) : Integer ;

    end;

//================================================================================
// Implementation
//================================================================================
implementation


uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  HEnt1;


//==================================================
// Definition des Variables
//==================================================
var
    TNotDel : TList;
    TDelGene : TList;
    Effacer : Boolean;
    NotEffacer : Boolean;

    HMGEN : array [0..21] of string = (
    {00}    'Compte supprimé',
    {01}    'Compte en cours d''utilisation',
    {02}    '2;Suppression des comptes généraux;Désirez-vous un compte-rendu des comptes détruits ?;Q;YNC;N;C;',
    {03}    '3;Suppression des comptes généraux;Désirez-vous un compte-rendu des comptes non détruits ?;Q;YNC;N;C;',
    {04}    'Ce compte est mouvementé ou dans une balance de situation.',
    {05}    'Ce compte possède des écritures analytiques',
    {06}    'Ce compte est un compte d''attente pour un axe analytique',
    {07}    'Ce compte est un compte de correspondance',
    {08}    'Ce compte est un compte d''écart de change pour les devises',
    {09}    'Ce compte est un compte de contrepartie d''un journal de banque',
    {10}    'Ce compte est associé à un mode de paiement',
    {11}    'Ce compte est un compte d''attente, de clôture d''ouverture ou d''écart Euro',
    {12}    'Ce compte est associé à un paramétrage de TVA',
    {13}    'Ce compte est référencé par des tiers',
    {14}    'Ce compte est associé à une fiche immobilisation',
    {15}    'Ce compte est renseigné dans les paramètres société',
    {16}    'Ce compte est utilisé dans une autre société',
    {17}    'Ce compte auxiliaire est référençé dans un guide d''écriture comptable',
    {18}    'Ce compte général est utilisé dans la ventilation comptable de la GC', {JP 27/12/05 : FQ 11919}
    {19}    'Ce compte est associé avec un tableau des variations', //Tableau des variations
    {20}    'Ce compte est associé avec une note de travail',
    {21}    'Ce compte est associé à un bien (spécificités agricoles)' {BTY 06/11/07 Agricole}
            );

    HMAUX : array [0..23] of string = (
    {00}    'Compte supprimé',
    {01}    'Compte en cours d''utilisation',
    {02}    '2;Suppression des comptes auxiliaires;Désirez-vous un compte rendu des comptes détruits ?;Q;YNC;N;C;',
    {03}    '3;Suppression des comptes auxiliaires;Désirez-vous un compte rendu des comptes non détruits ?;Q;YNC;N;C;',
    {04}    'Ce compte auxiliaire est mouvementé ou dans une balance de situation',
    {05}    'Ce compte auxiliaire est référençé dans un guide d''écriture comptable',
    {06}    'Ce compte auxiliaire est un compte de correspondance',
    {07}    'Ce compte auxiliaire est un client ou un fournisseur d''attente',
    {08}    'Ce compte auxiliaire est référencé dans une section analytique',
    {09}    'Ce compte auxiliaire est un utilisateur',
    {10}    'Ce compte est mouvementé en pièces de Gestion Commerciale',
    {11}    'Ce compte est référencé dans les activités',
    {12}    'Ce compte est référencé dans les ressources',
    {13}    'Ce compte est référencé en affaire',
    {14}    'Ce compte est référencé dans les actions',
    {15}    'Ce compte est référencé dans les perspectives',
    {16}    'Ce compte est référencé comme salarié',
    {17}    'Ce compte est référencé dans le catalogue',
    {18}    'Ce compte est mouvementé dans la paie comme salarié',
    {19}    'Ce compte est un compte auxiliaire payeur',
    {20}    'Ce compte auxiliaire est référencé dans une fiche immobilisation',
    {21}    'Ce compte est renseigné dans les paramètres société',
    {22}    'Ce compte auxiliaire est utilisé dans une autre société' ,
    {23}    'Ce compte auxiliaire est utilisé dans la gestion des créances douteuses' 
            );

    HMSEc : array [0..13] of string = (
    {00}    'Section supprimée',
    {01}    'Section en cours d''utilisation',
    {02}    '2;Suppression des Sections;Désirez-vous un compte-rendu des sections détruites ?;Q;YNC;N;C;',
    {03}    '3;Suppression des Sections;Désirez-vous un compte-rendu des sections non détruites ?;Q;YNC;N;C;',
    {04}    'La section n''a pas pu être supprimée',
    {05}    'Cette section possède des écritures analytiques',
    {06}    'Cette section possède des axes analytiques',
    {07}    'Cette section est une section de correspondance',
    {08}    'Section utilisée dans le paramétrage de ventilations types',
    {09}    'Section utilisée dans le paramétrage de ventilations par défaut de comptes généraux',
    {10}    'Section utilisée dans le paramétrage de ventilations par défaut de fiches d''immobilisations',
    {11}    'Cette section est utilisée dans une autre société',
    {12}    'Impossible de modifier une section de l''axe TVA', // GCO - 30/01/2006
    {13}    'Section utilisée dans la pré-ventilation analytique' // BVE 10.04.07 FQ 19553
            );

    HMJAL : array [0..13] of string = (
    {00}    'Journal supprimé',
    {01}    'Journal en cours d''utilisation',
    {02}    '2;Suppression des journaux;Désirez-vous un compte-rendu des comptes détruits ?;Q;YNC;N;C;',
    {03}    '3;Suppression des journaux;Désirez-vous un compte-rendu des comptes non détruits ?;Q;YNC;N;C;',
(* FQ 19592 BVE 10.04.07
    {04}    'Ce journal à des écritures analytiques',
END FQ 19592*)
    {04}    'Ce journal a des écritures',
    {05}    'Ce journal a des écritures comptables',
    {06}    'Ce journal est réservé (ouverture, fermeture, dotations, échéances...)',
    {07}    'Ce journal a des écritures comptables pour la gestion commerciale',
    {08}    '8;Suppression des journaux;Vous ne pouvez pas supprimer ce journal : il est mouvementé par des écritures analytiques.;W;O;O;O;',
    {09}    '9;Suppression des journaux;Vous ne pouvez pas supprimer ce journal : il est mouvementé par des écritures comptables.;W;O;O;O;',
    {10}    '13;Suppression des journaux;Vous ne pouvez pas supprimer ce journal : il est référencé dans une pièce commerciale.;W;O;O;O;',
    {11}    '14;Suppression des journaux;Vous ne pouvez pas supprimer ce journal : il est référencé dans le paramétrage des pièces commerciales.;W;O;O;O;',
    {12}    'Ce journal est utilisé dans une autre société',
    {13}    'Ce journal est utilisé dans les guides'
            );



//==================================================
// Fonction Hors class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Lancement de la destruction d'un compte general
Mots clefs ... :
*****************************************************************}
Function DetruitGene(Code,Libelle : String) : Byte;
var
    EvtSup : TEvtSup;
    X,Y : DelInfo;
    Bases, Dossier : String ;
begin
    EvtSup        := TEvtSup.Create;
    EvtSup.Code   := Code;
    EvtSup.FType  := fbGene ;

    // test si le compte est supprimable
    result := EvtSup.EstCompteSupprimable ;

    // suppression
    if (Result = 0) then
    begin
        if (Transactions(EvtSup.DegageGene,5) <> oeOK) then
        begin
            MessageAlerte('Suppression impossible');
            Result := 17;
        end
        else
        begin
            ExecuteSQL('DELETE FROM VENTIL WHERE V_COMPTE="'+Code+'" AND V_NATURE LIKE"GE%"');
            ExecuteSQL('DELETE FROM CLIENGENEMODELA WHERE CLA_GENERAL="'+Code+'"');

            If EstMultiSoc and EstTablePartagee('GENERAUX') and not EstTablePartagee('BANQUECP') then
            begin //YMO 25/10/2006 Multidossiers : Suppression dans tous les dossiers
                Bases:=GetBasesMS('',True);
                While Bases<>'' do
                begin
                      Dossier:=ReadTokenSt(Bases);
                      ExecuteSql('DELETE FROM '+GetTableDossier(Dossier,'BANQUECP')
                                +' WHERE BQ_GENERAL="'+Code+'"');
                end;
            end
            else
              ExecuteSql('DELETE FROM BANQUECP WHERE BQ_GENERAL="'+Code+'"');
            //AND BQ_NODOSSIER = "'+V_PGI.NoDossier+'"');
        end;
    end;

    // creation du compte rendu
    if  (Result = 0) then
    begin
        X := DelInfo.Create;
        X.LeCod := Code;
        X.LeLib := Libelle;
        X.LeMess := HMGEN[0];
        TDelGene.Add(X);
        Effacer := True;
    end
    else
    begin
        Y := DelInfo.Create;
        Y.LeCod := Code;
        Y.LeLib := Libelle;
        {JP 27/12/05 : FQ 11919 : Dans certains cas, un message plus précis est souhaité}
        if EvtSup.MessageErreur <> '' then Y.LeMess := EvtSup.MessageErreur
                                      else Y.LeMess := HMGEN[Result] ;
        TNotDel.Add(Y);
        NotEffacer := True;
    end;

    EvtSup.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Lancement de la destruction d'un compte auxiliaire
Mots clefs ... :
*****************************************************************}
Function DetruitTiers(St,RefGC,Lib : String) : Byte;
var
    EvtSup : TEvtSup;
    X,Y : DelInfo;
begin
    EvtSup := TEvtSup.Create;
    EvtSup.Code := St;
    EvtSup.GCCode := RefGc;
    EvtSup.FType  := fbAux ;

    // test si le compte est supprimable
    result := EvtSup.EstCompteSupprimable ;

    // suppression
    if (Result = 0) then
    begin
        if (Transactions(EvtSup.DegageTiers,5) <> oeOK) then
        begin
            MessageAlerte('Suppression impossible');
            Result := 17;
        end;
    end;

    // creation du compte rendu
    if (Result = 0) then
    begin
        X := DelInfo.Create;
        X.LeCod := St;
        X.LeLib := Lib;
        X.LeMess := HMAUX[0];
        TDelGene.Add(X);
        Effacer := True;
    end
    else
    begin
        Y := DelInfo.Create;
        Y.LeCod := St;
        Y.LeLib := Lib;
        {JP 27/12/05 : FQ 11919 : Dans certains cas, un message plus précis est souhaité}
        if EvtSup.MessageErreur <> '' then Y.LeMess := EvtSup.MessageErreur
                                      else Y.LeMess := HMAUX[Result] ;
        TNotDel.Add(Y);
        NotEffacer := True;
    end;

    EvtSup.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Lancement de la destruction d'une section
Mots clefs ... :
*****************************************************************}
Function DetruitSection(Stc,Sta,lib : String) : Byte;
var
    EvtSup : TEvtSup;
    X,Y : DelInfo;
begin
    EvtSup := TEvtSup.Create;
    EvtSup.Code := Stc;
    EvtSup.Codax := Sta;
    EvtSup.FType  := fbSect ;

    // test si le compte est supprimable
    result := EvtSup.EstCompteSupprimable ;

    // suppression
    if (Result = 0) then
    begin
        if (Transactions(EvtSup.Degagesection,5) <> oeOK) then
        begin
            MessageAlerte('Suppression impossible');
            Result := 1;
        end;
    end;

    // creation du compte rendu
    if (Result = 0) then
    begin
        X := DelInfo.Create;
        X.LeCod := Stc;
        X.LeLib := lib;
        X.LeMess := HMSec[0];
        TDelGene.Add(X);
        Effacer := True;
    end
    else
    begin
        Y := DelInfo.Create;
        Y.LeCod := Stc;
        Y.LeLib := lib;
        {JP 27/12/05 : FQ 11919 : Dans certains cas, un message plus précis est souhaité}
        if EvtSup.MessageErreur <> '' then Y.LeMess := EvtSup.MessageErreur
                                      else Y.LeMess := HMSec[Result] ;
        TNotDel.Add(Y);
        NotEffacer := True;
    end;

    EvtSup.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Lancement de la destruction d'un journal
Mots clefs ... :
*****************************************************************}
Function DetruitJournal(Stc,Lib : String):Byte ;
var
    EvtSup : TEvtSup;
    X,Y : DelInfo;
begin
    EvtSup := TEvtSup.Create;
    EvtSup.Code := Stc;
    EvtSup.FType  := fbJal ;

    // test si le compte est supprimable
    result := EvtSup.EstCompteSupprimable ;

    // suppression
    if (Result = 0) then
    begin
        if (Transactions(EvtSup.DegageJournal,5) <> oeOK) then
        begin
            MessageAlerte('Suppression impossible');
            Result := 17;
        end;
    end;

    // creation du compte rendu
    if (Result = 0) then
    begin
        X := DelInfo.Create;
        X.LeCod := Stc;
        X.LeLib := '';
        X.LeMess := HMJAL[0];
        TDelGene.Add(X);
        Effacer := True;
    end
    else
    begin
        Y := DelInfo.Create;
        Y.LeCod := Stc;
        Y.LeLib := '';
        Y.LeMess := HMJAL[Result];
        TNotDel.Add(Y);
        NotEffacer := True;
    end;

    EvtSup.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/11/2002
Modifié le ... :   /  /
Description .. : Supprime les enregistrements sélectionnés dans vFListe
Suite ........ : on utilise VTob pour avoir les infos de l' enregistrement
Suite ........ : vType = fbGene, fbAux, FbJal, fbSect
Mots clefs ... :
*****************************************************************}
procedure CPSupprimeListeEnreg(vTobFListe : Tob ; vType : TFichierBase);
var
    i : integer;
    Q1 : TQuery;
begin
    // message d'avertisement !
    if (PGIAsk('Vous allez supprimer définitivement les informations.#13#10Confirmez vous l''opération ?', '') <> mrYes) then exit;

    TDelGene := TList.Create ; // Liste des comptes supprimés
    TNotDel := TList.Create ; // Liste des comptes non suprrimés
    Effacer := False ;
    NotEffacer := False ;

    // pour chaque element .. destruction
    for i := 0 to vTobFListe.Detail.Count-1 do
    begin
        if (vTobFListe.Detail[i].GetValue('SELECT') = 'X') then
        begin
            case vType of
                fbGene : if (DetruitGene(vTobFListe.Detail[i].GetValue('G_GENERAL'),vTobFListe.Detail[i].GetValue('G_LIBELLE')) > 0) then break;
                fbJal :  if (DetruitJournal(vTobFListe.Detail[i].GetValue('J_JOURNAL'),vTobFListe.Detail[i].GetValue('J_LIBELLE')) > 0) then break;
                fbSect : if (DetruitSection(vTobFListe.Detail[i].GetValue('S_SECTION'),vTobFListe.Detail[i].GetValue('S_AXE'),vTobFListe.Detail[i].GetValue('S_LIBELLE')) > 0) then break;
                fbAux :
                    begin
                        Q1 := OpenSql ('SELECT T_TIERS FROM TIERS WHERE T_AUXILIAIRE="' + vTobFListe.Detail[i].GetValue('T_AUXILIAIRE') + '"',TRUE);
                        if (not Q1.EOF) then if (DetruitTiers(vTobFListe.Detail[i].GetValue('T_AUXILIAIRE'), Q1.FindField('T_TIERS').AsString, vTobFListe.Detail[i].GetValue('T_LIBELLE')) > 0) then begin ferme(Q1); break; end;
                        Ferme(Q1);
                    end;
            end;
        end;
    end;

    // rapport
    case vType of
        fbGene :
            begin
                if (Effacer)    then if (HShowMessage(HMGEN[2],'','') = mrYes) then RapportDeSuppression(TDelGene,1);
                if (NotEffacer) then if (HShowMessage(HMGEN[3],'','') = mrYes) then RapportDeSuppression(TNotDel,1);
            end;
        fbAux :
            begin
                if (Effacer)    then if (HShowMessage(HMAUX[2],'','') = mrYes) then RapportDeSuppression(TDelGene,1);
                if (NotEffacer) then if (HShowMessage(HMAUX[3],'','') = mrYes) then RapportDeSuppression(TNotDel,1);
            end;
        fbJal :
            begin
                if (Effacer)    then if (HShowMessage(HMJAL[2],'','') = mrYes) then RapportDeSuppression(TDelGene,1);
                if (NotEffacer) then if (HShowMessage(HMJAL[3],'','') = mrYes) then RapportDeSuppression(TNotDel,1);
            end;
        fbSect :
            begin
                if (Effacer)    then if (HShowMessage(HMSEC[2],'','') = mrYes) then RapportDeSuppression(TDelGene,1);
                if (NotEffacer) then if (HShowMessage(HMSEC[3],'','') = mrYes) then RapportDeSuppression(TNotDel,1);
            end;
    end;

    // destruction
    VideListe(TDelGene);
    VideListe(TNotDel);
    TDelGene.Free;
    TNotDel.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Supprime dans une THGrid
Mots clefs ... :
*****************************************************************}
{$IFDEF EAGLCLIENT}
procedure SupprimeListeEnreg(L : THGrid; Q : TQuery; St,Titre : string);
{$ELSE}
procedure SupprimeListeEnreg(L : THDBGrid; Q : TQuery; St,Titre : string);
{$ENDIF}
var
    i : integer;
    Q1 : TQuery;
    vType : TFichierBase;
begin
    // si rien de selectionné !
    if ((L.NbSelected = 0) and (not L.AllSelected)) then
    begin
        MessageAlerte('Aucun élément sélectionné.');
        exit;
    end;

    // message d'avertisement !
    if (HShowMessage('0;'+Titre+';'+TraduireMemoire('Vous allez supprimer définitivement les informations. Confirmez vous l''opération ?')+';Q;YN;N;N;','','') <> mrYes) then exit; 

    TDelGene := TList.Create;
    TNotDel := TList.Create;
    Effacer := False;
    NotEffacer := False;

    // set du type ;)
    if (St ='GENERAUX')     then vType := fbGene
    else if (St ='SECTION') then vType := fbSect
    else if (St ='JOURNAL') then vType := fbJal
    else if (St = 'TIERS')  then vType := fbAux
    else vType := fbNone; // valeur n'import ki fait pas de traitement derrier ;)

    // destruction
    if (L.AllSelected) then // si tout ??
    begin
        Q.First;

        while (Not Q.EOF) do
        begin
            MoveCur(False);

            case vType of
                fbGene : DetruitGene(Q.FindField('G_GENERAL').AsString,Q.FindField('G_LIBELLE').AsString);
                fbSect : DetruitSection(Q.FindField('S_SECTION').AsString,Q.FindField('S_AXE').AsString,Q.FindField('S_LIBELLE').AsString);
                fbJal :  DetruitJournal(Q.FindField('J_JOURNAL').AsString,Q.FindField('J_LIBELLE').AsString);
                fbAux :
                    begin
                        Q1 := OpenSql ('SELECT T_TIERS FROM TIERS Where T_AUXILIAIRE="'+ Q.FindField('T_AUXILIAIRE').AsString+'"',TRUE);
                        if (not Q1.EOF) then DetruitTiers(Q.FindField('T_AUXILIAIRE').AsString,Q1.FindField('T_TIERS').AsString,Q.FindField('T_LIBELLE').AsString);
                        Ferme(Q1);
                    end;
            end;

            Q.Next;
        end;

        L.AllSelected:=False;
    end
    else // si pas tout ....
    begin
        InitMove(L.NbSelected,'');

        for i := 0 to L.NbSelected-1 do
        begin
            MoveCur(False);

            L.GotoLeBookMark(i);
{$IFDEF EAGLCLIENT}
            Q.Seek(L.Row-1);
{$ENDIF}

            // attention .... ici en pas CWAS avec la grille de la mul de suppression on a pas forcement les bonnes info ...
            case vType of
                fbGene : DetruitGene(Q.FindField('G_GENERAL').AsString,Q.FindField('G_LIBELLE').AsString);
                fbSect : DetruitSection(Q.FindField('S_SECTION').AsString,Q.FindField('S_AXE').AsString,Q.FindField('S_LIBELLE').AsString);
                fbJal :  DetruitJournal(Q.FindField('J_JOURNAL').AsString,Q.FindField('J_LIBELLE').AsString);
                fbAux :
                    begin
                        Q1 := OpenSql ('SELECT T_TIERS FROM TIERS WHERE T_AUXILIAIRE="'+ Q.FindField('T_AUXILIAIRE').AsString+'"',TRUE);
                        if (not Q1.EOF) then DetruitTiers(Q.FindField('T_AUXILIAIRE').AsString,Q1.FindField('T_TIERS').AsString,Q.FindField('T_LIBELLE').AsString);
                        Ferme(Q1);
                    end;
            end;
        end;

        L.ClearSelected;
    end;

    FiniMove;

    // rapport
    case vType of
        fbGene :
            begin
                if (Effacer)    then if (HShowMessage(HMGEN[2],'','') = mrYes) then RapportDeSuppression(TDelGene,1);
                if (NotEffacer) then if (HShowMessage(HMGEN[3],'','') = mrYes) then RapportDeSuppression(TNotDel,1);
            end;
        fbAux :
            begin
                if (Effacer)    then if (HShowMessage(HMAUX[2],'','') = mrYes) then RapportDeSuppression(TDelGene,1);
                if (NotEffacer) then if (HShowMessage(HMAUX[3],'','') = mrYes) then RapportDeSuppression(TNotDel,1);
            end;
        fbJal :
            begin
                if (Effacer)    then if (HShowMessage(HMJAL[2],'','') = mrYes) then RapportDeSuppression(TDelGene,1);
                if (NotEffacer) then if (HShowMessage(HMJAL[3],'','') = mrYes) then RapportDeSuppression(TNotDel,1);
            end;
        fbSect :
            begin
                if (Effacer)    then if (HShowMessage(HMSEC[2],'','') = mrYes) then RapportDeSuppression(TDelGene,1);
                if (NotEffacer) then if (HShowMessage(HMSEC[3],'','') = mrYes) then RapportDeSuppression(TNotDel,1);
            end;
    end;

    // destruction
    VideListe(TDelGene);
    VideListe(TNotDel);
    TDelGene.Free;
    TNotDel.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Point d'entre script
Mots clefs ... :
*****************************************************************}
procedure AGLSupprimeListCpte(parms: array of variant; nb: integer );
var
    F : TFMul;
{$IFDEF EAGLCLIENT}
    Liste : THGrid;
{$ELSE}
    Liste : THDBGrid;
{$ENDIF}
    Query : TQuery;
begin
    F := TFMul(Longint(Parms[0]));
    if (F = Nil) then exit;

{$IFDEF EAGLCLIENT}
    Liste := THGrid(F.FindComponent('FListe'));
{$ELSE}
    Liste := THDBGrid(F.FindComponent('FListe'));
{$ENDIF}
    if (Liste = Nil) then exit;

{$IFDEF EAGLCLIENT}
    Query := F.Q.TQ;
{$ELSE}
    Query := F.Q;
{$ENDIF}
    if (Query = Nil) then exit;

    SupprimeListeEnreg(Liste,Query,parms[1],F.Caption);
end;


{ TEvtSup }

//==================================================
// Fonction de la class
//==================================================
//=========================
// Fonction de Suppression
//=========================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Suppression effective d'un compte general
Mots clefs ... :
*****************************************************************}
procedure TEvtSup.DegageGene;
begin
  // GCO - 08/02/2007
  ExecuteSql('DELETE FROM CREVGENERAUX WHERE CRG_GENERAL = "' + Code + '"');
  ExecuteSql('DELETE FROM CREVBLOCNOTE WHERE CBN_NATURE = "GEN" AND CBN_CODE = "' + Code + '"');
  // FIN GCO

  if (ExecuteSQL('DELETE FROM GENERAUX WHERE G_GENERAL="'+Code+'"') <> 1) then V_PGI.IoError := oeUnknown;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 24/08/2005
Description .. : Suppression effective d'un compte auxiliaire
Suite ........ : - LG  - 24/08/2005 - suppression de la mise a jour ds 
Suite ........ : l'annuaire qd on supprime le tiers
Mots clefs ... : 
*****************************************************************}
procedure TEvtSup.DegageTiers;
begin
    if (ExecuteSQL('DELETE FROM TIERS WHERE T_AUXILIAIRE="'+Code+'"') <> 1) then V_PGI.IoError := oeUnknown
    else
    begin
        ExecuteSQL('DELETE FROM CONTACT WHERE C_AUXILIAIRE="'+Code+'"') ;
        ExecuteSQL('DELETE FROM RIB WHERE R_AUXILIAIRE="'+Code+'"') ;
        // Gescom
        // CA - 07/12/2005 - Cas des bases optimisées sans table PIECE
        if not EstBasePclAllegee then
        begin
          ExecuteSQL('DELETE FROM TIERSPIECE WHERE GTP_TIERS="'+GCCode+'"') ;
          ExecuteSQL('DELETE FROM ARTICLETIERS WHERE GAT_REFTIERS="'+GCCode+'"') ;
          ExecuteSQL('DELETE FROM TARIF WHERE GF_TIERS="'+GCCode+'"') ;
          ExecuteSQL('DELETE FROM PROSPECTS WHERE RPR_AUXILIAIRE="'+Code+'"') ;
        end;
        ExecuteSQL('DELETE FROM ADRESSES WHERE ADR_TYPEADRESSE="TIE" AND ADR_REFCODE="'+GCCode+'"') ;
        ExecuteSQL('DELETE FROM LIENSOLE WHERE LO_TABLEBLOB="T" AND LO_IDENTIFIANT="'+GCCode+'"') ;
        ExecuteSQL('DELETE FROM TIERSCOMPL WHERE YTC_TIERS="'+GCCode+'"') ;
        //mcd 17/05/05 champ auxi supprimé  ExecuteSQL('UPDATE ANNUAIRE SET ANN_TIERS="", ANN_AUXILIAIRE="" WHERE ANN_TIERS="'+GCCode+'"') ;
        if V_PGI.InBaseCommune then
         ExecuteSQL('UPDATE ANNUAIRE SET ANN_TIERS="" WHERE ANN_TIERS="'+GCCode+'"') ;
        // Paie
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Suppression effective d'une section
Mots clefs ... :
*****************************************************************}
procedure TEvtSup.Degagesection;
begin
    if (ExecuteSQL('DELETE FROM SECTION WHERE S_SECTION="'+Code+'" AND S_AXE="'+Codax+'"') <> 1) then V_PGI.IoError := oeUnknown;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Suppression effective d'un journal
Mots clefs ... :
*****************************************************************}
procedure TEvtSup.DegageJournal;
begin
    if (ExecuteSQL('DELETE FROM JOURNAL WHERE J_JOURNAL="'+Code+'"') <> 1) then V_PGI.IoError := oeUnknown;
end;


//=========================
// Fonction de Verification
//=========================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Verification sur les comptes auxiliaires et generaux
Mots clefs ... :
*****************************************************************}
Function TEvtSup.EstMouvemente : Boolean;
begin
  Result := True;

  Case FType of
   // generaux
    fbGene : begin
             Result := ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="'+Code+'"'
                                  + ' AND ((EXISTS(SELECT E_GENERAL FROM ' + GetTableDossier( FDossier, 'ECRITURE' ) + ' WHERE E_GENERAL="'+Code+'"))'
                                  + ' OR (EXISTS(SELECT Y_GENERAL FROM ' + GetTableDossier( FDossier, 'ANALYTIQ' ) + ' WHERE Y_GENERAL="'+Code+'"))'
                                  + ' OR (EXISTS(SELECT BSE_COMPTE1 FROM ' + GetTableDossier( FDossier, 'CBALSITECR' ) + ' WHERE BSE_COMPTE1="'+Code+'")))');
             end ;

    // auxiliaire
    fbAux : begin
            Result := ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="'+Code+'"'
                                  + ' AND ((EXISTS(SELECT E_AUXILIAIRE FROM ' + GetTableDossier( FDossier, 'ECRITURE' ) + ' WHERE E_AUXILIAIRE="'+Code+'" ))'
                                  + ' OR (EXISTS(SELECT BSE_COMPTE1,BSE_COMPTE2 FROM ' + GetTableDossier( FDossier, 'CBALSITECR' ) + ' WHERE BSE_COMPTE1="'+Code+'" OR BSE_COMPTE2="'+Code+'")))') ;
            end ;

   // journaux
    fbJal :  begin
             Result := ExisteSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_JOURNAL="'+Code+'"'
                                  + ' AND ((EXISTS(SELECT E_JOURNAL FROM ' + GetTableDossier( FDossier, 'ECRITURE' ) + ' WHERE E_JOURNAL="'+Code+'"))'
                                  + ' OR (EXISTS(SELECT Y_JOURNAL FROM ' + GetTableDossier( FDossier, 'ANALYTIQ' ) + ' WHERE Y_JOURNAL="'+Code+'")))' ) ;
             end ;

    // Section
    fbSect :  Result := ExisteSQL('SELECT Y_SECTION FROM ' + GetTableDossier( FDossier, 'ANALYTIQ' ) + ' WHERE Y_SECTION="' + Code + '" AND Y_AXE="' + CodAx + '"');

    end ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Verification sur les comptes auxiliaires
Mots clefs ... :
*****************************************************************}
Function  TEvtSup.EstUnPayeur : Boolean;
begin
  Result := ExisteSQL('SELECT T_PAYEUR FROM TIERS WHERE T_PAYEUR="' + Code + '"');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Verification sur les comptes auxiliaires
Mots clefs ... :
*****************************************************************}
Function TEvtSup.EstDansPiece : Boolean;
begin

  Result := False ;

  // CA - 07/12/2005 - Cas des bases optimisées sans table PIECE
  if EstBasePclAllegee then exit;

  Case FType of

    fbAux : Result := ExisteSQL('SELECT GP_TIERS FROM ' + GetTableDossier( FDossier, 'PIECE')
                                  + ' WHERE GP_TIERS="' + GCCode
                                    + '" OR GP_TIERSLIVRE="' + GCCode
                                    + '" OR GP_TIERSFACTURE="' + GCCode + '"');

    fbJal : Result := Presence( GetTableDossier( FDossier, 'PIECE') , 'GP_JALCOMPTABLE', Code ) ;

    end ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Verification sur les comptes auxiliaires
Mots clefs ... :
*****************************************************************}
Function TEvtSup.EstDansActivite : Boolean;
begin
  // CA - 07/12/2005 - Cas des bases optimisées sans table ACTIVITE
  if not EstBasePclAllegee then
    Result := ExisteSQL('SELECT ACT_TIERS FROM ' + GetTableDossier( FDossier, 'ACTIVITE') + ' WHERE ACT_TIERS="' + GCCode + '"' )
  else Result := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Verification sur les comptes auxiliaires
Mots clefs ... :
*****************************************************************}
Function TEvtSup.EstDansRessource : Boolean;
begin
  Result := ExisteSQL('SELECT ARS_AUXILIAIRE FROM ' + GetTableDossier( FDossier, 'RESSOURCE') + ' WHERE ARS_AUXILIAIRE="' + Code + '"');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Verification sur les comptes auxiliaires
Mots clefs ... :
*****************************************************************}
Function TEvtSup.EstDansActions : Boolean;
begin
  // CA - 07/12/2005 - Cas des bases optimisées sans table ACTIONS
  if not EstBasePclAllegee then
    Result := ExisteSQL('SELECT RAC_AUXILIAIRE FROM ' + GetTableDossier( FDossier, 'ACTIONS' ) + ' WHERE RAC_AUXILIAIRE="'+Code+'"')
  else Result := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Verification sur les comptes auxiliaires
Mots clefs ... :
*****************************************************************}
Function TEvtSup.EstDansPersp : Boolean;
begin
  // CA - 07/12/2005 - Cas des bases optimisées sans table PERSPECTIVES
  if not EstBasePclAllegee then
    Result :=  ExisteSQL('SELECT RPE_AUXILIAIRE FROM ' + GetTableDossier( FDossier, 'PERSPECTIVES') + ' WHERE RPE_AUXILIAIRE="' + Code + '"')
  else Result := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Verification sur les comptes auxiliaires
Mots clefs ... :
*****************************************************************}
Function TEvtSup.EstDansAffaire : Boolean;
begin
  // CA - 07/12/2005 - Cas des bases optimisées sans table AFFAIRE
  if not EstBasePclAllegee then
    Result := ExisteSQL('SELECT AFF_TIERS FROM ' + GetTableDossier( FDossier, 'AFFAIRE') + ' WHERE AFF_TIERS="' + GCCode + '"')
  else Result := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Verification sur les comptes auxiliaires
Mots clefs ... :
*****************************************************************}
Function TEvtSup.EstDansCata : Boolean;
begin
  // CA - 07/12/2005 - Cas des bases optimisées sans table CATALOGU
  if not EstBasePclAllegee then
    Result := ExisteSQL('SELECT GCA_TIERS FROM ' + GetTableDossier( FDossier, 'CATALOGU') + ' WHERE GCA_TIERS="' + GCCode + '"')
  else Result := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Verification sur les comptes auxiliaires
Mots clefs ... :
*****************************************************************}
Function TEvtSup.EstDansPaie : Boolean;
begin
  Result := ExisteSQL('SELECT PSA_AUXILIAIRE FROM ' + GetTableDossier( FDossier, 'SALARIES') + ' WHERE PSA_AUXILIAIRE="' + Code + '"') ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Verification sur les comptes auxiliaires
Mots clefs ... :
*****************************************************************}
Function TEvtSup.EstDansMvtPaie : Boolean;
begin
  Result := ExisteSQL('SELECT PPU_SALARIE FROM ' + GetTableDossier( FDossier, 'PAIEENCOURS')
                           + ' LEFT JOIN ' + GetTableDossier( FDossier, 'SALARIES') + ' ON PPU_SALARIE=PSA_SALARIE WHERE PSA_AUXILIAIRE="' + Code + '"');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Verification sur les sections
Mots clefs ... :
*****************************************************************}
Function TEvtSup.EstDansSection : Boolean ;
begin
  Result := ExisteSQL('SELECT S_MAITREOEUVRE,S_CHANTIER FROM ' + GetTableDossier( FDossier, 'SECTION') +
                     ' WHERE S_MAITREOEUVRE="' + Code + '" OR S_CHANTIER="' + Code + '"' );
end;

Function TEvtSup.EstDansGCD : Boolean ;
begin
  Result := ExisteSQL('SELECT CGC_AUXILIAIRE FROM ' + GetTableDossier( FDossier, 'CPGCDCUMULS') +
                     ' WHERE CGC_AUXILIAIRE="' + Code + '" ') ;
end;

// modif mbo - 17/10/2006 - ajout tests sur les comptes de subvention d'investissement
// BTY 07/11/07 FQ 21792 ajout test compte est compte de charge d'un crédit-bail (i_comptelie)
function TEvtSup.EstDansAmortissement: boolean;
var
CptSub : string;
lg_cpt : integer;
begin

  Result := False ;

  Case FType of
    fbGene :
    begin
         Result := ExisteSQL ( 'SELECT I_IMMO FROM ' + GetTableDossier( FDossier, 'IMMO') + ' WHERE '+
                                      'I_COMPTEIMMO="'+Code+'" OR I_COMPTEAMORT="'+Code+'" OR I_COMPTEDOTATION="'+Code+
                                      '" OR I_COMPTEDEROG="'+Code+'" OR I_REPRISEDEROG="'+Code+'" OR I_PROVISDEROG="'+Code+
                                      '" OR I_DOTATIONEXC="'+Code+'" OR I_VACEDEE="'+Code+'" OR I_AMORTCEDE="'+Code+
                                      '" OR I_COMPTELIE="'+Code+'" OR I_CPTSBVB="'+Code+
                                      '" OR I_REPEXPLOIT="'+Code+'" OR I_REPEXCEP="'+Code+'" OR I_VAOACEDEE="'+Code+'"');

       // ajout mbo pour le compte de reprise de subvention
       if (result = false) and (copy(Code, 1,3) = '139') then
       begin
          lg_cpt := length (Code);
          CptSub := '13' + copy(Code, 4, (lg_cpt-3));

          Result := ExisteSQL ('SELECT * FROM IMMO WHERE I_CPTSBVB LIKE"'+ CptSub+'%" ');

          if Result = false then
          begin
             CptSub := '138' + copy(Code, 4, (lg_cpt-3));
             Result := ExisteSQL ('SELECT * FROM IMMO WHERE I_CPTSBVB LIKE"'+ CptSub+'%" ');
          end;
       end;

       // ajout mbo pour le compte de quote part de subvention
       if (result = false) and (copy(Code, 1,3) = '777') then
       begin
          lg_cpt := length (Code);
          CptSub := '13' + copy(Code, 4, (lg_cpt-3));

          Result := ExisteSQL ('SELECT * FROM IMMO WHERE I_CPTSBVB LIKE"'+ CptSub+'%" ');

          if Result = false then
          begin
             CptSub := '138' + copy(Code, 4, (lg_cpt-3));
             Result := ExisteSQL ('SELECT * FROM IMMO WHERE I_CPTSBVB LIKE"'+ CptSub+'%" ');
          end;
       end;
    end;
    //

    fbAux  :   Result := ExisteSQL('SELECT I_IMMO FROM ' + GetTableDossier( FDossier, 'IMMO' )
                                 + ' WHERE I_ORGANISMECB="' + Code + '"' ) ;

    //PGR 08/02/2006 Ctl code journal différent de SO_IMMOJALDOTDEF et SO_IMMOJALECHDEF
    fbJal  :   Result := ( Code = GetParamsocSecur('SO_IMMOJALDOTDEF','') ) or
                         ( Code = GetParamsocSecur('SO_IMMOJALECHDEF','') );

    end ;

end;

////////////////////////////////////////////////////////////////////////////////
// BTY 06/11/07 Ne pas supprimer un compte s'il est dans IMMOREGFR passage forfait au réel des agricoles
function TEvtSup.EstDansAgricole: boolean;
begin
  Result := False ;

  case FType of
    fbGene :
    begin
       Result := ExisteSQL ( 'SELECT IR_COMPTEIMMO FROM ' + GetTableDossier( FDossier, 'IMMOREGFR') + ' WHERE '+
                             'IR_COMPTEREF="'+Code+'" OR IR_COMPTEAMORT="'+Code+'" OR IR_COMPTEDOTATION="'+Code+
                             '" OR IR_COMPTEDEROG="'+Code+'" OR IR_REPRISEDEROG="'+Code+'" OR IR_PROVISREDOG="'+Code+
                             '" OR IR_DOTATIONEXC="'+Code+'" OR IR_AMORTCEDE="'+Code+
                             '" OR IR_REPEXPLOIT="'+Code+'" OR IR_REPEXCEP="'+Code+'"');
    end;
  end ;
end;

////////////////////////////////////////////////////////////////////////////////


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 22/07/2005
Modifié le ... :   /  /
Description .. : Test si l'enregistrement est utilisé dans les autes bases du
Suite ........ : regroupement multisociété
Mots clefs ... :
*****************************************************************}
Function TEvtSup.EstDansMultiSoc : Boolean ;
var lStGroup : String ;
    lStBase  : String ;
begin
  result := False ;
  if not EstMultiSoc then Exit ;
  if not EstTablePartagee ( fbToTable ( FType ) ) then Exit ;

  // Récupération du regroupement multisoc
  lStGroup := GetBasesMS ;
  if Trim( lStGroup ) = '' then Exit ;

  lStBase := ReadTokenSt( lStGroup ) ;

  // Parcours des bases
  While (lStBase <> '') do
    begin

    // Sur la base courante, le traitement classique a déjà été fait
    if lStBase <> V_PGI.SchemaName then 
      // Pour les autres bases, on effectue les recherches
      if EstDansBase( lStBase ) > 0 then
        begin
        result := True ;
        Exit ;
        end ;

    lStBase := ReadTokenSt( lStGroup ) ;

    end ;


end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 22/07/2005
Modifié le ... :   /  /
Description .. : Test si l'enregistrement est utilisé dans les autes bases du
Suite ........ : regroupement multisociété
Mots clefs ... :
*****************************************************************}
Function TEvtSup.EstDansBase( vBase : String ) : Integer ;
begin

  result := 0 ;

  FDossier := vBase ;

  Case FType of

    fbGene : begin
             // test si le compte est supprimable
             if EstMouvemente             then Result := 4
             else if EstEcrGuide          then Result := 17  // k
             else if EstDansAxe           then Result := 6   // ok
             else if ( not EstTablePartagee('JOURNAL') ) and EstCpteJournal  then Result := 9
             else if ( not EstTablePartagee('TIERS') ) and EstCpteCollRef    then Result := 13
             else if EstDansAmortissement then Result := 14 ;

             if result <> 0 then result := 16 ;

             end ;

    fbAux : begin
            if EstMouvemente                            then Result := 4 // ok
            else if EstEcrGuide                         then Result := 5 // ok
            // ?? Bizarre celui la a vérifier
            else if EstCpteCorresp                      then Result := 6
            // Test dans section uniquement si la table SECTION n'est pas partagée
            else if ( ( not EstTablePartagee( 'SECTION' ) ) and EstDansSection ) then Result := 8
            else if EstDansUtilisat                     then Result := 9
            else if EstDansPiece                        then Result := 10 // k
            else if EstDansActivite                     then Result := 11 // k
            else if EstDansRessource                    then Result := 12 // k
            else if EstDansAffaire                      then Result := 13
            else if EstDansActions                      then Result := 14 // k
            else if EstDansPersp                        then Result := 15 // k
            else if EstDansPaie                         then Result := 16 // k
            else if EstDansCata                         then Result := 17 // k
            else if EstDansMvtPaie                      then Result := 18 // k
            else if EstDansAmortissement                then Result := 20
            else if EstDAnsGCD                          then Result := 23 ;


            if result <> 0 then
              result := 22 ;
            end ;

    fbSect : begin
             if EstMouvemente        then Result := 5
             else result := EstVentilType ;

             if result <> 0 then
               result := 11 ;

             end ;

    fbJal : begin
            if EstMouvemente                then Result := 4
            else if EstDansSouche           then Result := 7  // ok
            else if EstDansParamPiece       then Result := 10
            else if EstDansParamPieceCompl  then Result := 11 ;

            if result <> 0 then
              result := 12 ;

            end ;

    end ;

  FDossier := '' ;

end ;


function  TEvtSup.EstCpteJournal : Boolean ;
begin
  Result := ExisteSQL('SELECT J_JOURNAL FROM ' + GetTableDossier( FDossier, 'JOURNAL') + ' WHERE J_CONTREPARTIE="'+ Code +'" '+'AND ((J_NATUREJAL="BQE") OR (J_NATUREJAL="CAI") OR (J_NATUREJAL="REG"))');
end ;

function  TEvtSup.EstCpteCollRef : Boolean ;
begin
  result := Presence( GetTableDossier( FDossier, 'TIERS') , 'T_COLLECTIF', Code );
end ;


function  TEvtSup.EstVentilType : Integer ;
var lQuery : TQuery ;
begin

  // Test si présence dans le paramétrage des ventilations types
  // GCO - 04/09/2004 - FQ 12635
  Result := 0 ;

  {JP 27/12/05 : FQ 11919 : On va préciser le type de guide et le compte associé}
  MessageErreur := '';

  try
    lQuery := OpenSql('SELECT V_NATURE, V_COMPTE FROM ' + GetTableDossier( FDossier, 'VENTIL') + ' WHERE V_SECTION = "'+ Code + '" ORDER BY V_NATURE', True);
    if not lQuery.Eof then
    begin
      if Pos('GE', UpperCase(lQuery.FindField('V_NATURE').AsString)) > 0 then begin
        {JP 27/12/05 : FQ 11919 : précision du compte associé}
        MessageErreur := 'Section utilisée dans la ventilation du compte général ' + lQuery.FindField('V_COMPTE').AsString;
        Result := 9;
      end
      else if Pos('IM', UpperCase(lQuery.FindField('V_NATURE').AsString)) > 0 then begin
        Result := 10;
        {JP 27/12/05 : FQ 11919 : précision du compte associé}
        MessageErreur := 'Section utilisée dans la ventilation de la  fiches d''immobilisation ' + lQuery.FindField('V_COMPTE').AsString;
      end
      else if Pos('TY', UpperCase(lQuery.FindField('V_NATURE').AsString)) > 0 then begin
        {JP 27/12/05 : FQ 11919 : précision du compte associé}
        MessageErreur := 'Section utilisée dans la ventilation type ' + lQuery.FindField('V_COMPTE').AsString;
        Result := 8;
      end
      else if (Pos('SA', UpperCase(lQuery.FindField('V_NATURE').AsString)) > 0) or
              (Pos('RC', UpperCase(lQuery.FindField('V_NATURE').AsString)) > 0) or
              (Pos('PG', UpperCase(lQuery.FindField('V_NATURE').AsString)) > 0) or
              (Pos('RR', UpperCase(lQuery.FindField('V_NATURE').AsString)) > 0) then begin
        {JP 27/12/05 : FQ 11919 : précision du compte associé}
        MessageErreur := 'Section utilisée dans la ventilation ' + lQuery.FindField('V_COMPTE').AsString + ' de la paie / GRH';
        Result := 8;
      end
      else if (Pos('HA', UpperCase(lQuery.FindField('V_NATURE').AsString)) > 0) or
              (Pos('HV', UpperCase(lQuery.FindField('V_NATURE').AsString)) > 0) or
              (Pos('ST', UpperCase(lQuery.FindField('V_NATURE').AsString)) > 0) or
              (Pos('EV', UpperCase(lQuery.FindField('V_NATURE').AsString)) > 0) or
              (Pos('RV', UpperCase(lQuery.FindField('V_NATURE').AsString)) > 0) or
              (Pos('EA', UpperCase(lQuery.FindField('V_NATURE').AsString)) > 0) or
              (Pos('RA', UpperCase(lQuery.FindField('V_NATURE').AsString)) > 0) or
              (Pos('VS', UpperCase(lQuery.FindField('V_NATURE').AsString)) > 0) then begin
        {JP 27/12/05 : FQ 11919 : précision du compte associé}
        MessageErreur := 'Section utilisée dans la ventilation ' + lQuery.FindField('V_COMPTE').AsString +
                         ' de Gestion commerciale et Gestion d''affaires'; {JP 21/02/06 : Suite de la FQ 11919}
        Result := 8;
      end
      else begin
        MessageErreur := 'Section utilisée dans un autre produit du PGI';
        result := 8 ;
      end;
    end;
  finally
    Ferme( lQuery );
  end;
{ FQ 19553 BVE 10.04.07 }
  if ExisteSQL('SELECT YVA_SECTION FROM VENTANA WHERE YVA_SECTION = "' + Code + '" AND YVA_AXE = "' + Codax + '"') then
  begin
     MessageErreur := 'Section utilisée dans la pré-ventilation analytique';
     result := 13 ;
  end;
{ END FQ 19553 }
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/01/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
function  TEvtSup.EstSectionAxeTva : Boolean ;
var lStAxeTva : string;
begin
  lStAxeTva := GetParamSocSecur('SO_CPPCLAXETVA', '');

  if lStAxeTva = '' then
  begin // ParamSoc pas renseigné
    lStAxeTva := GetColonneSQL('AXE', 'X_AXE', 'X_LIBELLE = "TVA"');
    SetParamSoc('SO_CPPCLAXETVA', lStAxeTva);
  end;

  Result := (Codax = lStAxeTva);
end;

////////////////////////////////////////////////////////////////////////////////

Function  TEvtSup.EstDansAxe : Boolean ;
var i : integer ;
begin
  Result := False ;

  Case FType of

    // AXE toujours partagée si SECTION partagée donc traitement spécifique non nécessaire
    fbSect : result := ( Code = VH^.Cpta[ AxeToFb(CodAx) ].Attente ) ;

    fbGene : if (FDossier='') or (FDossier=V_PGI.SchemaName) or ( EstTablePartagee('AXE') ) then
               begin
               for i:=1 to 5 do
                 if VH^.Cpta[AxeToFb('A'+IntToStr(i))].AxGenAttente = Code then
                   begin
                   Result := True;
                   Exit;
                   end
               end
             else
               result := ExisteSQL('SELECT X_AXE FROM ' + GetTableDossier( FDossier, 'AXE') + ' WHERE X_GENEATTENTE="' + Code + '"' ) ;
    end ;

end ;


////////////////////////////////////////////////////////////////////////////////

Function  TEvtSup.EstCpteCorresp : Boolean ;
begin
  Result := False ;
  Case FType of
    fbSect        : result := ExisteSQL('SELECT S_CORRESP1,S_CORRESP2 FROM ' + GetTableDossier( FDossier, 'SECTION') + ' WHERE (S_CORRESP1="'+Code+'" OR S_CORRESP2="'+Code+'") and (S_AXE="'+CodAx+'")') ;
    fbGene, fbAux : result := Presence( GetTableDossier( FDossier, 'CORRESP'), 'CR_CORRESP', Code ) ;
    end ;
end ;


////////////////////////////////////////////////////////////////////////////////
Function  TEvtSup.EstEcrGuide : Boolean ;
var
  Q : TQuery;
  s : string;
begin
  Result := False;
  (*
  Case FType of
    fbGene  : result := Presence( GetTableDossier( FDossier, 'ECRGUI'), 'EG_GENERAL',    Code ) ;
    fbAux   : result := Presence( GetTableDossier( FDossier, 'ECRGUI'), 'EG_AUXILIAIRE', Code ) ;
    end ;*)
  {JP 27/12/05 : FQ 11919 : On va préciser le guide concernée}
  case FType of
    fbGene : Q := OpenSQL('SELECT EG_TYPE, EG_GUIDE FROM ECRGUI WHERE EG_GENERAL = "' + Code + '"', True, -1, '', False, FDossier);
    fbAux  : Q := OpenSQL('SELECT EG_TYPE, EG_GUIDE FROM ECRGUI WHERE EG_AUXILIAIRE = "' + Code + '"', True, -1, '', False, FDossier);
  end;

  try
    if not Q.EOF then begin
           if Q.FindField('EG_TYPE').AsString = 'ABO' then s := 'd''abonnement'
      else if Q.FindField('EG_TYPE').AsString = 'DEC' then s := 'de règlement'
      else if Q.FindField('EG_TYPE').AsString = 'ENC' then s := 'de règlement'
      else if Q.FindField('EG_TYPE').AsString = 'NOR' then s := 'de saisie'
      else if Q.FindField('EG_TYPE').AsString = 'POI' then s := 'de pointage';
      Result := True;
      MessageErreur := 'Le compte est utilisé dans le guide ' + s + ' (' + Q.FindField('EG_GUIDE').AsString + ')';
    end;
  finally
    Ferme(Q);
  end;
end ;

////////////////////////////////////////////////////////////////////////////////


Function TEvtSup.EstDansSouche : Boolean ;
begin
  result := ExisteSQL( 'SELECT SH_SOUCHE FROM ' + GetTableDossier( FDossier, 'SOUCHE' ) + ' WHERE SH_JOURNAL ="' + Code + '"' ) ;
end ;

Function TEvtSup.EstDansUtilisat : Boolean ;
begin
  result := Presence( GetTableDossier( FDossier, 'UTILISAT' ),'US_AUXILIAIRE', Code) ;
end ;

function TEvtSup.EstDansParamPiece      : Boolean;
begin
  // CA - 07/12/2005 - Cas des bases optimisées sans table PIECE
  if not EstBasePclAllegee then
    result := Presence( GetTableDossier( FDossier, 'PARPIECE' ),'GPP_JOURNALCPTA', Code)
  else result := False;
end;

function TEvtSup.EstDansParamPieceCompl : Boolean;
begin
  // CA - 07/12/2005 - Cas des bases optimisées sans table PIECE
  if not EstBasePclAllegee then
    result := Presence( GetTableDossier( FDossier, 'PARPIECECOMPL' ),'GPC_JOURNALCPTA', Code)
  else result := false;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 30/10/2006
Modifié le ... :   /  /
Description .. : Recherche du code journal dans les guides
Mots clefs ... : FQ19030
*****************************************************************}
function TEvtSup.EstDansGuide : Boolean;
begin
  Result := ExisteSQL('SELECT GU_GUIDE FROM GUIDE WHERE GU_JOURNAL="'+Code+'"');
end ;

function TEvtSup.EstCpteModepaie      : Boolean ;
begin
  Result := Presence( GetTableDossier( FDossier, 'MODEPAIE' ), 'MP_GENERAL', Code ) ;
end ;

function TEvtSup.EstCpteTva           : Boolean ;
begin
  Result := EstTvaTpf( Code, True ) ;
  if Not Result then Result := EstTvaTpf( Code, False ) ;
end ;

function TEvtSup.EstCpteDevise        : Boolean ;
begin
  Result := ExisteSQL('SELECT D_DEVISE FROM DEVISE WHERE(D_CPTLETTRDEBIT="'+Code+'") OR (D_CPTLETTRCREDIT="'+Code+'") '+'OR (D_CPTPROVDEBIT="'+Code+'") OR (D_CPTPROVCREDIT="'+Code+'")');
end ;

{JP 27/12/05 : FQ 11919 : Test sur la table des paramètres compta GC
{---------------------------------------------------------------------------------------}
function TEvtSup.EstCodeCptaGC : Boolean;
{---------------------------------------------------------------------------------------}
begin
  // CA - 07/12/2005 - Cas des bases optimisées sans table PIECE
  if not EstBasePclAllegee then
    Result := ExisteSQL('SELECT * FROM ' + GetTableDossier(FDossier, 'CODECPTA') + ' WHERE ' +
                      'GCP_CPTEGENEACH = "' + Code + '" OR ' +
                      'GCP_CPTEGENESCACH = "' + Code + '" OR ' +
                      'GCP_CPTEGENESCVTE = "' + Code + '" OR ' +
                      'GCP_CPTEGENESTOCK = "' + Code + '" OR ' +
                      'GCP_CPTEGENEVARSTK = "' + Code + '" OR ' +
                      'GCP_CPTEGENEVTE = "' + Code + '" OR ' +
                      'GCP_CPTEGENREMACH = "' + Code + '" OR ' +
                      'GCP_CPTEGENREMVTE = "' + Code + '" ')
  else Result := False;
end;

////////////////////////////////////////////////////////////////////////////////

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/04/2007
Modifié le ... :   /  /    
Description .. : - LG - 27/04/2007 - FB 20146 - on controle la presence des 
Suite ........ : champs ds GCD
Mots clefs ... : 
*****************************************************************}
Function TEvtSup.EstCompteSupprimable : Integer ;
begin

  result := 0 ;
  MessageErreur := ''; {JP 27/12/05 : FQ 11919}

  Case FType of

    fbGene :  begin
              // test si le compte est supprimable
              if EstMouvemente                 then Result := 4
              else if EstEcrGuide              then Result := 17 // k
              else if EstDansAxe               then Result := 6
//              else if EstCpteCorresp           then Result := 7
              else if EstCpteDevise            then Result := 8
              else if EstCpteJournal           then Result := 9
              else if EstCpteModepaie          then Result := 10
              else if EstCpteTva               then Result := 12
              else if EstCpteCollRef           then Result := 13
              else if EstDansAmortissement     then Result := 14
              else if EstDansSociete           then Result := 15
              else if EstDansMultiSoc          then result := 16
              else if EstCodeCptaGC            then Result := 18 {JP 27/12/05 : FQ 11919}
              else if EstTableauVar            then Result := 19
              else if EstNoteTravail           then Result := 20
              else if EstDansAgricole          then Result := 21 {BTY 06/11/07 Agricole}

              ;
              end ;

    fbAux : begin
            // test si le compte est supprimable
            if      EstMouvemente                                then Result := 4
            else if EstEcrGuide                                  then Result := 5 // k
//            else if EstCpteCorresp                               then Result := 6
            else if EstDansSection                               then Result := 8
            else if EstDansUtilisat                              then Result := 9
            else if EstUnPayeur                                  then Result := 19
            else if EstDansPiece                                 then Result := 10 // k
            else if EstDansActivite                              then Result := 11 // k
            else if EstDansRessource                             then Result := 12 // k
            else if EstDansAffaire                               then Result := 13
            else if EstDansActions                               then Result := 14 // k
            else if EstDansPersp                                 then Result := 15 // k
            else if EstDansPaie                                  then Result := 16 // k
            else if EstDansCata                                  then Result := 17 // k
            else if EstDansMvtPaie                               then Result := 18 // k
            else if EstDansAmortissement                         then Result := 20
            else if EstDansSociete                               then Result := 21
            else if EstDansMultiSoc                              then result := 22
            else if EstDAnsGCD                                   then Result := 23 ;
            end ;

    fbSect : begin
             // test si le compte est supprimable
             if EstMouvemente                                    then Result := 5
             else if EstDansAxe                                  then Result := 6
//             else if EstCpteCorresp                              then Result := 7
             else result := EstVentilType ;

             if result = 0 then
               if EstDansMultiSoc                                then result := 11 ;

             // GCO - 30/01/2006 - FQ 17216
             if result = 0 then
             begin
               if GetParamSocSecur('SO_CPPCLSAISIETVA', False)  then
               begin
                 if EstSectionAxeTva                             then result := 12;
               end;
             end;


             end ;

    fbJal : begin
            // test si le compte est supprimable
            if EstMouvemente                                     then Result := 4
            else if EstDansSociete                               then Result := 6
            else if EstDansSouche                                then Result := 7
            else if EstDansParamPiece                            then Result := 10
            else if EstDansParamPieceCompl                       then Result := 11
            else if EstDansMultiSoc                              then result := 12
            //PGR 08/02/2006 Ctl code journal différent de SO_IMMOJALDOTDEF et SO_IMMOJALECHDEF
            else if EstDansAmortissement                         then Result := 6
            else if EstDansGuide                                 then Result := 13;
            end ;

    end ;


end ;

////////////////////////////////////////////////////////////////////////////////

Function  TEvtSup.EstDansSociete       : Boolean ;
begin
  Result := False ;
  Case FType of
  //FQ19029 YMO 31/10/2006 Recherche dans les paramètres société
  //correspondant aux tablettes de domaine compta et de préfixe journal
    fbJal  : Result := ExisteSQL('SELECT SOC_NOM FROM PARAMSOC, DECOMBOS '
                     +'WHERE SOC_DATA="'+Code+'" AND SOC_DESIGN LIKE "%"||DO_COMBO||"%" '
                     +'AND DO_DOMAINE="C" AND DO_PREFIXE="J"');
   {
    result := ( Code = GetParamsocSecur('SO_JALOUVRE','') )      or
                         ( Code = GetParamsocSecur('SO_JALFERME','') )      or
                         ( Code = GetParamsocSecur('SO_JALVTP','') )        or
                         ( Code = GetParamsocSecur('SO_JALATP','') )        or
                         ( Code = GetParamsocSecur('SO_JALREPBALAN','') ) ;
   }

    fbGene :   Result :=  (Code = GetParamSocSecur('SO_CPTEAMORTINF','') ) or
                          (Code = GetParamSocSecur('SO_CPTEAMORTSUP','') ) or
                          (Code = GetParamSocSecur('SO_CPTECBSUP','') )    or
                          (Code = GetParamSocSecur('SO_CPTEDEPOTINF','') ) or
                          (Code = GetParamSocSecur('SO_CPTEDEPOTSUP','') ) or
                          (Code = GetParamSocSecur('SO_CPTEDEROGINF','') ) or
                          (Code = GetParamSocSecur('SO_CPTEDEROGSUP','') ) or
                          (Code = GetParamSocSecur('SO_CPTEDOTEXCINF','') ) or
                          (Code = GetParamSocSecur('SO_CPTEDOTEXCSUP','') ) or
                          (Code = GetParamSocSecur('SO_CPTEDOTINF','') )    or
                          (Code = GetParamSocSecur('SO_CPTEDOTSUP','') )    or
                          (Code = GetParamSocSecur('SO_CPTEEXPLOITINF','') ) or
                          (Code = GetParamSocSecur('SO_CPTEEXPLOITSUP','') ) or
                          (Code = GetParamSocSecur('SO_CPTEFININF','') )     or
                          (Code = GetParamSocSecur('SO_CPTEFINSUP','') )     or
                          (Code = GetParamSocSecur('SO_CPTEIMMOINF','') )    or
                          (Code = GetParamSocSecur('SO_CPTEIMMOSUP','') )    or
                          (Code = GetParamSocSecur('SO_CPTELOCINF','') )     or
                          (Code = GetParamSocSecur('SO_CPTELOCSUP','') )     or
                          (Code = GetParamSocSecur('SO_CPTEPROVDERINF','') ) or
                          (Code = GetParamSocSecur('SO_CPTEPROVDERSUP','') ) or
                          (Code = GetParamSocSecur('SO_CPTEREPDERINF','') )  or
                          (Code = GetParamSocSecur('SO_CPTEREPDERSUP','') )  or
                          (Code = GetParamSocSecur('SO_CPTEREPEXCINF','') )  or
                          (Code = GetParamSocSecur('SO_CPTEREPEXCSUP','') )  or
                          (Code = GetParamSocSecur('SO_CPTEVACEDEEINF','') ) or
                          (Code = GetParamSocSecur('SO_CPTEVACEDEESUP','') ) or
                          (Code = GetParamSocSecur('SO_DEFCOLCDIV','') )     or
                          (Code = GetParamSocSecur('SO_DEFCOLCLI','') )      or
                          (Code = GetParamSocSecur('SO_DEFCOLDDIV','') )     or
                          (Code = GetParamSocSecur('SO_DEFCOLDIV','') )      or
                          (Code = GetParamSocSecur('SO_DEFCOLFOU','') )      or
                          (Code = GetParamSocSecur('SO_DEFCOLSAL','') )      or
                          (Code = GetParamSocSecur('SO_GENATTEND','') )      or
                          (Code = GetParamSocSecur('SO_FERMEBEN','') )       or
                          (Code = GetParamSocSecur('SO_FERMEBIL','') )       or
                          (Code = GetParamSocSecur('SO_FERMEPERTE','') )     or
                          (Code = GetParamSocSecur('SO_OUVREBEN','') )       or
                          (Code = GetParamSocSecur('SO_OUVREBIL','') )       or
                          (Code = GetParamSocSecur('SO_OUVREPERTE','') )     or
                          (Code = GetParamSocSecur('SO_PRODEB1','') )        or
                          (Code = GetParamSocSecur('SO_PRODEB2','') )        or
                          (Code = GetParamSocSecur('SO_PRODEB3','') )        or
                          (Code = GetParamSocSecur('SO_PRODEB4','') )        or
                          (Code = GetParamSocSecur('SO_PRODEB5','') )        or
                          (Code = GetParamSocSecur('SO_PROFIN1','') )        or
                          (Code = GetParamSocSecur('SO_PROFIN2','') )        or
                          (Code = GetParamSocSecur('SO_PROFIN3','') )        or
                          (Code = GetParamSocSecur('SO_PROFIN4','') )        or
                          (Code = GetParamSocSecur('SO_PROFIN5','') )        or
                          (Code = GetParamSocSecur('SO_RESULTAT','') )       or
                          (Code = GetParamSocSecur('SO_LETCHOIXGEN','') )    or
                          (Code = GetParamSocSecur('SO_LETCHOIXGENC','') )   or
                          (Code = GetParamSocSecur('SO_ICCCOMPTECAPITAL','') ) or
                          (Code = GetParamSocSecur('SO_TELETVABQE1','') )      or
                          (Code = GetParamSocSecur('SO_TELETVABQE2','') )      or
                          (Code = GetParamSocSecur('SO_TELETVABQE3','') )      or
                          // DOMAINE PAIE (P)
                          (Code = GetParamSocSecur('SO_PGCPTNETAPAYER','') ) or
                          // DOMAINE GESCOM (G)
                          (Code = GetParamSocSecur('SO_GCCPTEESCACH','') ) or
                          (Code = GetParamSocSecur('SO_GCCPTEESCVTE','') ) or
                          (Code = GetParamSocSecur('SO_GCCPTEREMACH','') ) or
                          (Code = GetParamSocSecur('SO_GCCPTEESCVTE','') ) or
                          (Code = GetParamSocSecur('SO_GCCPTEHTACH','') )  or
                          (Code = GetParamSocSecur('SO_GCCPTEHTVTE','') )  or
                          (Code = GetParamSocSecur('SO_GCCPTEHTVTE','') )  or
                          (Code = GetParamSocSecur('SO_GCCPTEHTVTE','') )  or
                          (Code = GetParamSocSecur('SO_GCCPTEPORTACH','') ) or
                          (Code = GetParamSocSecur('SO_GCCPTEPORTVTE','') ) or
                          (Code = GetParamSocSecur('SO_GCCPTERGVTE','') )   or
                          (Code = GetParamSocSecur('SO_GCCPTESTOCK','') )   or
                          (Code = GetParamSocSecur('SO_GCCPTEVARSTK','') );

    fbAux  :   Result :=  (Code = GetParamSocSecur('SO_CLIATTEND','') ) or
                          (Code = GetParamSocSecur('SO_FOUATTEND','') ) or
                          (Code = GetParamSocSecur('SO_SALATTEND','') ) or
                          (Code = GetParamSocSecur('SO_DIVATTEND','') ) or
                          // Domaine (GC)
                          (Code = GetParamSocSecur('SO_GCFOUCPTADIFF','') ) or
                          (Code = GetParamSocSecur('SO_GCCLICPTADIFF','') ) or
                          (Code = GetParamSocSecur('SO_GCFOUCPTADIFFPART','') ) or
                          (Code = GetParamSocSecur('SO_GCCLICPTADIFFPART','') );

    end ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/11/2006
Modifié le ... :   /  /    
Description .. : Test si le compte à supprimer est associé 
Suite ........ : à un tableau des variations
Mots clefs ... : 
*****************************************************************}
function TEvtSup.EstTableauVar : Boolean;
begin
 Result := ExisteSQL('SELECT 1 FROM CPTABLEAUVAR WHERE CTV_GENERAL="'+Code+'"');
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 20/11/2006
Modifié le ... :   /  /    
Description .. : test si une note de travail est associé avec le compte
Mots clefs ... :
*****************************************************************}
function  TEvtSup.EstNoteTravail       : Boolean;
begin
  Result := ExisteSQL('SELECT 1 FROM CPNOTETRAVAIL WHERE CNO_GENERAL="'+Code+'"');
end;
////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////
//================================================================================
// Initialization
//================================================================================
{ TEvtSup }

{ TEvtSup }
Initialization
    RegisterAglProc('SupprimeListeCpte',TRUE,1,AGLSupprimeListCpte);
end.
