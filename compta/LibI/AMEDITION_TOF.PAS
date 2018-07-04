{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 24/02/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AMEDITION ()
Mots clefs ... : TOF;AMEDITION

Suite ........ : MBO - 08/09/2005 - FQ 14430 Liste mutations avec uniquement mut compte
                 (ajout xx_where sur fiche amedition)
Suite......... : MBO - 18/11/2005 - ajout liste des dépréciations/Reprises
                 TGA - 07/12/2005 - ajout liste des changements de conditions d'amortissement
                 BTY - 12/2005    - FQ 15243 Ajout check-box Commentaire de l'immo
                 PGR - 27/12/2005 - FQ 15637 - ajout liste des valorisations
                 BTY - 01/06      - FQ 17172 - les immos ne s'impriment pas dans l'ordre des codes immos i_immo
                 MBO - 13/01/2006 - FQ 13036 - ajout état des engagements
                 MBO - 17/01/2006 - FQ 17315 - ajout état dépréciations pour S1 + tables 0 à 9 dans le select
                                             + ajout état des changements de conditions
                 PGR - 17/01/2006 - ajout état des valorisations pour S1
                 Tga - 09/02/2006 - suppression interdiction des éditions sous Oracle
                 BTY - 02/06      - FQ 13035 - ajout état des dotations simulées pour S5 et S1 et CWAS
                 BTY - 02/06      - Mauvais fonctionnement rupture sur Regroupement pour les engagements
                 TGA - 16/02/2006 - ajout liste des amortissements variables
                 MVG - 13/04/2006 - correction ordre SQL pour SERIE1
                 MBO - 19/06/2006 - FQ 13036 - zone date arrêté non modifiable si etat engagt et dot simulées
                 MVG 17/08/2006 FQ 18636 - ajout BLOCNOTE
                 MVG 18/08/2006 FQ 18637 - ajout BLOCNOTE et modification fonctionnement de l'état d'inventaire pour CBL
                 TGA 13/09/2006 Gestion des restrictions utilsateurs sur les établissements
                 BTY 09/06 Restrictions utilisateur non respectées en restauration de filtre
                 BTY 09/06 FQ 18828 En S1, les critères et ruptures de l'utilisateur ne fonctionnent pas :
                     exemple une édition demandée pour l'immo 001 édite toutes les immos
                 BTY 09/06 FQ 18860 En S1, Répartition des dotations ne présentent pas les amortissements variables
                 BTY 09/06 FQ 18891 En S1, Taxe Professionnelle n'afifche aucun montant + aucune Nature de Bien
                 MVG 30/10/06 FQ 19048 + maquette IEG
                 BTY 11/06 FQ 19036 Divers modifs sur la Taxe professionnelle
                 MBO 03/07 FQ 17512 chantier fiscal
                 MVG 05/07 FQ 19798 Paramètrage des états : En Entreprise, tjs possible. En cabinet, si l'on est avec le mot de passe du jour
                 BTY 06/07 FQ 19927 Avec un filtre par défaut ayant un modèle de lettre d'inventaire paramétrée, on édite quand même la lettre standard Cegid
                 MBO 24/07/2007 FQ 19171 - pour liste sorties et +/- values : ajout test sur date d'arrêté
                 BTY 09/07 Lettre d'inventaire IVL : se positionner dans l'onglet Mise en page et griser certains champs
                 MBO 17/09/07 FQ 21422 - IRE : modif du select : suppression (i_nonded<>"X" et méthode=VAR)
                 MBO 18/09/07 FQ 19913 - Autoriser une date édition <> fin d'exo pour dotations simulées (IDS)
                 BTY 13/11/07 Nouvel état IFR des passages forfaits à réel des biens agricoles
                 BTY 21/11/07 FQ 21899 IDD (DPI agricoles) : agrandir l'écran + lancer l'état automatiquement (sans faire F9)
                 BTY 21/11/07 FQ 18889 IDD (DPI agricoles) : présenter trié sur iz_date d'attribution de la DPI
*****************************************************************}
Unit AMEDITION_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     MaineAgl,
     eQRS1,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     fe_main,
     QRS1,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
{$IFDEF EAGLCLIENT}
     utileAGL,
     uLanceProcess,
     uHTTP,
{$ELSE}
     uEdtComp,
{$ENDIF}
     LookUp,
     uTOB,
     {$IFDEF SERIE1}
     {$ELSE}
     TofMeth,
     CritEdt,
     {$ENDIF}
     AGLInit,
     LicUtil
  {$IFDEF SERIE1}
  {$ELSE}
     , Ent1
  {$ENDIF}
     ,HTB97;

Type
  {$IFDEF SERIE1}
  TOF_AMEDITION = Class (TOF)
  {$ELSE}
  TOF_AMEDITION = Class (TOF_METH)
  {$ENDIF}
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    {$IFNDEF SERIE1}
    procedure OnChangeFiltre(Sender : TObject ); override ;
    {$ENDIF}
    private
      FCodeEtat : string;
      FNatureEtat : string;
      FTRupture : TOB;

      // TGA 12/07/2006
      // FQ 21899 Fvalidation: TNotifyEvent;

     // BTY 09/06 Restrictions utilisateur pour les filtres aussi
     // // TGA 19/13/2006 =========================================
     // {$IFNDEF SERIE1}
     // FComboEtab : THMultiValComboBox;
     // function  GetComboEtab : THMultiValComboBox; //gestion des réstrictions Etablissements et à défaut des ParamSoc}
     // procedure GereEtablissement; //on s'assure que le filtre coincide avec les restrictions utilisateurs sur l'établissement
     // procedure ControlEtab;
     // property  ComboEtab : THMultiValComboBox read GetComboEtab write FComboEtab;
     // {$ENDIF}
     // // ========================================================

      procedure OnChangeRupture(Sender: TObject) ;
      // mbo 19.01.06 - correction warning de compil : déplacement sauf pour S1 + ajout override
      {$IFDEF SERIE1}
      procedure OnChangeFiltre(Sender : TObject );
      {$ENDIF}
      procedure AssignHelpContext;
      procedure OnClickCompteRef ( Sender : TObject );
      procedure OnChangeNature ( Sender : TObject );
      procedure OnChangeEtat ( Sender : TObject );
      procedure InitZonesSuivantEtat;
      procedure InitRuptures;
      procedure ChargeListeRupture;
      procedure RenseigneComboRupture ( CB : THValComboBox; stSauf : string );
      procedure MajLookUpSection ( stAxe : string );
      procedure OnExitCompteRef ( Sender : TObject );
      procedure OnExitDateAchatFin ( Sender : TObject );  // FQ 19036
      procedure OnExitDateArrete ( Sender : TObject );  // FQ 19036
      procedure OnClickRegroupement ( Sender : TObject );
      // ajout mbo FQ 14430
      procedure OnClickMutcompte (Sender :TObject );
      {$IFDEF SERIE1}
      procedure OnChangeChampLibre(Sender: TObject);
    procedure OnNew;
      {$ENDIF}
  end ;

procedure AMLanceFiche_Edition ( Nature, Code : string );
{$IFDEF SERIE1}
{$ELSE}
procedure AMLanceFicheEtatChaine( vCritEdtChaine: TCritEdtChaine );
{$ENDIF}


Implementation

uses
    {$IFDEF MODENT1}
    CPProcMetier,
    CpTypeCons,
    {$ENDIF MODENT1}
      {$IFDEF SERIE1}
      S1Util, //YCP 25/08/05
      Ut0,    // BTY 09/09 pour IIF
      {$ELSE}
      CalcOle,
      {$ENDIF}
      ImOle,
      ImEdCalc,
      ImEnt;  // AMDatePassageREGFR

{$IFDEF EAGLCLIENT}
function CalcAmortServer(sf,sp : string) : variant ;
var T, TResponse : TOB;
begin
  T := TOB.Create ('',nil,-1);
  try
    { Ouverture du serveur }
    TResponse := AppServer.Request('Amortissement.AmEdOpen','',T,'','');
    if TResponse = nil then PGIBox('Amortissement ( Ouverture ) : serveur d''éditions introuvable');
    if TResponse<> nil then TResponse.Free;
  finally
    T.Free;
  end;
end;
{$ENDIF}

procedure AMLanceFiche_Edition ( Nature, Code : string );
var stMsg : string;
begin
  // TGA 09/02/2006
  //if (((Nature='ITE') or (Nature='ITF') or (nature='ITR') or (nature='IRE')) and (V_PGI.Driver = dbORACLE7))
  //  then PGIInfo( 'Cet état n''est pas utilisable sous Oracle 7. Veuillez utiliser l''état correspondant dans le menu Anciennes éditions' , 'Editions' )
  //else AGLLanceFiche ('AM','AMEDITION','',Code,Nature);

  // BTY 13/11/07 En eaglclient si on vient des états chainés, TheData n'est pas remis à nil
  // => empêche l'action du Timer sur le OnNew si on lance IFR/IDD
  TheData := nil;
  // BTY 13/11/07 Conditionner la Déclaration Agricole à la date de passage forfait à réel
  if not AMDatePassageREGFR (Nature, Code, stMsg) then
     PGIBox(stMsg)
  else
     AGLLanceFiche ('AM','AMEDITION','',Code,Nature);
end;

{$IFDEF SERIE1}
{$ELSE}
procedure AMLanceFicheEtatChaine( vCritEdtChaine: TCritEdtChaine );
var lCritEdtChaine : ClassCritEdtChaine;
    stMsg : string;
begin
  lCritEdtChaine := ClassCritEdtChaine.Create;
  lCritEdtchaine.CritEdtChaine := vCritEdtChaine;
  TheData := lCritEdtChaine;

  // BTY 13/11/07 Conditionner la Déclaration Agricole à la date de passage forfait à réel
  if not AMDatePassageREGFR (vCritEdtChaine.NatureEtat, vCritEdtChaine.CodeEtat, stMsg) then
     V_PGI.NoPrintDialog := True
  else
     AGLLanceFiche('AM', vCritEdtChaine.NomFiche, '', vCritEdtChaine.CodeEtat, vCritEdtChaine.NatureEtat );
end;
{$ENDIF}

procedure TOF_AMEDITION.OnNew ;
begin
  Inherited ;
{$IFDEF SERIE1}
TFQRS1(Ecran).FEtat.Value:=FCodeEtat;
TFQRS1(Ecran).ReloadComboEtat() ;
{$ELSE}

  // BTY 13/11/07 Lancer IFR automatiquement si pas états chaînés
  // FQ 21889
  if (TheData = nil) and ((FNatureEtat = 'IFR') or (FNatureEtat = 'IDD')) then
  begin
    TToolBarButton97(GetControl('BAGRANDIR',False)).Click; // agrandit l'écran
    FTimer.Enabled := True;  // lance le Timer de l'ancêtre TOF_METH qui lance l'édition
  end;
{$ENDIF SERIE1}
end ;

procedure TOF_AMEDITION.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AMEDITION.OnUpdate ;
{$IFDEF SERIE1}
var
  stSQL : string;
  stChampsIMMO, stChampsGENERAL : string;
  RecupWhereSQL, stRecupWhere, stRecupOrderBy, stWhere : string;  // BTY 09/06
{$ENDIF}

(*

  PLUS TARD POUR CALCUL CORRECT SUR N-1 *)
(*var Exo : TExoDate;
    DateArrete : TDateTime;*)

begin
(*  DateArrete := StrToDate(GetControlText( 'DATEARRETE'));
  if ImQuelDateDeExo(ImQuelExoDt(DateArrete), Exo) then
  begin
    SetControlText ( 'utoDATEDEBEXO', DateToStr(Exo.Deb));
    SetControlText ( 'DATEFINEXO',DateToStr(Exo.Fin));
  end else
  begin
    SetControlText ( 'DATEDEBEXO', DateToStr(VHImmo^.Encours.Deb));
    SetControlText ( 'DATEFINEXO',DateToStr(VHImmo^.Encours.Fin) );
  end;*)
  Inherited ;
{$IFDEF SERIE1}
  // BTY FQ 18828
  // Appliquer les choix multicritères renseignés par l'utilisateur :
  // les données complètent le début du WHERE
  // les ruptures redéfinies dans InitRuptures font le ORDER BY
  // => inutile d'ajouter un order by spécifique par état
  RecupWhereSQL := TFQRS1(Ecran).WhereSQL;
  stRecupWhere := ' WHERE ' + ReadTokenPipe (RecupWhereSQL, '|');
  stRecupOrderBy := ' ORDER BY ' + RecupWhereSQL; //ReadTokenPipe (RecupWhereSQL, '|');

  // mbo - FQ 17315 - 17.01.2006 ajout des tables de 0 à 9
  stChampsIMMO := 'I_COMPTEREF, I_IMMO, I_LIBELLE,I_DATEPIECEA, I_MONTANTHT, I_LIEUGEO,I_GROUPEIMMO,'+
                  'I_NATUREIMMO,I_TIERSA,I_METHODEECO,I_DUREEECO,I_TAUXECO,I_METHODEFISC,I_ETABLISSEMENT,'+
                  'I_VENTILABLE, I_TVARECUPEREE,I_TVARECUPERABLE,I_QUANTITE, I_BASEECO,I_TAUXFISC,'+
                  'I_DUREEFISC,I_BASEFISC, I_PLANACTIF,I_COMPTEDOTATION,I_QTCEDE, I_IMMOLIE, I_OPEMUTATION,'+
                  'I_OPEECLATEMENT,I_OPECESSION,I_REPRISEECO,I_OPECHANGEPLAN, I_OPELIEUGEO, I_OPEETABLISSEMENT,'+
                  'I_OPELEVEEOPTION, I_CHANGECODE,  I_OPEMODIFBASES,I_OPERATION, I_MONTANTBASEAMORT,I_MONTANTEXC,'+
                  'I_REPRISEFISCAL,I_REINTEGRATION,I_QUOTEPART,I_BASETAXEPRO,I_REVISIONECO,'+
                  'I_TABLE0,I_TABLE1,I_TABLE2,I_TABLE3,I_TABLE4,I_TABLE5,I_TABLE6,I_TABLE7,I_TABLE8,I_TABLE9';

  if  (FNatureEtat='IEG') then  // BTY 02/06 Engagements
      stChampsIMMO := stChampsIMMO + ',I_COMPTELIE,I_RESIDUEL,I_COMPTEIMMO'; // MVG FQ 18504

  if  (FNatureEtat='IDS') then  // BTY 02/06 FQ 13035 - Dotations simulées
      stChampsIMMO := stChampsIMMO + ',I_COMPTELIE,I_DEPOTGARANTIE,I_COMPTEIMMO'; // MVG FQ 18504

  if  (FNatureEtat='ITP') then  // BTY 09/06 FQ 18891
      stChampsIMMO := stChampsIMMO + ',I_NATUREBIEN,I_COMPTEIMMO,I_DATEFINCB';  // BTY FQ 19036

  if  ((FNatureEtat='IRD') OR (FNatureEtat = 'ITR') ) then // BTY 13/11/07 ) then oubliés // FQ 21422 MBO 17.09.07 OR (FNatureEtat ='IRE')) then  // MBO 03/07 FQ 17512
      stChampsIMMO := stChampsIMMO + ',I_NONDED';

  //BTY 13/11/07 Nouvel état IFR des passages forfaits à réel des biens agricoles
  if  (FNatureEtat='IFR') then
      stChampsIMMO := stChampsIMMO + ',I_ACHFOR,I_PRIXACFOR,I_PRIXACFORC,I_BLOCNOTE,I_PFR';

  stChampsGENERAL := ',G_GENERAL,G_LIBELLE';
  if (FNatureEtat='ILD') then  // FQ 17315 - mbo - 17.01.2006 -ajout état Dépréciations
  begin
    // MVG 07/07/06 FQ 18505
    stSQL := 'SELECT '+stChampsIMMO+stChampsGENERAL;
    stSQL := stSQL + ' FROM IMMO';
    stSQL := stSQL + ' LEFT JOIN IMMOLOG ON (I_IMMO=IL_IMMO)';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON G_GENERAL=I_COMPTEREF';
    stWhere := stRecupWhere + ' AND (I_ETAT<>"FER") AND IL_TYPEOP="DPR" AND(I_REVISIONECO<>0)' +
                     ' AND IL_DATEOP >= "'+ USDateTime(StrToDate(GetControlText('DATEDEBEXO')))+'"'+
                     ' AND IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"';
    // FQ 18828 stSQL := stSQL + ' ORDER BY I_IMMO';
    // Fin MVG 07/07/06

  end else if (FNatureEtat='ILC') then  //FQ 17316 - mbo - 17.01.2006 - chgt de conditions
  begin
    // Début MVG 07/07/06 FQ 18506
    stSQL := 'SELECT '+stChampsIMMO+stChampsGENERAL+
             ',IL_DATEOPREELLE,IL_METHODEECO,IL_REPRISEECO,IL_REPRISEFISC,IL_MONTANTDOT,IL_DUREEECO';
    stSQL := stSQL + ' FROM IMMO';
    stSQL := stSQL + ' LEFT JOIN IMMOLOG ON (I_IMMO=IL_IMMO) ';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON G_GENERAL=I_COMPTEREF';
    // FQ 18828 stSQL := stSQL + ' WHERE I_ETAT<>"FER"'  +
    stWhere := stRecupWhere + ' AND I_ETAT<>"FER"'  +
                     ' AND IL_TYPEOP="CDA" AND IL_DATEOP >= "'+ USDateTime(StrToDate(GetControlText('DATEDEBEXO')))+'"'+
                     ' AND IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"';
    // FQ 18828 stSQL := stSQL + ' ORDER BY I_IMMO';
    // fin MVG 07/07/06
  end else if (FNatureEtat='IEG') then  //FQ 17306 - mbo - 19.01.2006 - Liste des engagements
  begin
    stSQL := 'SELECT '+stChampsIMMO;
    stSQL := stSQL + ',G1.G_LIBELLE LIB1,G1.G_GENERAL,G2.G_LIBELLE LIB2,G2.G_GENERAL';
    stSQL := stSQL + ' FROM IMMO';
    stSQL := stSQL + ' LEFT JOIN IMMOLOG ON I_IMMO=IL_IMMO';
    stSQL := stSQL + ' LEFT JOIN GENERAUX G1 ON G1.G_GENERAL=I_COMPTELIE'; // MVG 18504
    stSQL := stSQL + ' LEFT JOIN GENERAUX G2 ON G2.G_GENERAL=I_COMPTEIMMO'; // MVG 18504
    // FQ 18828 stSQL := stSQL + ' WHERE (I_ETAT<>"FER") AND (I_NATUREIMMO ="CB")' +  // MVG 18504
    stWhere := stRecupWhere + ' AND (I_ETAT<>"FER") AND (I_NATUREIMMO ="CB" OR I_NATUREIMMO="LOC")' +
                     ' AND IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"';
    // FQ 18828 stSQL := stSQL + ' ORDER BY I_COMPTELIE,I_COMPTEIMMO,I_IMMO';

  end else if (FNatureEtat='IDS') then
  begin
    stSQL := 'SELECT '+stChampsIMMO;
    stSQL := stSQL + ',G1.G_LIBELLE LIB1,G1.G_GENERAL,G2.G_LIBELLE LIB2,G2.G_GENERAL';
    stSQL := stSQL + ' FROM IMMO';
    stSQL := stSQL + ' LEFT JOIN IMMOLOG ON I_IMMO=IL_IMMO';
    stSQL := stSQL + ' LEFT JOIN GENERAUX G1 ON G1.G_GENERAL=I_COMPTEREF'; // MVG 18504
    stSQL := stSQL + ' LEFT JOIN GENERAUX G2 ON G2.G_GENERAL=I_COMPTEIMMO'; // MVG 18504s
    // FQ 18828 stSQL := stSQL + ' WHERE (I_ETAT<>"FER") AND (I_NATUREIMMO ="CB")' + // MVG 18504
    stWhere := stRecupWhere + ' AND (I_ETAT<>"FER") AND (I_NATUREIMMO ="CB")' +
                     ' AND IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"';
    // FQ 18828 stSQL := stSQL + ' ORDER BY I_COMPTEIMMO,I_IMMO';

  end else if (FNatureEtat='ILV') then  // PGR - 17/01/2006 - ajout état des valorisations pour S1
  begin
    stSQL := 'SELECT '+stChampsIMMO+stChampsGENERAL;
    stSQL := stSQL + ' FROM IMMO ';
    stSQL := stSQL + ' LEFT JOIN IMMOLOG ON (IMMO.I_IMMO=IMMOLOG.IL_IMMO)';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON GENERAUX.G_GENERAL=IMMO.I_COMPTEREF';
    // FQ 18828 stSQL := stSQL + ' WHERE (I_NATUREIMMO = "CB" OR I_NATUREIMMO = "PRO") AND (I_QUANTITE<>0)' +
    stWhere := stRecupWhere + ' AND (I_NATUREIMMO = "CB" OR I_NATUREIMMO = "PRO") AND (I_QUANTITE<>0)' +
                     ' AND (IL_TYPEOP="ACQ") AND (I_ETAT<>"FER")' +
                     ' AND (IL_DATEOP<="' + USDateTime(StrToDate(GetControlText('DATEARRETE')))+'")';
    // FQ 18828  stSQL := stSQL + ' ORDER BY I_IMMO';

  end
  // BTY 09/06 Pas de Gestion des DPI en SERIE1
  {else if (FNatureEtat='IDD') then  // TGA 05/07/2006 - Déclaration des DPI
  begin
    stSQL := 'SELECT '+stChampsIMMO+stChampsGENERAL;
    stSQL := stSQL + ',IZ_IMMO,IZ_LIBELLE,IZ_DATE,IZ_NATURED,IZ_MONTANT FROM IMMOMVTD ';
    stSQL := stSQL + ' LEFT JOIN IMMO ON (IMMOMVTD.IZ_IMMO=IMMO.I_IMMO)';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON GENERAUX.G_GENERAL=IMMO.I_COMPTEREF';
    stSQL := stSQL + ' WHERE IZ_NATURED="IMM" AND I_NATUREIMMO="PRO"';
    stSQL := stSQL + ' AND (I_DATEDEBECO >= "' + USDateTime(StrToDate(GetControlText('DATEDEBEXO')))+'"'+
                     ' AND  I_DATEDEBECO <= "' + USDateTime(StrToDate(GetControlText('DATEFINEXO')))+'" )';
    stSQL := stSQL + ' ORDER BY I_DATEDEBECO, IZ_IMMO,IZ_DATE';
  end}
  else if ((FNatureEtat='ITE') or (FNatureEtat='ITF')) then
  begin
    // MVG 17/08/2006 FQ 18636 - ajout BLOCNOTE et Dates de début d'amortissement
    stSQL := 'SELECT IIF(IL_DATEOP<"'+ USDateTime(StrToDate(GetControlText('DATEDEBEXO')))+'",IL_TYPEOP||"_1",IL_TYPEOP||"_2") TYPEOPERATION,'+
              stChampsIMMO+stChampsGENERAL+
              ',IL_DATEOP,IL_TYPEOP,IL_VOCEDEE,IL_DOTCESSECO,IL_DOTCESSFIS,IL_VOCEDEE,I_BLOCNOTE,I_DATEDEBECO,I_DATEDEBFIS';
    // Fin FQ18636
    stSQL := stSQL + ' FROM IMMOLOG';
    stSQL := stSQL + ' LEFT JOIN IMMO ON IMMO.I_IMMO=IMMOLOG.IL_IMMO';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON GENERAUX.G_GENERAL=IMMO.I_COMPTEREF';
    // FQ 18828 stSQL := stSQL + ' WHERE (I_NATUREIMMO="PRO" OR I_NATUREIMMO="FI") AND (I_ETAT<>"FER") AND '+
    stWhere := stRecupWhere + ' AND (I_NATUREIMMO="PRO" OR I_NATUREIMMO="FI") AND (I_ETAT<>"FER") AND '+
                  ' (((IL_TYPEOP="ACQ" AND IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATEARRETE')))+'")'+
                  ' OR (IL_DATEOP >= "'+ USDateTime(StrToDate(GetControlText('DATEDEBEXO')))+'"'+
                  ' AND IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"'+
                  ' AND IL_TYPEOP LIKE "CE%")))';
    // FQ 18828 stSQL := stSQL + ' ORDER BY I_IMMO';
  end
  else if ((FNatureEtat='ITR')) then // BTY 13/11/07 ) then oubliés //  FQ 21422 - or (FNatureEtat='IRE')) then
  begin
    // MVG 17/08/2006 FQ 18636 - ajout BLOCNOTE et Dates de début d'amortissement
    stSQL := 'SELECT IIF(IL_DATEOP<"'+ USDateTime(StrToDate(GetControlText('DATEDEBEXO')))+'",IL_TYPEOP||"_1",IL_TYPEOP||"_2") TYPEOPERATION,'+
              stChampsIMMO+stChampsGENERAL+
              ',IL_DATEOP,IL_TYPEOP,IL_VOCEDEE,IL_DOTCESSECO,IL_DOTCESSFIS,IL_VOCEDEE,I_BLOCNOTE,I_DATEDEBECO,I_DATEDEBFIS';
    // Fin FQ18636
    stSQL := stSQL + ' FROM IMMOLOG';
    stSQL := stSQL + ' LEFT JOIN IMMO ON IMMO.I_IMMO=IMMOLOG.IL_IMMO';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON GENERAUX.G_GENERAL=IMMO.I_COMPTEREF';
    // FQ 18828 stSQL := stSQL + ' WHERE (I_NATUREIMMO="PRO" OR I_NATUREIMMO="FI") AND (I_ETAT<>"FER") AND '+
    stWhere := stRecupWhere + ' AND (I_NATUREIMMO="PRO" OR I_NATUREIMMO="FI") AND (I_ETAT<>"FER") AND ((I_NONDED<>"X") OR (I_METHODEECO="VAR")) '+
                  ' (((IL_TYPEOP="ACQ" AND IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATEARRETE')))+'")'+
                  ' OR (IL_DATEOP >= "'+ USDateTime(StrToDate(GetControlText('DATEDEBEXO')))+'"'+
                  ' AND IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"'+
                  ' AND IL_TYPEOP LIKE "CE%")))';
    // FQ 18828 stSQL := stSQL + ' ORDER BY I_IMMO';
  end  // BTY 13/11/07 end oublié
  else if (FNatureEtat='IRE') then  // ajout pour FQ 21422
  begin
    stSQL := 'SELECT IIF(IL_DATEOP<"'+ USDateTime(StrToDate(GetControlText('DATEDEBEXO')))+'",IL_TYPEOP||"_1",IL_TYPEOP||"_2") TYPEOPERATION,'+
              stChampsIMMO+stChampsGENERAL+
              ',IL_DATEOP,IL_TYPEOP,IL_VOCEDEE,IL_DOTCESSECO,IL_DOTCESSFIS,IL_VOCEDEE,I_BLOCNOTE,I_DATEDEBECO,I_DATEDEBFIS';
    stSQL := stSQL + ' FROM IMMOLOG';
    stSQL := stSQL + ' LEFT JOIN IMMO ON IMMO.I_IMMO=IMMOLOG.IL_IMMO';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON GENERAUX.G_GENERAL=IMMO.I_COMPTEREF';
    // FQ 21422 stWhere := stRecupWhere + ' AND (I_NATUREIMMO="PRO" OR I_NATUREIMMO="FI") AND (I_ETAT<>"FER") AND ((I_NONDED<>"X") OR (I_METHODEECO="VAR")) '+
                  stWhere := stRecupWhere + ' AND (I_NATUREIMMO="PRO" OR I_NATUREIMMO="FI") AND (I_ETAT<>"FER") '+
                  ' (((IL_TYPEOP="ACQ" AND IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATEARRETE')))+'")'+
                  ' OR (IL_DATEOP >= "'+ USDateTime(StrToDate(GetControlText('DATEDEBEXO')))+'"'+
                  ' AND IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"'+
                  ' AND IL_TYPEOP LIKE "CE%")))';

  // TGA Tableau des amortissements variables
  end else if (FNatureEtat='ITV') then
  begin
    // Début MVG 07/07/06 FQ 18506
    stSQL := 'SELECT IIF(IL_DATEOP<"'+ USDateTime(StrToDate(GetControlText('DATEDEBEXO')))+'",IL_TYPEOP||"_1",IL_TYPEOP||"_2") TYPEOPERATION,'+
              stChampsIMMO+stChampsGENERAL+
              ',IL_DATEOP,IL_TYPEOP,I_UNITEOEUVRE,I_INFOUO, I_BLOCNOTE,I_DATEDEBECO';
    stSQL := stSQL + ' FROM IMMOLOG';
    stSQL := stSQL + ' LEFT JOIN IMMO ON IMMO.I_IMMO=IMMOLOG.IL_IMMO';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON GENERAUX.G_GENERAL=IMMO.I_COMPTEREF';
    // FQ 18828 stSQL := stSQL + ' WHERE (I_NATUREIMMO="PRO" OR I_NATUREIMMO="CB") AND (I_ETAT<>"FER") AND (I_METHODEECO="VAR") AND '+
    stWhere := stRecupWhere + ' AND (I_NATUREIMMO="PRO" OR I_NATUREIMMO="CB") AND (I_ETAT<>"FER") AND (I_METHODEECO="VAR") AND '+
                  ' (((IL_TYPEOP="ACQ" AND IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATEARRETE')))+'")'+
                  ' OR (IL_DATEOP >= "'+ USDateTime(StrToDate(GetControlText('DATEDEBEXO')))+'"'+
                  ' AND IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"'+
                  ' AND IL_TYPEOP LIKE "CE%")))';
    // FQ 18828 stSQL := stSQL + ' ORDER BY I_COMPTEREF,TYPEOPERATION,IL_DATEOP,I_IMMO';
    // FIN MVG 17/08/2006 FQ 18636

  end else if (FNatureEtat='ITN') then    // mbo fq 17512 modif sur le where
  begin
    stSQL := 'SELECT '+stChampsIMMO+stChampsGENERAL;
    stSQL := stSQL + ' FROM IMMO';
    stSQL := stSQL + ' LEFT JOIN IMMOLOG ON (I_IMMO=IL_IMMO) ';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON G_GENERAL=I_COMPTEREF';
    // FQ 18828 stSQL := stSQL + ' WHERE I_ETAT<>"FER" AND IL_DATEOP<="'+USDateTime(StrToDate(GetControlText('DATEARRETE')))+'" AND IL_TYPEOP="ACQ"';
    stWhere := stRecupWhere + ' AND I_ETAT<>"FER" AND IL_DATEOP<="'+USDateTime(StrToDate(GetControlText('DATEARRETE')))+'" AND IL_TYPEOP="ACQ"';
    //stSQL := stSQL + ' AND (I_REINTEGRATION<>0 OR I_QUOTEPART<>0)'
    // FQ 17512 stWhere := stWhere + ' AND (I_REINTEGRATION<>0 OR I_QUOTEPART<>0)'
    stWhere := stWhere + ' AND (I_REINTEGRATION<>0 AND I_QUOTEPART = 0)';

  end else if (FNatureEtat='IRD') then   // ajout mbo pour réintégration déduction fq 17512
  begin
    stSQL := 'SELECT '+stChampsIMMO+stChampsGENERAL;
    stSQL := stSQL + ' FROM IMMO';
    stSQL := stSQL + ' LEFT JOIN IMMOLOG ON (I_IMMO=IL_IMMO) ';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON G_GENERAL=I_COMPTEREF';
    // FQ 18828 stSQL := stSQL + ' WHERE I_ETAT<>"FER" AND IL_DATEOP<="'+USDateTime(StrToDate(GetControlText('DATEARRETE')))+'" AND IL_TYPEOP="ACQ"';
    stWhere := stRecupWhere + ' AND I_ETAT<>"FER" AND IL_DATEOP<="'+USDateTime(StrToDate(GetControlText('DATEARRETE')))+'" AND IL_TYPEOP="ACQ"';
    stWhere := stWhere + ' AND I_REINTEGRATION=0 AND I_QUOTEPART=0';
    stWhere := stWhere + ' AND (I_SBVPRI<>0 OR I_SBVPRIC<>0 OR I_NONDED<>"-")';

  end else if (FNatureEtat='ITQ') then    // mbo fq 17512 NOUVEL ETAT QUOTE PART
  begin
    stSQL := 'SELECT '+stChampsIMMO+stChampsGENERAL;
    stSQL := stSQL + ' FROM IMMO';
    stSQL := stSQL + ' LEFT JOIN IMMOLOG ON (I_IMMO=IL_IMMO) ';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON G_GENERAL=I_COMPTEREF';
    stWhere := stRecupWhere + ' AND I_ETAT<>"FER" AND IL_DATEOP<="'+USDateTime(StrToDate(GetControlText('DATEARRETE')))+'" AND IL_TYPEOP="ACQ"';
    stWhere := stWhere + ' AND (I_REINTEGRATION=0 AND I_QUOTEPART <> 0)';

  end else if (FNatureEtat='ITO') then
  begin
    stSQL := 'SELECT '+stChampsIMMO+stChampsGENERAL;
    stSQL := stSQL + ' FROM IMMO';
    stSQL := stSQL + ' LEFT JOIN IMMOLOG ON (I_IMMO=IL_IMMO) ';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON G_GENERAL=I_COMPTEREF';
    // FQ 18828 stSQL := stSQL + ' WHERE I_ETAT<>"FER" AND IL_DATEOP<="'+USDateTime(StrToDate(GetControlText('DATEARRETE')))+'" AND IL_TYPEOP="ACQ"';
    stWhere := stRecupWhere + ' AND I_ETAT<>"FER" AND IL_DATEOP<="'+USDateTime(StrToDate(GetControlText('DATEARRETE')))+'" AND IL_TYPEOP="ACQ"';
    // FQ 18828 + FQ 18860 stSQL := stSQL + ' AND (I_METHODEECO="LIN" OR I_METHODEECO="DEG")';
    stWhere := stWhere + ' AND (I_METHODEECO="LIN" OR I_METHODEECO="DEG" OR I_METHODEECO="VAR")';
  end else

//  begin
  if (FNatureEtat='ILS') then
  begin
    stSQL := 'SELECT '+stChampsIMMO+stChampsGENERAL;
    stSQL := stSQL + ' FROM IMMO ';
    stSQL := stSQL + ' LEFT JOIN IMMOLOG ON (IMMO.I_IMMO=IMMOLOG.IL_IMMO)';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON GENERAUX.G_GENERAL=IMMO.I_COMPTEREF';
    // FQ 18828 stSQL := stSQL + ' WHERE IL_TYPEOP="ACQ" AND I_ETAT<>"FER" AND IL_DATEOP<="'
    stWhere := stRecupWhere +  ' AND IL_TYPEOP="ACQ" AND I_ETAT<>"FER" AND IL_DATEOP<="'
               + USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"';
    // FQ 18828 stSQL := stSQL + ' ORDER BY I_IMMO';

  end else
  if (FNatureEtat='ISO') or (FNatureEtat='IPM') then
  begin
    stSQL := 'SELECT '+stChampsIMMO+stChampsGENERAL;
    stSQL := stSQL + ',IL_DATEOP,IL_MONTANTCES,IL_VOCEDEE,IL_PVALUE,IL_MOTIFCES';
    stSQL := stSQL + ' FROM IMMOLOG';
    stSQL := stSQL + ' LEFT JOIN IMMO ON I_IMMO=IL_IMMO';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON G_GENERAL=I_COMPTEREF';
    // FQ 18828 stSQL := stSQL + ' WHERE IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATESORTIE_')))+'"';
    // FQ 19171 : ajout du test sur date op <= date d'arrêté - mbo - 24.07.07
    stWhere := stRecupWhere + ' AND IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATESORTIE_')))+'"' +
               ' AND IL_DATEOP>="'+ USDateTime(StrToDate(GetControlText('DATESORTIE')))+'"' +
               ' AND IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"' +
               ' AND IL_TYPEOP="CES"';
    // FQ 18828 stSQL := stSQL + ' ORDER BY IL_MOTIFCES,I_IMMO';

  end else
  if (FNatureEtat='IAC') then
  begin
    stSQL := 'SELECT '+stChampsIMMO+stChampsGENERAL;
    stSQL := stSQL + ' FROM IMMOLOG';
    stSQL := stSQL + ' LEFT JOIN IMMO ON I_IMMO=IL_IMMO';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON G_GENERAL=I_COMPTEREF';
    // FQ 18828 stSQL := stSQL + ' WHERE IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('I_DATEPIECEA_')))+'"';
    stWhere := stRecupWhere + ' AND IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('I_DATEPIECEA_')))+'"' +
               ' AND IL_DATEOP>="'+ USDateTime(StrToDate(GetControlText('I_DATEPIECEA')))+'"' +
               ' AND IL_TYPEOP="ACQ"';
    // FQ 18828 stSQL := stSQL + ' ORDER BY I_IMMO';
  end else
  if (FNatureEtat='IMU') then
  begin
    if (GetCheckBoxState('MUTCOMPTE')<>cbChecked) then
      begin
       stSQL := 'SELECT '+stChampsIMMO+stChampsGENERAL;
       stSQL := stSQL + ',IL_TYPEOP,IL_CPTEMUTATION,IL_CODEMUTATION';
       stSQL := stSQL + ' FROM IMMO';
       stSQL := stSQL + ' LEFT JOIN IMMOLOG ON (I_IMMO=IL_IMMO)';
       stSQL := stSQL + ' LEFT JOIN GENERAUX ON G_GENERAL=I_COMPTEREF';
       // FQ 18828 stSQL := stSQL + ' WHERE (I_ETAT<>"FER" AND IL_DATEOP<="'+USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"';
       stWhere := stRecupWhere +  ' AND (I_ETAT<>"FER" AND IL_DATEOP<="'+USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"' +
                  ' AND I_OPEMUTATION="X"  AND IL_TYPEOP="MUT")';
       // FQ 18828 stSQL := stSQL + ' ORDER BY I_IMMO';
       stRecupOrderBy := stRecupOrderBy + ',I_IMMO';
      end
    else
      begin
       stSQL := 'SELECT '+stChampsIMMO+stChampsGENERAL;
       stSQL := stSQL + ',IL_TYPEOP,IL_CPTEMUTATION,IL_CODEMUTATION';
       stSQL := stSQL + ' FROM IMMO';
       stSQL := stSQL + ' LEFT JOIN IMMOLOG ON (I_IMMO=IL_IMMO)';
       stSQL := stSQL + ' LEFT JOIN GENERAUX ON G_GENERAL=I_COMPTEREF';
       // FQ 18828 stSQL := stSQL + ' WHERE (I_ETAT<>"FER" AND IL_DATEOP<="'+USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"';
       stWhere := stRecupWhere + ' AND (I_ETAT<>"FER" AND IL_DATEOP<="'+USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"' +
                  ' AND I_OPEMUTATION="X"  AND IL_TYPEOP="MUT" AND IL_CPTEMUTATION <> "")';
       // FQ 18828 stSQL := stSQL + ' ORDER BY I_IMMO';
    end
  end else
  if (FNatureEtat='ITP') then
  begin
    stSQL := 'SELECT IIF(I_NATUREIMMO="LOC","OUI","NON") LOCATION,';
    stSQL := stSQL + stChampsIMMO+stChampsGENERAL;
    stSQL := stSQL + ' FROM IMMO';
    stSQL := stSQL + ' LEFT JOIN IMMOLOG ON I_IMMO=IL_IMMO ';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON G_GENERAL=I_COMPTEREF';
    // FQ 18828 stSQL := stSQL + ' WHERE I_ETAT<>"FER" AND I_NATUREIMMO<>"FI" AND IL_DATEOP<="'+
    // BTY FQ 19036
    stWhere := stRecupWhere + ' AND I_ETAT<>"FER" AND I_NATUREIMMO<>"FI" AND IL_DATEOP<="'+
                            USDateTime(StrToDate(GetControlText('DATEARRETE')))
                            +'" AND I_BASETAXEPRO<>0 AND I_OPELEVEEOPTION<>"X"'
                            +' AND IL_TYPEOP="ACQ" AND I_QUANTITE<>0'
                            +' AND ( ((I_NATUREIMMO="CB" OR I_NATUREIMMO="LOC") AND (I_DATEFINCB>="'+
                            USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"))'
                            +' OR (I_NATUREIMMO<>"CB" AND I_NATUREIMMO<>"LOC"))';
    // FQ 18828 stSQL := stSQL + ' ORDER BY I_IMMO';
  end else
  if (FNatureEtat='IPR') then
  begin
    stSQL := 'SELECT ' + stChampsIMMO+stChampsGENERAL;
    stSQL := stSQL + ' FROM IMMO';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON G_GENERAL=I_COMPTEREF';
    // FQ 18828 stSQL := stSQL + ' WHERE I_ETAT<>"FER" AND I_DATEPIECEA<="'
    stWhere := stRecupWhere + ' AND I_ETAT<>"FER" AND I_DATEPIECEA<="'
                            + USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"';
    // FQ 18828 stSQL := stSQL + ' ORDER BY I_IMMO';
  end else

  // BTY 10/06 Suramortissement des primes d'équipement
  if (FNatureEtat='ITS') then
  begin
    stSQL := 'SELECT '+stChampsIMMO+stChampsGENERAL;
    stSQL := stSQL + ' FROM IMMO ';
    stSQL := stSQL + ' LEFT JOIN IMMOLOG ON (IMMO.I_IMMO=IMMOLOG.IL_IMMO)';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON GENERAUX.G_GENERAL=IMMO.I_COMPTEREF';
    stWhere := stRecupWhere +  ' AND IL_TYPEOP="ACQ" AND I_ETAT<>"FER" AND IL_DATEOP<="'
               + USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"'
               + ' AND (I_SBVPRI<> 0 OR I_SBVPRIC <> 0)';
  end else
  // BTY 10/06 Amortissement des subventions
  if (FNatureEtat='ITB') then
  begin
    stSQL := 'SELECT '+stChampsIMMO+stChampsGENERAL;
    stSQL := stSQL + ' FROM IMMO ';
    stSQL := stSQL + ' LEFT JOIN IMMOLOG ON (IMMO.I_IMMO=IMMOLOG.IL_IMMO)';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON GENERAUX.G_GENERAL=IMMO.I_COMPTEREF';
    stWhere := stRecupWhere +  ' AND IL_TYPEOP="ACQ" AND I_ETAT<>"FER" AND IL_DATEOP<="'
               + USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"'
               + ' AND (I_SBVMT<> 0 OR I_SBVMTC <> 0)';
  end else
  // BTY 13/11/07 Immos issues du passage forfait au réel pour les biens agricoles
  if (FNatureEtat='IFR') then
  begin
    stSQL := 'SELECT '+stChampsIMMO;
    stSQL := stSQL + ',G1.G_LIBELLE LIB1,G1.G_GENERAL,G2.G_LIBELLE LIB2,G2.G_GENERAL';
    stSQL := stSQL + ' FROM IMMO';
    stSQL := stSQL + ' LEFT JOIN IMMOLOG ON I_IMMO=IL_IMMO';
    stSQL := stSQL + ' LEFT JOIN GENERAUX G1 ON G1.G_GENERAL=I_COMPTEREF';
    stSQL := stSQL + ' LEFT JOIN GENERAUX G2 ON G2.G_GENERAL=I_COMPTEIMMO';
    stWhere := stRecupWhere + ' AND (I_ETAT<>"FER") AND (I_NATUREIMMO ="PRO") AND (I_PFR="X")'
                   + ' AND IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"';
  end else
  if (FNatureEtat='IVE') then
  begin
  // BTY 09/06 Nature IVE présente 2 codes état IVE et IVV (inventaire vierge)
  if (FCodeEtat='IVE') then
  begin
    // MVG 18/08/2006 FQ 18637
    // La double jointure sur la table IMMOLOG ne fonctionne pas en ACCESS.
    // On récupére aucun enregistrement (opération=ACQ et opération=CES)
    // Modification du fonctionnement pour ne plus imprimer les immos déjà sorties
    // dans l'exercice
    stSQL := 'SELECT '+ stChampsIMMO+stChampsGENERAL;
    stSQL := stSQL + ' ,I_BLOCNOTE';
    stSQL := stSQL + ' FROM IMMO';
    stSQL := stSQL + ' LEFT JOIN IMMOLOG IL1 ON (I_IMMO=IL1.IL_IMMO)';
    stSQL := stSQL + ' LEFT JOIN GENERAUX ON G_GENERAL=I_COMPTEREF';
    // FQ 18828 stSQL := stSQL + ' WHERE I_ETAT<>"FER" AND IL1.IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"';
    stWhere := stRecupWhere + ' AND I_ETAT<>"FER" AND IL1.IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"' +
               ' AND IL1.IL_TYPEOP="ACQ" AND I_QUANTITE<>0';
    // FQ 18828 stSQL := stSQL + ' ORDER BY I_IMMO';

    // ancien code
    //stSQL := 'SELECT '+ stChampsIMMO+stChampsGENERAL;
    //stSQL := stSQL + ' ,IL2.IL_DATEOP DATESORTIE, IL2.IL_MOTIFCES,IL2.IL_MONTANTCES';
    //stSQL := stSQL + ' FROM IMMO';
    //stSQL := stSQL + ' LEFT JOIN IMMOLOG IL1 ON (I_IMMO=IL1.IL_IMMO)';
    //stSQL := stSQL + ' LEFT JOIN IMMOLOG IL2 ON (I_IMMO=IL2.IL_IMMO)';
    //stSQL := stSQL + ' LEFT JOIN GENERAUX ON G_GENERAL=I_COMPTEREF';
    //stSQL := stSQL + ' WHERE I_ETAT<>"FER" AND IL1.IL_DATEOP<="'+ USDateTime(StrToDate(GetControlText('DATEARRETE')))+'"';
    //stSQL := stSQL + ' AND IL1.IL_TYPEOP="ACQ"  AND IL2.IL_TYPEOP="CES"';
    //stSQL := stSQL + ' ORDER BY I_IMMO';
    // Fin MVG 18/08/2006 FQ 18637

  end;
  end;
  // FQ 18828 TFQRS1(Ecran).WhereSQL := trim(stSQL);
  if StSQL<>'' then //YCP 17/10/2006 sinon on ecrase par exemple les ITE
     TFQRS1(Ecran).WhereSQL := trim(stSQL) + stWhere + stRecupOrderBy;
{$ENDIF}
end ;

procedure TOF_AMEDITION.OnLoad ;
begin
  Inherited ;

end ;

procedure TOF_AMEDITION.OnArgument (S : String ) ;
var stTous : string;
    I_LieuGeo : THMultiValComboBox;
    Sender: TObject;
begin
  Inherited ;
  // TGA 13/07/06 pour warning
  Sender := TFQRS1(ECRAN);

  VHImmo^.PlanInfo.Free ; VHImmo^.PlanInfo:=nil ;
  TFQRS1(ECRAN).NatureEtat := S;
  FCodeEtat := TFQRS1(ECRAN).CodeEtat;
  FNatureEtat := TFQRS1(ECRAN).NatureEtat;
  ProcZoomEdt := ImZoomEdtEtatImmo ;
  {$IFDEF EAGLCLIENT}
  ProcCalcEdt := nil;
  CalcAmortServer('','') ;
  {$ELSE}
  ProcCalcEdt := CalcOLEEtatImmo ;
  {$ENDIF}
  TFQRS1(ECRAN).NatureEtat := S;
  TFQRS1(ECRAN).FFiltres.OnChange := OnChangeFiltre;
  THEdit(GetControl('I_COMPTEREF')).OnExit := OnExitCompteRef;
  THEdit(GetControl('I_COMPTEREF_')).OnExit := OnExitCompteRef;
  THEdit(GetControl('I_DATEPIECEA_')).OnExit := OnExitDateAchatFin; // FQ 19036
  THEdit(GetControl('DATEARRETE')).OnExit := OnExitDateArrete;   // FQ 19036
  SetControlText ( 'DATEDEBEXO',DateToStr(VHImmo^.EnCours.Deb) );
  SetControlText ( 'DATEFINEXO',DateToStr(VHImmo^.EnCours.Fin) );
  SetControlText ( 'DATEARRETE',DateToStr(VHImmo^.EnCours.Fin) );
  // Par défaut , pas de rupture sur l'état
  ChargeListeRupture;
  RenseigneComboRupture ( nil, '');
  THValComboBox(GetControl('RUPTURE1')).ItemIndex := 0;
  THValComboBox(GetControl('RUPTURE2')).ItemIndex := 0;
  THValComboBox(GetControl('RUPTURE1')).OnChange := OnChangeRupture;
  THValComboBox(GetControl('RUPTURE2')).OnChange := OnChangeRupture;
  TCheckBox(GetControl('AVECREGROUPEMENT')).OnClick := OnClickRegroupement;
  // ajout mbo fq 14430
  TCheckBox(GetControl('MUTCOMPTE')).OnClick := OnClickMutcompte;

  // Valeurs par défaut des critères
  SetControlText('I_QUALIFIMMO','R');
  SetControlText('I_NATUREIMMO','');
  stTous := '<<' + TraduireMemoire('Tous') + '>>';
  I_LieuGeo := THMultiValComboBox(GetControl('I_LIEUGEO'));
  I_LieuGeo.SelectAll;
  if I_LieuGeo <> nil then if I_LieuGeo.Tous then I_LieuGeo.Text := stTous;
  // BTY 09/06 Restrictions utilisateur -> voir plus bas
  //I_Etablissement := THMultiValComboBox(GetControl('I_ETABLISSEMENT'));
  //I_Etablissement.SelectAll;
  //if I_Etablissement <> nil then if I_Etablissement.Tous then I_Etablissement.Text := stTous;

  // fq 14430 mbo 07.09.2005
  SetControlVisible('MUTCOMPTE', False);

  {$IFDEF SERIE1}
  SetControlVisible('I_ETABLISSEMENT',False);
  SetControlVisible('TI_ETABLISSEMENT',False);
  TFQRS1(Ecran).ParamEtat := False;
  TFQRS1(Ecran).CritAvancesVisibled:=false ;
  TFQRS1(Ecran).Avances.TabVisible:=false ;
  {$ELSE}
  //TFQRS1(Ecran).ParamEtat := ((not (ctxPCL in V_PGI.PGIContexte)) or (V_PGI.PassWord=CryptageSt(DayPass(Date))) or (FNatureEtat='IVL'));  //  FQ 19798 - MVG 24/05/2007, En Entreprise, tjs possible. En cabinet, si l'on est avec le mot de passe du jour
  // Suite 19/06/2007

TFQRS1(Ecran).ParamEtat := ((ctxPCL in V_PGI.PGIContexte) and (V_PGI.PassWord=CryptageSt(DayPass(Date)))) or (FNatureEtat='IVL') or (not(ctxPCL in V_PGI.PGIContexte));
  {$ENDIF}
  SetControlVisible('BEURO',False);
  SetControlVisible('BPROG',False);

  SetControlText('I_COMPTEREF', VHImmo^.CpteImmoInf);
  SetControlText('I_COMPTEREF', VHImmo^.CpteLocSup);

  { Initialisation contextuelle des critères }
  SetControlText('I_DATEPIECEA', DateToStr(iDate1900) ) ;
  SetControlText('I_DATEPIECEA_',DateToStr(VHImmo^.Encours.Fin) ) ;
  if (FCodeEtat = 'IAC') then
  begin
    SetControlText('I_DATEPIECEA',DateToStr(VHImmo^.Encours.Deb) ) ;
    SetControlText('I_DATEPIECEA_',DateToStr(VHImmo^.Encours.Fin) ) ;
  end;
  if (FCodeEtat = 'ISO') or (FNatureEtat='IPM') then
  begin
    SetControlText('DATESORTIE',DateToStr(VHImmo^.Encours.Deb) ) ;
    SetControlText('DATESORTIE_',DateToStr(VHImmo^.Encours.Fin) ) ;
  end;

  THCritMaskEdit(GetControl('I_COMPTEREF')).OnElipsisClick := OnClickCompteRef;
  THCritMaskEdit(GetControl('I_COMPTEREF_')).OnElipsisClick := OnClickCompteRef;
  THValComboBox(GetControl('I_NATUREIMMO')).OnChange := OnChangeNature;
  THValComboBox(GetControl('FETAT')).OnChange := OnChangeEtat;
  {$IFDEF SERIE1}
  THValComboBox(GetControl('TABLELIBRE1')).OnChange := OnChangeChampLibre;
  THValComboBox(GetControl('TABLELIBRE2')).OnChange := OnChangeChampLibre;
  THValComboBox(GetControl('TABLELIBRE3')).OnChange := OnChangeChampLibre;
  {$ENDIF}

  {$IFDEF SERIE1}
  ImLibellesTableLibre(TTabSheet(GetControl('PCHAMPSLIBRES')),'TT_TABLELIBREIMMO','','I') ;
  SetControlVisible('PTABLESLIBRES',False);
  {$ELSE}
  ImLibellesTableLibre(TTabSheet(GetControl('PTABLESLIBRES')),'TI_TABLE','I_TABLE','I') ;
  { Table libre 9 utilisée pour le stockage 'éligible TP' }
  SetControlVisible ('I_TABLE9',False);
  SetControlVisible ('TI_TABLE9',False);
  SetControlVisible('PCHAMPSLIBRES',False);
  {$ENDIF}

  InitZonesSuivantEtat;

  InitRuptures;

  OnChangeNature ( nil );

  MajLookUpSection ( '' );

  { Mise à jour des cases à cocher saut de page }
  SetControlEnabled ( 'SAUTPAGERUPTURE1',GetControlText('RUPTURE1')<>'');
  SetControlEnabled ( 'SAUTPAGERUPTURE2',GetControlText('RUPTURE2')<>'');
  if GetControlText('RUPTURE1')='' then SetControlChecked ( 'SAUTPAGERUPTURE1', False);
  if GetControlText('RUPTURE2')='' then SetControlChecked ( 'SAUTPAGERUPTURE2', False);


  // BTY 09/06 Restrictions utilisateur -> voir plus bas
  //{$IFNDEF SERIE1}
  //PositionneEtabUser(THMultiValComboBox(GetControl('I_ETABLISSEMENT')));
  //{$ENDIF}

  // TGA 07/07/2006
  if (FCodeEtat = 'IDD') then
    Begin
      //  Sauvegarde de l'action valider
      // FQ 21899 Fvalidation := TToolbarButton97(GetControl('BValider')).OnClick;
      // Dates d'achat = dates exo non modifiables
      SetControlText ( 'I_DATEPIECEA',DateToStr(VHImmo^.EnCours.Deb) );
      THEdit(GetControl('I_DATEPIECEA_')).Enabled := False;
      SetControlText ( 'I_DATEPIECEA_',DateToStr(VHImmo^.EnCours.Fin) );
      THEdit(GetControl('I_DATEPIECEA')).Enabled := False;
      // Dates d'arrêté = date fin exo non modifiable
      SetControlText ( 'DATEARRETE',DateToStr(VHImmo^.EnCours.Fin) );
      THEdit(GetControl('DATEARRETE')).Enabled := False;
      // Nature immo = 'PRO
      THValComboBox(GetControl('I_NATUREIMMO')).Enabled := False;
      SetControlText('I_NATUREIMMO','PRO'); // FQ 21899
      // Fourchette de compte pour nature PRO
      SetControlText('I_COMPTEREF_',VHImmo^.CpteImmoSup);
      THEdit(GetControl('I_COMPTEREF')).Enabled := False;
      SetControlText('I_COMPTEREF',VHImmo^.CpteImmoInf);
      THEdit(GetControl('I_COMPTEREF_')).Enabled := False;
      // FQ 21899 FValidation(Sender);
    End;
    // 09/2007
    if (FNatureEtat = 'IVL') then
    begin
      SetActiveTabSheet('Option');
      TCheckBox(GetControl('SAUTPAGECOMPTE')).Enabled := False;
      THCheckBox(GetControl('FREDUIRE')).Enabled := False;
      THCheckBox(GetControl('FLISTE')).Enabled := False;
      TCheckBox(GetControl('ETATSYNTH')).Enabled := False;
      TCheckBox(GetControl('RAPPELCRITERES')).Enabled := False;
      TCheckBox(GetControl('IMPRIMERDATE')).Enabled := False;
      SetControlChecked ( 'IMPRIMERDATE', False);
    end;
    // BTY 13/11/07 Edition Déclaration agricole
    if (FNatureEtat = 'IFR') then
    Begin
      // Codes immos bloqués
      THEdit(GetControl('I_IMMO')).Enabled := False;
      THEdit(GetControl('I_IMMO_')).Enabled := False;
      // Dates d'arrêté bloquée
      SetControlText ( 'DATEARRETE',DateToStr(VHImmo^.EnCours.Fin) );
      THEdit(GetControl('DATEARRETE')).Enabled := False;
      // Dates d'achat cachées
      SetControlText ( 'I_DATEPIECEA',DateToStr(VHImmo^.EnCours.Deb) );
      SetControlText ( 'I_DATEPIECEA_',DateToStr(VHImmo^.EnCours.Fin) );
      THEdit(GetControl('I_DATEPIECEA')).Enabled := False;
      THEdit(GetControl('I_DATEPIECEA_')).Enabled := False;
      // Nature 'PRO'
      SetControlText('I_NATUREIMMO','PRO');
      THValComboBox(GetControl('I_NATUREIMMO')).Enabled := False;
      // Fourchette de compte pour nature PRO
      SetControlText('I_COMPTEREF_',VHImmo^.CpteImmoSup);
      THEdit(GetControl('I_COMPTEREF')).Enabled := False;
      SetControlText('I_COMPTEREF',VHImmo^.CpteImmoInf);
      THEdit(GetControl('I_COMPTEREF_')).Enabled := False;
    end;

    // BTY 09/06 Utiliser la fonction générale de ImEnt
    // TGA 13/09/2006  gestion des restrictions utilisateurs sur les établissements
    {$IFNDEF SERIE1}
    //FComboEtab := nil;
    //GetComboEtab;
    //GereEtablissement;
    //ControlEtab;
    ImGereEtablissement (THMultiValComboBox(GetControl('I_ETABLISSEMENT')) );
    {$ELSE}
    // Bloquer la date d'arrêté si taxe professionnelle
    if (FNatureEtat='ITP') then
       THEdit(GetControl('DATEARRETE')).Enabled := False;
    {$ENDIF}
end ;


procedure TOF_AMEDITION.OnClose ;
{$IFDEF EAGLCLIENT}
var T, TResponse : TOB;
{$ENDIF}
begin
  {$IFDEF EAGLCLIENT}
  T := TOB.Create ('',nil,-1);
  try
    TResponse := AppServer.Request('Amortissement.AmEdClose','',T,'','');
    if TResponse = nil then PGIBox('Amortissement ( Fermeture ) : serveur d''éditions introuvable');
    TResponse.Free;
    delay(100); // on laisse le temps au serveur de se fermer
  finally
    T.Free;
  end;
  {$ENDIF}
  ProcZoomEdt:=ImZoomEdtEtat ;
  ProcCalcEdt:=ImCalcOLEEtat ;
  FTRupture.Free;
  VHImmo^.PlanInfo.Free ; VHImmo^.PlanInfo:=nil ;
  Inherited ;
end ;

procedure TOF_AMEDITION.OnDisplay () ;
begin
  Inherited ;
  { Table libre 9 utilisée pour le stockage 'éligible TP' }
  SetControlVisible ('I_TABLE9',False);
  SetControlVisible ('TI_TABLE9',False);
end ;

procedure TOF_AMEDITION.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_AMEDITION.OnChangeRupture(Sender: TObject);
var St : string;
begin
  if THValComboBox(Sender).Name='RUPTURE1' then RenseigneComboRupture ( THValComboBox(GetControl('RUPTURE2')), GetControlText('RUPTURE1'))
  else RenseigneComboRupture ( THValComboBox(GetControl('RUPTURE1')), GetControlText('RUPTURE2'));
  if (GetControlText('RUPTURE1')<>'') then
  begin
    if (Copy(GetControlText('RUPTURE1'),1,9)='S_SECTION') then
      SetControlText('XX_RUPTURE1',Copy(GetControlText('RUPTURE1'),1,9))
    else SetControlText('XX_RUPTURE1',GetControlText('RUPTURE1'));
    if GetControlText('XX_RUPTURE1')='S_SECTION' then
    begin
      St := THValComboBox(GetControl('RUPTURE1')).Text;
      SetControlText('S_AXE','A'+St[Length(St)]);
    end else SetControlText('S_AXE','');
  end
  else
  begin
    SetControlText('XX_RUPTURE1','');
    SetControlText('S_AXE','');
  end;
  if (GetControlText('RUPTURE2')<>'') then
  begin
    if (Copy(GetControlText('RUPTURE2'),1,9)='S_SECTION') then
      SetControlText('XX_RUPTURE2',Copy(GetControlText('RUPTURE2'),1,9))
    else SetControlText('XX_RUPTURE2',GetControlText('RUPTURE2'));
    if GetControlText('XX_RUPTURE2')='S_SECTION' then
    begin
      St := THValComboBox(GetControl('RUPTURE2')).Text;
      SetControlText('S_AXE','A'+St[Length(St)]);
    end else if GetControlText('XX_RUPTURE1')<>'S_SECTION' then SetControlText('S_AXE','');
  end
  else
  begin
    SetControlText('XX_RUPTURE2','');
    if GetControlText('XX_RUPTURE1')<>'S_SECTION' then SetControlText('S_AXE','');
  end;
  MajLookUpSection(GetControlText('S_AXE'));
  { Mise à jour des cases à cocher saut de page }
  SetControlEnabled ( 'SAUTPAGERUPTURE1',GetControlText('RUPTURE1')<>'');
  SetControlEnabled ( 'SAUTPAGERUPTURE2',GetControlText('RUPTURE2')<>'');
  if GetControlText('RUPTURE1')='' then SetControlChecked ( 'SAUTPAGERUPTURE1', False);
  if GetControlText('RUPTURE2')='' then SetControlChecked ( 'SAUTPAGERUPTURE2', False);
end;

procedure TOF_AMEDITION.AssignHelpContext;
begin
{$IFDEF SERIE1}
    if (fCodeEtat='FIM') then Ecran.HelpContext:=550500
  else if (fCodeEtat='DOT') then Ecran.HelpContext:=552100
  else if (fCodeEtat='DEC') then Ecran.HelpContext:=552200
  else if (fCodeEtat='DFI') then Ecran.HelpContext:=552300
  else if (fCodeEtat='DER') then Ecran.HelpContext:=552400
  else if (fCodeEtat='INV') then Ecran.HelpContext:=551200
  else if (fCodeEtat='ACQ') then Ecran.HelpContext:=551100
  else if (fCodeEtat='LSI') then Ecran.HelpContext:=551300
  else if (fCodeEtat='STS') then Ecran.HelpContext:=551400
  else if (fCodeEtat='PMV') then Ecran.HelpContext:=551600
  else if (fCodeEtat='MUT') then Ecran.HelpContext:=551500
  else if (fCodeEtat='DCB') then Ecran.HelpContext:=553500
  else if (fCodeEtat='ECB') then Ecran.HelpContext:=553600
  else if (fCodeEtat='PTP') then Ecran.HelpContext:=553900
  else if (fCodeEtat='PRE') then Ecran.HelpContext:=553100
  else if (fCodeEtat='PRF') then Ecran.HelpContext:=553200
  else if (fCodeEtat='ETI') then Ecran.HelpContext:=554000 ;
{$ELSE}
  if (fCodeEtat='FIM') then Ecran.HelpContext := 2505000
  else if (fCodeEtat='DOT') then Ecran.HelpContext := 2521000
  else if (fCodeEtat='DEC') then Ecran.HelpContext := 2522000
  else if (fCodeEtat='DFI') then Ecran.HelpContext := 2523000
  else if (fCodeEtat='DER') then Ecran.HelpContext := 2524000
  else if (fCodeEtat='INV') then Ecran.HelpContext := 2512000
  else if (fCodeEtat='ACQ') then Ecran.HelpContext := 2511000
  else if (fCodeEtat='LSI') then Ecran.HelpContext := 2513000
  else if (fCodeEtat='STS') then Ecran.HelpContext := 2514000
  else if (fCodeEtat='PMV') then Ecran.HelpContext := 2516000
  else if (fCodeEtat='MUT') then Ecran.HelpContext := 2515000
  else if (fCodeEtat='DCB') then Ecran.HelpContext := 2535000
  else if (fCodeEtat='ECB') then Ecran.HelpContext := 2536000
  else if (fCodeEtat='PTP') then Ecran.HelpContext := 2539000
  else if (fCodeEtat='PRE') then Ecran.HelpContext := 2531000
  else if (fCodeEtat='PRF') then Ecran.HelpContext := 2532000
  else if (fCodeEtat='ETI') then Ecran.HelpContext := 2540000;
{$ENDIF}
end;

procedure TOF_AMEDITION.OnClickCompteRef(Sender: TObject);
var CpteInf,CpteSup : string;
    sWhere : string;
    Nature : string;
begin
  Nature := GetControlText('I_NATUREIMMO');
  if Nature='CB' then begin CpteSup := VHImmo^.CpteCBSup; CpteInf := VHImmo^.CpteCBInf;end
  else if Nature='LOC' then begin CpteSup := VHImmo^.CpteLocSup; CpteInf := VHImmo^.CpteLocInf; end
  else if Nature='PRO' then begin CpteSup := VHImmo^.CpteImmoSup;CpteInf := VHImmo^.CpteImmoInf;end
  else if Nature='FI'  then begin CpteSup := VHImmo^.CpteFinSup; CpteInf := VHImmo^.CpteFinInf; end
  else begin CpteSup := VHImmo^.CpteLocSup; CpteInf :=VHImmo^.CpteImmoInf;end;
  sWhere:='G_GENERAL<="'+CpteSup+'" AND G_GENERAL>="'+CpteInf+'"';
  LookUpList (TControl (Sender),'Comptes généraux','GENERAUX','G_GENERAL','G_LIBELLE',sWhere,'G_GENERAL',True,0) ;
end;

procedure TOF_AMEDITION.OnChangeNature(Sender: TObject);
var Nature : string;
begin
  Nature := GetControlText ('I_NATUREIMMO');
  if Nature='CB' then
  begin
    SetControlText('I_COMPTEREF_',VHImmo^.CpteCBSup);
    SetControlText('I_COMPTEREF',VHImmo^.CpteCBInf);
  end
  else if Nature='LOC' then
  begin
    SetControlText('I_COMPTEREF_',VHImmo^.CpteLocSup);
    SetControlText('I_COMPTEREF',VHImmo^.CpteLocInf);
  end
  else if Nature='PRO' then
  begin
    SetControlText('I_COMPTEREF_',VHImmo^.CpteImmoSup);
    SetControlText('I_COMPTEREF',VHImmo^.CpteImmoInf);
  end
  else if Nature='FI'  then
  begin
    SetControlText('I_COMPTEREF_',VHImmo^.CpteFinSup);
    SetControlText('I_COMPTEREF',VHImmo^.CpteFinInf);
  end else
  begin
    SetControlText('I_COMPTEREF_',VHImmo^.CpteLocSup);
    SetControlText('I_COMPTEREF',VHImmo^.CpteImmoInf);
  end;
end;

procedure TOF_AMEDITION.OnChangeEtat(Sender: TObject);
begin
  InitZonesSuivantEtat;
  InitRuptures;

  // FQ 19927 La présence du filtre par défaut envoie OnChangeEtat => affecter ici
  TFQRS1(ECRAN).CodeEtat := THValComboBox(GetControl('FETAT')).value;
  fCodeEtat := TFQRS1(ECRAN).CodeEtat;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 24/06/2003
Modifié le ... : 01/07/2003
Description .. : Description des natures
Suite ........ : - AMS : Liste simplifiée
Suite ........ : - AMA : Liste des acquisitions
Suite ........ : - AME : Tableau d'amortissement économique
Suite ........ : - AMT : Liste des sorties
Suite ........ : - AMP : prévisionnel économique et fiscal ( AME :
Suite ........ : économique, AMF : fiscal ).
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOF_AMEDITION.InitZonesSuivantEtat;
begin
  TFQRS1(ECRAN).FNomFiltre := 'QRF'+FNatureEtat;
  if (FNatureEtat='ILS') then TFQRS1(ECRAN).Caption := 'Liste simplifiée des immobilisations'
  else if (FNatureEtat='ITE') then TFQRS1(ECRAN).Caption := 'Tableau d''amortissement économique'
  else if (FNatureEtat='ITF') then TFQRS1(ECRAN).Caption := 'Tableau d''amortissement fiscal'
  else if (FNatureEtat='ITN') then TFQRS1(ECRAN).Caption := 'Plafond de déductibilité fiscale'  //modif fq 17512 mbo
  else if (FNatureEtat='ITQ') then TFQRS1(ECRAN).Caption := 'Quote-part personnelle'   // ajout fq 17512
  else if (FNatureEtat='IRD') then TFQRS1(ECRAN).Caption := 'Réintégrations/déductions fiscales'  // ajout fq 17512
  else if (FNatureEtat='ITR') then TFQRS1(ECRAN).Caption := 'Suivi des amortissements dérogatoires'
  else if (FNatureEtat='ITV') then TFQRS1(ECRAN).Caption := 'Suivi des amortissements variables' // ajout TGA 16/02/2006
  else if (FNatureEtat='IDD') then TFQRS1(ECRAN).Caption := 'Déclaration des DPI'               // ajout  TGA 05/07/2006
  else if (FNatureEtat='IRE') then TFQRS1(ECRAN).Caption := 'Récapitulatif des amortissements'
  else if (FNatureEtat='ITO') then TFQRS1(ECRAN).Caption := 'Répartition des dotations'
  else if (FNatureEtat='IAC') then TFQRS1(ECRAN).Caption := 'Liste des acquisitions'
  else if (FNatureEtat='ISO') then TFQRS1(ECRAN).Caption := 'Liste des sorties'
  else if (FNatureEtat='IMU') then TFQRS1(ECRAN).Caption := 'Liste des mutations'
  else if (FNatureEtat='IPM') then TFQRS1(ECRAN).Caption := 'Plus ou moins values'
  else if (FNatureEtat='ILD') then TFQRS1(ECRAN).Caption := 'Liste des dépréciations/reprises'  // ajout mbo 17.11.05
  else if (FNatureEtat='ILC') then TFQRS1(ECRAN).Caption := 'Liste des changements de conditions'  // ajout TGA 07.12.05
  else if (FNatureEtat='ILV') then TFQRS1(ECRAN).Caption := 'Liste des valorisations'  // ajout PGR - FQ 15637 - 27.12.05
  else if (FNatureEtat='IVE') then TFQRS1(ECRAN).Caption := 'Inventaire'
  else if (FNatureEtat='IVL') then TFQRS1(ECRAN).Caption := 'Lettre d''inventaire'
  else if (FNatureEtat='IET') then TFQRS1(ECRAN).Caption := 'Etiquettes'  // ajout TGA 06.02.06
  else if (FNatureEtat='ITP') then TFQRS1(ECRAN).Caption := 'Etat préparatoire à la taxe professionnelle'
  else if (FNatureEtat='IEG') then TFQRS1(ECRAN).Caption := 'Etat des engagements'  // ajout mbo - fq 13036 - 13.01.06
  else if (FNatureEtat='IDS') then TFQRS1(ECRAN).Caption := 'Etat des dotations simulées'  // BTY 02/06 FQ 13035
  else if (FNatureEtat='ITS') then TFQRS1(ECRAN).Caption := 'Suramortissement des primes d''équipement'  // BTY 10/06 Primes équipement
  else if (FNatureEtat='ITB') then TFQRS1(ECRAN).Caption := 'Amortissement des subventions'  // BTY 10/06 Subventions
  else if (FNatureEtat='IRC') then TFQRS1(ECRAN).Caption := 'Remplacement de composants' // ajout mbo 26.06.07
  else if (FNatureEtat='IFR') then TFQRS1(ECRAN).Caption := 'Immobilisations issues du passage du forfait au réel' // BTY 13/11/07
  else if FNatureEtat = 'IPR' then
  begin
    if FCodeEtat = 'PRE' then TFQRS1(ECRAN).Caption := 'Prévision économique sur 5 ans'
    else if FCodeEtat = 'PRF' then TFQRS1(ECRAN).Caption := 'Prévision fiscale sur 5 ans';
  end;
  UpdateCaption(ECRAN) ;

  // ajout mbo 24.08.05 FQ 14659
  SetControlEnabled('ETATSYNTH',(FNatureEtat<>'IVE') and (FNatureEtat<>'IVL')); // BTY 09/2007 lettre inventaire

  // ajout mbo 08.09.05 FQ 14430 Check mutation compte uniquement
  SetControlVisible('MUTCOMPTE',(FNatureEtat = 'IMU'));

  SetControlVisible ('TI_DATECESSION', (FCodeEtat = 'ISO') or (FNatureEtat='IPM'));
  SetControlVisible ('DATESORTIE', (FCodeEtat = 'ISO') or (FNatureEtat='IPM'));
  SetControlVisible ('TI_DATECESSION_', (FCodeEtat = 'ISO') or (FNatureEtat='IPM'));
  SetControlVisible ('DATESORTIE_', (FCodeEtat = 'ISO') or (FNatureEtat='IPM'));
  SetControlVisible ('INFOCOMP',(FNatureEtat='ITE') or (FNatureEtat='ITV') or (FNatureEtat='IDD') or (FNatureEtat='ITF') or ( FNatureEtat='AMF') or (FNatureEtat='IVE'));

  // BTY 12/05 FQ 15243 Modif Caption cause check-box ajoutée
  //if (FNatureEtat = 'IVE') then SetControlCaption('INFOCOMP','Impression des totalisations');
  if (FNatureEtat = 'IVE') then SetControlCaption('INFOCOMP','Avec les totalisations');
  // BTY 12/05 FQ 15243 Impression du commentaire sur l'inventaire
  // Commentaire imprimé avec les infos complémentaires dans les tableaux ITE et ITF
  SetControlVisible ('INFOBN', (FCodeEtat = 'IVE'));

  SetControlVisible ('AVECMOTIF',(FNatureEtat='IPM') or (FNatureEtat='ISO'));

  // fq 13036 mbo 19.06.06 date de fin = date fin exercice pour état des engagements et dot simulées
  //SetControlEnabled('DATEARRETE',(FNatureEtat<>'IEG') and (FNatureEtat<> 'IDS'));
  // TGA 12/07 ajout IDD

  // fq 19913 - mbo - SetControlEnabled('DATEARRETE',(FNatureEtat<>'IEG') and (FNatureEtat<> 'IDS') and (FNatureEtat<> 'IDD'));
  // BTY 13/11/07 ajout IFR
  SetControlEnabled('DATEARRETE',(FNatureEtat<>'IEG') and (FNatureEtat<> 'IDD') and (FNatureEtat<>'IFR'));
  SetControlVisible ('DATEARRETE', (FNatureEtat<>'IPR'));
  SetControlVisible ('TDATEARRETE', (FNatureEtat<>'IPR'));

  if (FCodeEtat='DOT') or (FCodeEtat='DEC') or (FCodeEtat='DFI') or (FCodeEtat='DDE') or (FCodeEtat='LSI') then
    SetControlText('I_NATUREIMMO','PRO')
  else if (FCodeEtat='DCB') or (FCodeEtat='IEG') or (FCodeEtat='IDS') then   // modif mbo fq 13036 or (FCodeEtat='ECB') then
  // BTY 02/06 FQ 13035 Dotations simulées
    SetControlText('I_NATUREIMMO','CB');

  // MVG 30/10/06 FQ 19048
  if (FCodeEtat='IEG') then SetControlText('I_NATUREIMMO','CB;LOC');

  SetControlVisible('FLISTE',(FCodeEtat <> 'ETI') or (FCodeEtat <> 'FIN'));

  if ((FNatureEtat='ITE') or (FNatureEtat='ITF') or (FNatureEtat='IRE') ) then
    SetControlProperty('I_NATUREIMMO','Plus',' AND CO_CODE<>"CB" AND CO_CODE<>"LOC"')
  else SetControlProperty('I_NATUREIMMO','Plus','');

  // Mise à jour des contextes d'aide
  AssignHelpContext;
end;


procedure TOF_AMEDITION.RenseigneComboRupture ( CB : THValComboBox; stSauf : string );
var i,j : integer;
    iMax : integer;
    stTmp : string;
begin
  stTmp := '';
  if CB =  nil then iMax := 2 else iMax := 1;
  for i:=1 to iMax do
  begin
    if CB = nil then CB := THValComboBox(GetControl('RUPTURE'+IntToStr(i),True))
    else stTmp := CB.Value;
    CB.OnChange := nil;
    if CB <> nil then
    begin
      CB.Items.Clear;
      CB.Values.Clear;
      for j:=0 to FTRupture.Detail.Count - 1 do
      begin
        if (stSauf ='' ) or (Copy(FTRupture.Detail[j].GetValue('VALUE'),1,9)<> Copy(stSauf,1,9)) then
        begin
          CB.Items.Add ( FTRupture.Detail[j].GetValue('ITEM'));
          CB.Values.Add ( FTRupture.Detail[j].GetValue('VALUE'));
        end;
      end;
      if stTmp<> '' then CB.Value := stTmp else CB.ItemIndex := 0;
      CB.OnChange := OnChangeRupture;
      CB := nil;
    end;
  end;
end;

procedure TOF_AMEDITION.ChargeListeRupture;
{$IFDEF SERIE1}
  Procedure GetLibelleTableLibre(Pref : String ; LesLib : HTStringList) ;
  Var QLoc : TQuery ;
  BEGIN
    QLoc:=OpenSql('select CC_CODE,CC_LIBELLE From CHOIXCOD where CC_TYPE="LCL" and CC_CODE Like"'+Pref+'%" Order by CC_CODE',True) ; //XMG 25/03/03
    LesLib.Clear ;
    While Not QLoc.Eof do
      BEGIN LesLib.Add(QLoc.Fields[1].AsString) ; QLoc.Next ; END ;
    Ferme(QLoc) ;
  END ;
{$ENDIF}

var T : TOB;
    i : integer;
    LesLibelles : HTStringList;
    stLibelle : string;
    {$IFDEF SERIE1}
    {$ELSE}
    St : string;
    Q : TQuery;
    {$ENDIF}
begin
  FTRupture := TOB.Create ('', nil, -1);
  { Pas de rupture }
  T := TOB.Create ('', FTRupture, - 1);
  T.AddChampSupValeur ('ITEM',TraduireMemoire('<<Aucune>>'));
  T.AddChampSupValeur ('VALUE','');
//  { Rupture sur compte ou regroupement }
//  T := TOB.Create ('', FTRupture, - 1);

  // FQ 13036 - pour l'état des engagements la check regroupement est tjrs dispo
  // mais on ne propose ni regroupement ni compte dans les combo ruptures
  // FQ 13035 - BTY 02/06 Même chose Dotations simulées des CB
if (FNatureEtat <> 'IEG') then
begin
 { Rupture sur compte ou regroupement }
  T := TOB.Create ('', FTRupture, - 1);
  if (GetCheckBoxState('AVECREGROUPEMENT')=cbChecked) then
  begin   { Regroupement est coché ==> on propose la rupture sur compte }
    T.AddChampSupValeur ('ITEM',TraduireMemoire('Compte'));
    //T.AddChampSupValeur ('VALUE','I_COMPTEREF');
    // BTY 02/06 I_COMPTEREF=I_COMPTEIMMO pour les maquettes sauf celles de Crédit-Bail
    // => Pour IDS I_COMPTEREF=I_COMPTELIE => mettre I_COMPTEIMMO
    if (FNatureEtat = 'IDS') then
       T.AddChampSupValeur ('VALUE','I_COMPTEIMMO')
    else
       T.AddChampSupValeur ('VALUE','I_COMPTEREF');
  end else
  begin   { Regroupement n'est pas coché ==> on propose la rupture sur regroupement }
    T.AddChampSupValeur ('ITEM',TraduireMemoire('Regroupement'));
    T.AddChampSupValeur ('VALUE','I_GROUPEIMMO');
  end;
end;

  { Rupture sur établissement }
  T := TOB.Create ('', FTRupture, - 1);
  T.AddChampSupValeur ('ITEM',TraduireMemoire('Etablissement'));
  T.AddChampSupValeur ('VALUE','I_ETABLISSEMENT');
  { Rupture sur lieu géographique }
  T := TOB.Create ('', FTRupture, - 1);
  T.AddChampSupValeur ('ITEM',TraduireMemoire('Lieu géographique'));
  T.AddChampSupValeur ('VALUE','I_LIEUGEO');
  { Ruptures sur tables libres }
  LesLibelles:=HTStringList.Create ;
  GetLibelleTableLibre('I',LesLibelles);
  for i:=0 to LesLibelles.Count-1 do
  begin
    {$IFDEF SERIE1}
    stLibelle := LesLibelles.Strings[i];
    {$ELSE}
    St:=LesLibelles.Strings[i] ;
    stLibelle := ReadTokenSt(St);
    if St = 'X' then
    {$ENDIF}
    begin
      T := TOB.Create ('', FTRupture, - 1);
      T.AddChampSupValeur ('ITEM',stLibelle);
      {$IFDEF SERIE1}
      T.AddChampSupValeur ('VALUE', 'I_CHAMPLIBRE'+IntToStr(i));
      {$ELSE}
      T.AddChampSupValeur ('VALUE', 'I_TABLE'+IntToStr(i));
      {$ENDIF}
    end;
  end;
  LesLibelles.Free;
  { Ruptures sur sections }
  {$IFDEF SERIE1}
  {$ELSE}
  Q := OpenSQL ('SELECT DISTINCT(V_NATURE) FROM VENTIL WHERE V_NATURE LIKE "IM%"',True);
  try
    while not Q.Eof do
    begin
      T := TOB.Create ('', FTRupture, - 1);
      T.AddChampSupValeur ('ITEM', 'Axe '+(Q.Fields[0].AsString)[Length(Q.Fields[0].AsString)]);
      T.AddChampSupValeur ('VALUE', 'S_SECTION'+(Q.Fields[0].AsString)[Length(Q.Fields[0].AsString)]);
      Q.Next;
    end;
  finally
    Ferme (Q);
  end;
  {$ENDIF}
end;

procedure TOF_AMEDITION.MajLookUpSection(stAxe: string);
begin
  SetControlVisible('TS_SECTION',stAxe<> '');
  SetControlVisible('S_SECTION',stAxe<> '');
  SetControlVisible('TS_SECTION_',stAxe<> '');
  SetControlVisible('S_SECTION_',stAxe<> '');
  if stAxe='A1' then
  begin
    SetControlProperty ( 'S_SECTION','DataType','TZSECTION');
    SetControlProperty ( 'S_SECTION_','DataType','TZSECTION');
  end
  else if stAxe <> '' then
  begin
    SetControlProperty ( 'S_SECTION','DataType','TZSECTION'+stAxe[2]);
    SetControlProperty ( 'S_SECTION_','DataType','TZSECTION'+stAxe[2]);
  end else
  begin
    SetControlText('S_SECTION','');
    SetControlText('S_SECTION_','');
  end;
end;

procedure TOF_AMEDITION.OnChangeFiltre(Sender: TObject);
begin
  {$IFNDEF SERIE1}
  Inherited ;  // ajout mbo 19.01.06 pour correction warning à la compil
  {$ENDIF}
  THValComboBox(GetControl('RUPTURE1')).OnChange := nil;
  THValComboBox(GetControl('RUPTURE2')).OnChange := nil;
  TFQRS1(ECRAN).FFiltresChange(Sender);
  OnChangeRupture(GetControl('RUPTURE1'));
  OnChangeRupture(GetControl('RUPTURE2'));
  InitZonesSuivantEtat;
  InitRuptures;
  THValComboBox(GetControl('RUPTURE1')).OnChange := OnChangeRupture;
  THValComboBox(GetControl('RUPTURE2')).OnChange := OnChangeRupture;

  // BTY 09/06 Revoir les établissements au cas où restauration incompatible
  // avec les restrictions utilisateur
  {$IFNDEF SERIE1}
  ImControlEtab (THMultiValComboBox(GetControl('I_ETABLISSEMENT')) );
  {$ENDIF}
end;

procedure TOF_AMEDITION.OnExitCompteRef(Sender: TObject);
begin
  if Trim(THEdit(Sender).Text) = '' then Exit;
  if Length(THEdit(Sender).Text) < VHImmo^.Cpta[fbGene].Lg then
    THEdit(Sender).Text := ImBourreEtLess( THEdit(Sender).Text, fbGene);
end;

// FQ 19036  Pour la Taxe professionelle, la borne supérieure de la date achat
// ne doit pas dépasser la date d'arrêté
procedure TOF_AMEDITION.OnExitDateAchatFin(Sender: TObject);
begin
  if  (FNatureEtat='ITP') then
    begin
      if StrToDate(GetControlText('I_DATEPIECEA_')) > StrToDate(GetControlText('DATEARRETE')) then
         SetControlText('I_DATEPIECEA_', GetControlText('DATEARRETE'));
    end;
end;

// FQ 19036  Pour la Taxe professionelle, la borne supérieure de la date achat
// ne doit pas dépasser la date d'arrêté
procedure TOF_AMEDITION.OnExitDateArrete(Sender: TObject);
begin
  if  (FNatureEtat='ITP') then
    begin
      if StrToDate(GetControlText('I_DATEPIECEA_')) > StrToDate(GetControlText('DATEARRETE')) then
         SetControlText('I_DATEPIECEA_', GetControlText('DATEARRETE'));
    end;
end;

procedure TOF_AMEDITION.OnClickRegroupement(Sender: TObject);
begin

// fq 13036 - pour l'état des engagements on peut choisir regroupement +
// 2 autres niveaux de rupture
   if (FNatureEtat <> 'IEG') then
   begin
     { Réinitialisation des choix de rupture }
     SetControlText('XX_RUPTURE1','');
     SetControlText('XX_RUPTURE2','');
     SetControlText('XX_RUPTURE3','');
     ChargeListeRupture;
     RenseigneComboRupture ( nil, '');
     THValComboBox(GetControl('RUPTURE1')).ItemIndex := 0;
     OnChangeRupture(THValComboBox(GetControl('RUPTURE1')));
     THValComboBox(GetControl('RUPTURE2')).ItemIndex := 0;
     OnChangeRupture(THValComboBox(GetControl('RUPTURE2')));
     // InitRuptures; BTY 02/06
   end;
   // BTY 02/06 Attention : Check Regroupement actionnée
   // => toujours recalculer les nivx de rupture
   InitRuptures;
end;

// mbo fq 14430
procedure TOF_AMEDITION.OnCLickMutcompte(Sender: TObject);
begin
    if (GetCheckBoxState('MUTCOMPTE')<>cbChecked) then
       SetControlText('XX_WHERE', '')
    else
       SetControlText('XX_WHERE', 'IL_CPTEMUTATION <> ""')
end;
// fin mbo

procedure TOF_AMEDITION.InitRuptures;
var bAvecRegroupement : boolean;
{$IFDEF SERIE1}
    TypeOperation : string;
    TDateW : TDateTime;
{$ENDIF}

begin
  bAvecRegroupement := (GetCheckBoxState('AVECREGROUPEMENT')=cbChecked);
  // if bAvecRegroupement then SetControlText ('XX_RUPTURE3','I_GROUPEIMMO')
  // else SetControlText ('XX_RUPTURE3','I_COMPTEREF');
  // BTY 02/06 FQ 13035 I_COMPTEREF=I_COMPTELIE pour les CB => mettre I_COMPTEIMMO
  if bAvecRegroupement then
     SetControlText ('XX_RUPTURE3','I_GROUPEIMMO')
  else if (FNatureEtat='IDS') then
       SetControlText ('XX_RUPTURE3','I_COMPTEIMMO')
       else
       SetControlText ('XX_RUPTURE3','I_COMPTEREF');
  if ((FNatureEtat='ITE') or (FNatureEtat='ITV') or (FNatureEtat='ITF') or (FNatureEtat='ITR') or (FNatureEtat='IRE')) then
  begin
    // BTY FQ 18828 ORDER BY sur l'alias TYPEOPERATION ne fonctionne pas en S1
    //SetControlText('XX_RUPTURE4','TYPEOPERATION');
    {$IFDEF SERIE1}
    TDateW := StrToDate(GetControlText('DATEDEBEXO'));
    TypeOperation :=
     'IIF(IL_DATEOP<"' + USDateTime(TDateW)+ '", IL_TYPEOP||"_1", IL_TYPEOP||"_2")';
    SetControlText('XX_RUPTURE4', TypeOperation);
    {$ELSE}
    SetControlText('XX_RUPTURE4','TYPEOPERATION');
    {$ENDIF}
    SetControlText('XX_RUPTURE5','I_IMMO');
    SetControlText('XX_RUPTURE6','');
  end else if (FNatureEtat='IPM') or (FNatureEtat='ISO') then
  begin
    SetControlText('XX_RUPTURE3','IL_MOTIFCES');
    if bAvecRegroupement then SetControlText('XX_RUPTURE4','I_GROUPEIMMO')
    else SetControlText('XX_RUPTURE4','I_COMPTEREF');
    SetControlText('XX_RUPTURE5','I_IMMO');
    SetControlText('XX_RUPTURE6','');
  end else if (FNatureEtat='ITP') then
  begin
    // BTY FQ 18828 ORDER BY sur l'alias LOCATION ne fonctionne pas en S1
    //SetControlText('XX_RUPTURE3','LOCATION');
    {$IFDEF SERIE1}
    SetControlText('XX_RUPTURE3', 'IIF(I_NATUREIMMO="LOC","OUI","NON")' );
    {$ELSE}
    SetControlText('XX_RUPTURE3','LOCATION');
    {$ENDIF}
    SetControlText('XX_RUPTURE4','I_NATUREBIEN');
    if bAvecRegroupement then SetControlText('XX_RUPTURE5','I_GROUPEIMMO')
    else SetControlText('XX_RUPTURE5','I_COMPTEREF');
    SetControlText('XX_RUPTURE6','I_IMMO');
  end else if (FNatureEtat='IVE') then
  begin
    SetControlText('XX_RUPTURE4','IL1.IL_DATEOP');
    SetControlText('XX_RUPTURE5','I_IMMO');
    SetControlText('XX_RUPTURE6','');
  end else if (FnatureEtat='IEG') then             // ajout mbo - fq 13036
  begin
    SetControlText('XX_RUPTURE4','I_COMPTEIMMO');
    SetControlText('XX_RUPTURE5','I_IMMO');
    SetControlText('XX_RUPTURE6','');
  end else if (FnatureEtat='IDD') then             // FQ 18889
  begin
    SetControlText('XX_RUPTURE4','I_IMMO');
    SetControlText('XX_RUPTURE5','IZ_DATE');
    SetControlText('XX_RUPTURE6','');
  end
  else begin
    SetControlText('XX_RUPTURE4','I_IMMO');
    SetControlText('XX_RUPTURE5','');
    SetControlText('XX_RUPTURE6','');
  end;
  // BTY 01/06 FQ 17172 Ne pas écraser la Rupture 6 sur code immo pour la Taxe Pro
  // SetControlText('XX_RUPTURE6','');
  //if (FNatureEtat <>'ITP') then
  //   SetControlText('XX_RUPTURE6','');
end;

{$IFDEF SERIE1}
procedure TOF_AMEDITION.OnChangeChampLibre(Sender: TObject);
begin
  if THValComboBox(Sender).Name = 'TABLELIBRE1' then
    SetControlText('I_TABLE0', GetControlText('TABLELIBRE1'))
  else if THValComboBox(Sender).Name = 'TABLELIBRE2' then
    SetControlText('I_TABLE1', GetControlText('TABLELIBRE2'))
  else if THValComboBox(Sender).Name = 'TABLELIBRE3' then
    SetControlText('I_TABLE2', GetControlText('TABLELIBRE3'));
end;
{$ENDIF}


//==================================================================
//Gestion des restrictions utilsateurs sur les établissements
//==================================================================


// Récupération de la combo gérant les établissements
//{$IFNDEF SERIE1}
{function TOF_AMEDITION.GetComboEtab : THMultiValComboBox;
var
  n : Integer;
begin
  if FComboEtab = nil then begin
    for n := 0 to Ecran.ComponentCount - 1 do
      if (Pos('ETAB', Ecran.Components[n].Name) > 0) and
         (Ecran.Components[n] is THMultiValComboBox) then begin
        FComboEtab := THMultiValComboBox(Ecran.Components[n]);
        Break;
      end;
  end;
  Result := FComboEtab;
end;
}
// Gestion des établissements
{procedure TOF_AMEDITION.GereEtablissement;
begin
  if Assigned(ComboEtab) then begin
    //Si l'on ne gère pas les établissement ...
    if not VH^.EtablisCpta  then begin
      //... on affiche l'établissement par défaut
      SetControlText(ComboEtab.Name, VH^.EtablisDefaut);
      //... on désactive la zone
      SetControlEnabled(ComboEtab.Name, False);
    end

    //On gère l'établisement, donc ...
    else begin
      //... On commence par regarder les restrictions utilisateur
      PositionneEtabUser(ComboEtab);
      //... s'il n'y a pas de restrictions, on reprend le paramSoc
      if GetControlText(ComboEtab.Name) = '' then begin
        //... on affiche l'établissement par défaut
        SetControlText(ComboEtab.Name, VH^.EtablisDefaut);
        //... on active la zone
        SetControlEnabled(ComboEtab.Name, True);
      end;
    end;
  end;
end;
}
// on s'assure que le filtre ne va pas à l'encontre des restrictions utlisateurs
{procedure TOF_AMEDITION.ControlEtab;
var
  Eta : string;
begin
  if not Assigned(ComboEtab) then Exit;
  //S'il n'y a pas de gestion des établissement, logiquement, on ne force pas l'établissement !!!
  if not VH^.EtablisCpta then Exit;

  Eta := EtabForce;
  //S'il y a une restriction utilisateur et qu'elle ne correspond pas au contenu de la combo ...
  if (Eta <> '') and (Eta <> GetControlText(ComboEtab.Name)) then begin
    //... on affiche l'établissement des restrictions
    SetControlText(ComboEtab.Name, Eta);
    //... on désactive la zone
    SetControlEnabled(ComboEtab.Name, False);
  end;
end;
}
//{$ENDIF}


Initialization
  registerclasses ( [ TOF_AMEDITION ] ) ;
end.
