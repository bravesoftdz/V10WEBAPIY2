unit MajHalley;

interface

uses ComCtrls, StdCtrls;

function MAJHalleyAvant(VSoc: Integer; MajLab1, MajLab2: TLabel; MajJauge: TProgressBar): boolean;
function MAJHalleyApres(VSoc: Integer; MajLab1, MajLab2: TLabel; MajJauge: TProgressBar): boolean;
function MAJHalleyIsPossible(VSoc: Integer; MajLab1, MajLab2: TLabel; MajJauge: TProgressBar): boolean;
function MAJHalleyPendant(VSoc: Integer; NomTable: string): boolean;

procedure MAJStandardPaie;

implementation

uses HEnt1, HCtrls, MajTable, Forms, DB,
{$IFDEF VER150} {D7}
  Variants,
{$ENDIF}
  {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
  MajStruc, HQry, SysUtils, Classes,
  DBCtrls, ParamSoc, Dialogs,

  MajHalleyUtil, MajHalleyProc,
  uHListe,Ent1, UtilGa, HmsgBox, STKMoulinette, yTarifsBascule, Wjetons,yTarifsCommun, GcLienPiece,
  uTob , utablesged ;

var L1, L2: TLabel;
  J: TProgressBar;
  Vers_affaire,Vers_Activite,vers_ligne, vers_article,
  Vers_ParCaisse, Vers_ParPiece, Vers_JoursEtabEvt, Vers_ClavierEcran, Vers_JoursCaisse: Integer;
  Vers_TarifPer, Vers_Dispo, Vers_TarifMode, Vers_ListeInvent, Vers_Etabliss, Vers_TarifTypMode, Vers_Depots: Integer;
  Vers_ParPieceCompl, Vers_ProfilArt, Vers_JoursEtab, Vers_CtrlCaisMt, Vers_Tache, Vers_ArticleCompl,
  Vers_PropTransfEnt, Vers_ExportAscii, Vers_Piece,Vers_Affcdeentete,Vers_Affcdeligne,Vers_AfRevision,Vers_AfCumul,
  Vers_Afftiers,Vers_Afplanning : Integer;

  IsDossierPCL : boolean;

  TobWPF: Tob; { MajAvant690 et MajVer 690: Mémorisation d'infos de tables qui doivent disparaitre (WOA, WEA) pour transfert dans nouvelle table WPF}

  {==============================================================================}
  {=========================== PROCEDURES MAJ AVANT =============================}
  {==============================================================================}

procedure MajAvant545;
begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE LIKE "RT%"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "AFMULFACTIERSAFF"');
end;

procedure MajAvant547;
begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE LIKE "PG%"');
end;

procedure MajAvant550;
begin
  Vers_ParCaisse := TableToVersion('PARCAISSE');
  Vers_ParPiece := TableToVersion('PARPIECE');
  Vers_JoursEtabEvt := TableToVersion('JOURSETABEVT');
  Vers_ClavierEcran := TableToVersion('CLAVIERECRAN');
  Vers_JoursCaisse := TableToVersion('JOURSCAISSE');
  Vers_TarifPer := TableToVersion('TARIFPER');
  Vers_Dispo := TableToVersion('DISPO');
  Vers_TarifMode := TableToVersion('TARIFMODE');
  Vers_Listeinvent := TableToVersion('LISTEINVENT');
  Vers_Etabliss := TableToVersion('ETABLISS');
  Vers_ParPieceCompl := TableToVersion('PARPIECECOMPL');
end;

procedure MajAvant560;
begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE LIKE "PGMULMVTAB%"');
end;

procedure MajAvant561;
begin
  // ajout champ d'index
  if TableExiste('AFCUMUL') then AddChamp('AFCUMUL', 'ACU_NUMPIECE', 'VARCHAR(35)', 'ACU', False, '', 'Numéro de pièce', 'LD', False);
end;

procedure MajAvant562;
begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "RTMULACTIONSPRO"');
  if TableExiste('MOTIFSORTIEPAY') then
  begin
    ExecuteSQL('DELETE FROM MOTIFSORTIEPAY');
    AddChamp('MOTIFSORTIEPAY', 'PMS_PREDEFINI', 'COMBO', 'PMS', False, 'CEG', 'Standard', '', False);
    AddChamp('MOTIFSORTIEPAY', 'PMS_NODOSSIER', 'VARCHAR(17)', 'PMS', False, '000000', 'Standard', '', False);
  end;
  if TableExiste('MOTIFENTREESAL') then
  begin
    ExecuteSQL('DELETE FROM MOTIFENTREESAL');
    AddChamp('MOTIFENTREESAL', 'PME_PREDEFINI', 'COMBO', 'PME', False, 'CEG', 'Standard', '', False);
    AddChamp('MOTIFENTREESAL', 'PME_NODOSSIER', 'VARCHAR(17)', 'PME', False, '000000', 'Standard', '', False);
  end;

  if TableExiste('JUTYPEPER') then ExecuteSQL('DELETE FROM JUTYPEPER');

end;

procedure MajAvant565;
begin
  Vers_TarifMode := TableToVersion('TARIFMODE');
  Vers_TarifTypMode := TableToVersion('TARIFTYPMODE');
end;

procedure MajAvant566;
begin
  Vers_TarifMode := TableToVersion('TARIFMODE');
  Vers_TarifTypMode := TableToVersion('TARIFTYPMODE');
  Vers_ProfilArt := TableToVersion('PROFILART');
end;

procedure MajAvant569;
var SQL: string;
begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "PGDUCSENTETE"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "PGINSTITUTIONS"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "PGDUCSAFFECT"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "PGCOTSANSDUCS"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "PGLANCEBULLAVEC"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "PGEMULMVTABSR"');
  if TableExiste('CUMULRUBRIQUE') then
  begin
    SQL := 'UPDATE CUMULRUBRIQUE SET PCR_PREDEFINI="STD" WHERE PCR_PREDEFINI="CEG" ';
    SQL := SQL + ' AND ((PCR_RUBRIQUE like "%1") OR (PCR_RUBRIQUE like "%3")) ';
    SQL := SQL +
      ' AND NOT EXISTS (SELECT PCR_RUBRIQUE FROM CUMULRUBRIQUE C2 WHERE C2.PCR_PREDEFINI="STD" AND PCR_NATURERUB=C2.PCR_NATURERUB AND PCR_RUBRIQUE=C2.PCR_RUBRIQUE AND PCR_CUMULPAIE=C2.PCR_CUMULPAIE )';
    ExecuteSQL(SQL);
  end;
  MajJurPredef;
end;

procedure MajAvant582;
begin
  Vers_TarifMode := TableToVersion('TARIFMODE');
  Vers_TarifTypMode := TableToVersion('TARIFTYPMODE');
  Vers_ProfilArt := TableToVersion('PROFILART');
  Vers_ParCaisse := TableToVersion('PARCAISSE');
  Vers_ParPieceCompl := TableToVersion('PARPIECECOMPL');
  Vers_JoursEtab := TableToVersion('JOURSETAB');
  Vers_JoursCaisse := TableToVersion('JOURSCAISSE');
  Vers_CtrlCaisMt := TableToVersion('CTRLCAISMT');
  TestDropTable('HRPREFACT'); // recréée proprement dans cette version
  TestDropTable('EAITRANSFERTS');
  TestDropTable('EAISATELLITES');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "AFMULFACTIERSAFF"');
  if TableExiste('HISTOSALARIE') and (tableToVersion('HISTOSALARIE') < 111) then
  begin
    AddChamp('HISTOSALARIE', 'PHS_DATEAPPLIC', 'DATE', 'PHS', False, Idate1900, '', '', False);
    if ChampPhysiqueExiste('HISTOSALARIE', 'PHS_DATEAPPLICATION') then
      ExecuteSQL('UPDATE HISTOSALARIE SET PHS_DATEAPPLIC=PHS_DATEAPPLICATION');
  end;
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE="PGDUCSENTETE" OR LI_LISTE="PGVISITEMEDICALE" OR LI_LISTE="PGMULMVTSALARIE" OR LI_LISTE="PGVENTILAREM" OR LI_LISTE="PGMULSALDAS" OR LI_LISTE="PGDADSPERIODE" OR LI_LISTE = "PGTAUXAT" OR LI_LISTE = "PGMULHISTOSAL"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE="RTRESSOURCES" OR LI_LISTE="RTMULTIERSMAILING" OR LI_LISTE="RTTIERSMAILINGFIC" OR LI_LISTE="RTMULTIERSACTIONS" OR LI_LISTE="RTMULQUALITECONTA"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE="GCCATANA"');
  InsertChoixExt('SPE', '160101', 'GC;GCETATLIBRE;;;;', 'FIC', '');
  TestDropTable('CPPRORATA');
end;

procedure MajAvant585;
begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCCATANA","GCTRACABILITELIG","GCTRACABILITELOT","GCTRACABILITESERI","GCVERIFLIGNESERIE","GCSERIESTOCK","GCSERIELIGNE","GCSERIECOMPO","GCMULPIECE","GCDISPOSERIE")');
end;

procedure MajAvant590;
begin
  Vers_ParCaisse := TableToVersion('PARCAISSE');
  Vers_Jourscaisse := TableToVersion('JOURSCAISSE');
  Vers_ParPiece := TableToVersion('PARPIECE');

  ExecuteSQL('DELETE FROM PARPIECE WHERE GPP_NATUREPIECEG = "FCF"');
  ExecuteSQL('DELETE FROM PARPIECEGRP WHERE GPR_NATUREPIECEG = "CF"');

  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "GCMULPIEDECHEFO"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "GCMULPIECARTFO"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "GCAFFECTTRANSFERT"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "GCMULPIECEDEVFO"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "GCMULANNULTICFO"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "GCMULCHQDIFFO"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "GCVMULTIERSFO"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "GCMULPIECEFO"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "GCSITUFLASHFO"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "GCRECAPVENFO"');

  // DC 23/10/02
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCMULARTDIM_MODE","GCMULARTICLE_MODE")');
end;

procedure MajAvant591;
begin
  Vers_Tache := TableToVersion('TACHE');
  if TableToVersion('TACHE') < 107 then
  begin
    TestDropTable('TACHE');
    if TableToVersion('TACHERESSOURCE') < 102 then TestDropTable('TACHERESSOURCE');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE in ("AFTACHE_MUL","AFLISTETACHES","AFLISTETACHESRES")');
    if TableToVersion('AFCUMUL') < 107 then TestDropTable('AFCUMUL');
  end;
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE in ("RTMULTIERSACTIONS","YYCONTACTTIERS","RTMULQUALITECONTA","AFMULPIECEF1","PGDUCSINIT","PGENVOISOCIAL")');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE in ("PGVALIDRESPONSFOR","PGTVRESPONSFOR","PGASSEDICFAITES","PGCOMPETENCESTAGE","PGCONTRATTRAV","PGDUCSAFFECT","PGMULCOEFF","PGVENTILACOTISAT","PGVENTILAREM")');

  //TRESO
  if IncoherTable('COUPLESVM', 'TCM') then MAJDropTable('COUPLESVM');
  if IncoherTable('COURSVM', 'TCS') then MAJDropTable('COURSVM');
  if IncoherTable('PARAMVM', 'TPV') then MAJDropTable('PARAMVM');
  if IncoherTable('VALEURSMOBILIERES', 'TVM') then MAJDropTable('VALEURSMOBILIERES');
  if IncoherTable('SOLDEBANCAIRE', 'TSB') then MAJDropTable('SOLDEBANCAIRE');

  //PCL
  MoveCommunToChoixcode('VOT', 'VOT')
end;

procedure MajAvant595;
var SQL: string;
(*
  procedure VerifTableImmo(NomTable, NomChamp: string);
  begin
    if NomTable = '' then Exit;
    if NomChamp = '' then Exit;
    if not TableExiste(NomTable) then exit;
    BEGINTRANS;
    ExecuteSQL('UPDATE ' + NomTable + ' SET ' + NomChamp + '="" WHERE ' + NomChamp + '<>"" ');
    COMMITTRANS;
  end;
*)
begin
  // GRC
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE in  ("RTMULOPERATIONS","RTMULCONTMAILING","RTMULTIERSMAILING","RTMULQUALITE")');

  //GI-GA
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE in ("AFTACHE_MUL","AFMULPLANNING","AFMULFACTIERSAFF","AFMULFACTAFFAPP","AFLIGPLANNING","AFTACHELIG","AFFTIERS","AFMUL_EACTIVITE1","AFREGRLIGNE_MUL")');
  GIGAChangeNomAffaire;

  // MODE
  Vers_ParCaisse := TableToVersion('PARCAISSE');
  Vers_Jourscaisse := TableToVersion('JOURSCAISSE');
  Vers_ParPiece := TableToVersion('PARPIECE');
  Vers_ArticleCompl := TableToVersion('ARTICLECOMPL');
  Vers_Etabliss := TableToVersion('ETABLISS');
  Vers_Depots := TableToVersion('DEPOTS');

  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE in ("GCMULARTDIM_MODE", "GCMULARTICLE_MODE", "GCMULDISPDIM_MODE", "GCMULDISPDIM_MOS5", "GCMULDISPART_MODE", "GCMULDISPART_MOS5", "MULDISPO_DIST", "ARTICLEDIM_MODE", "ARTICLE_MODE")');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "GCMULPIECARTFO"');

  (*
  // GP : Remise à Niveau Structure IMMO
  VerifTableImmo('IAIDSAI', 'IAI_CODEETAB');
  VerifTableImmo('IMOREF', 'IRF_USERSAI');
  VerifTableImmo('IMOREF', 'IRF_USERMAJ');
  VerifTableImmo('IMOREF', 'IRF_CODEETAB');
  VerifTableImmo('IMOVI1', 'IV1_CODEETAB');
  *)

  // PAIE
  SQL := 'DELETE FROM LISTE WHERE LI_LISTE IN ("PGMULMVTABS","PGEMULMVTABS","PGEMULMVTABSR"';
  SQL := SQL + ',"PGMODIFCOEFFSAL","PGMODIFCOEFF","PGCONVOCFORMATION","PGFORMATIONS","PGMULINSCFOR"';
  SQL := SQL + ',"PGPLAFONDFORM","PGMULFRAISFORMATI","PGPRESENCEFORM","PGSAISIEFRAISFORM"';
  SQL := SQL + ',"PGSAISIFORMATIONS","PGVALIDINSSESSION","PGCONSULTINSCBUD","PGINSCFORBUDGET"';
  SQL := SQL + ',"PGVALIDRESPONSFOR","PGSTAGESESSION","PGTVSESSIONSTAGE","PGSESSIONANIMAT"';
  SQL := SQL + ',"PGSESSIONSTAGE","PGSTAGEFORM","PGCOMPETENCESTAGE","PGOBJECTIFSTAGE","PGTVSTAGE","PGERECAPSAL")';
  ExecuteSQL(SQL);

  //Compta
  SQL := 'DELETE FROM LISTE WHERE LI_LISTE IN("CPDECBOREDT","CPDECBOREDTNC","CPDECCHEQUEEDT","CPDECCHEQUEEDTNC","CPDECDIV", ';
  SQL := SQL + '"CPDECTOUS","CPDECTRAITEBQE","CPDECTRAITEENC","CPDECTRAITEESC","CPDECTRAITEPOR",';
  SQL := SQL + '"CPDECVIREDT","CPDECVIREMENT","CPDECVIRIN","CPDECVIRINEDT", ';
  SQL := SQL + '"CPENCCBBQE","CPENCCBPOR","CPENCCHQBQE","CPENCCHQPOR","CPENCDIV", ';
  SQL := SQL + '"CPENCPREEDT","CPENCPRELEVEMENT","CPENCTOUS","CPENCTRAITEBQE","CPENCTRAITEEDT", ';
  SQL := SQL + '"CPENCTRAITEEDTNC","CPENCTRAITEENC","CPENCTRAITEESC","CPENCTRAITEPOR", ';
  SQL := SQL + '"TRFDOSSIER","CPCONSGENE","CPCONSAUXI" )';
  ExecuteSQL(SQL);
end;

procedure MajAvant602;
begin
  if TableToVersion('ACTIVITEPAIE') < 102 then
  begin
    TestDropTable('ACTIVITEPAIE');
  end;
  if TableExiste('ACTIVITE') then
  begin
    AddChamp('ACTIVITE', 'ACT_NUMLIGNEUNIQUE', 'INTEGER', 'ACT', False, 0, 'num unique de la ligne par affaire', '', False);
    // PL le 07/03/03 :
    GIGAIndexActivite;
  end;
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("AFMULMODIFLOTAFFA","PRT_JOURFERIE","AFALIGNAFF","AFLIGPLANNING","AFMULAFFAIREMULTI","AFMULACTIVITEGLO2","AFMULACTIVITEVIS1","AFMULFACTIERSAFF")');
  // MODE
  Vers_ParCaisse := TableToVersion('PARCAISSE');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "GCMULPIECE"');
  Vers_PropTransfEnt := TableToVersion('PROPTRANSFENT');
  Vers_ExportAscii := TableToVersion('EXPORTASCII_ENTETE');
  Vers_Piece := TableToVersion('PIECE');
end;

procedure MajAvant604;
begin
MajGPAOAvant;
end;

procedure MajAvant605;
begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "AFMULACTIVITECON1"');
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE LIKE "PGCONVCOL%"');
PGMinConvPaie;
if TableExiste('QDONFILTRE') AND (TableToVersion('QDONFILTRE') < 5) AND (not ChampPhysiqueExiste('QDONFILTRE','QDO_CHAINE')) then
  ExecuteSQL('DELETE From DECHAMPS Where DH_NOMCHAMP="QDO_CHAINE"') ;
end;

procedure MajAvant606;
begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE like "BT%"') ;
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE in  ("GCPARFOULIG","AFLISTETACHES","PGDUCSENTETE")') ;
ExecuteSQL('DELETE FROM GRAPHS WHERE GR_GRAPHE = "TBTANALDEV"') ;
ExecuteSQL('DELETE FROM PARPIECE WHERE GPP_NATUREPIECEG IN ("ABT","DBT","FBT","ETU","CBT","PBT")') ;
end;

procedure MajAvant607;
begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE in  ("RTMULPROJETS","RTMULPERSPRECH","RTMULACTIONSRAP","PGDUCSENTETE","GCDISPOSERIE")') ;
end;

procedure MajAvant608;
begin
  Vers_Affaire := TableToVersion('AFFAIRE');
  Vers_Activite := TableToVersion('ACTIVITE');
  Vers_Article := TableToVersion('ARTICLE');
  Vers_ligne := TableToVersion('LIGNE');
end;

procedure MajAvant609;
begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE in  ("GCMULCONSUMOUV","PRT_PGORGANISME","PGTICKETBULL")') ;
end;

procedure MajAvant610;
begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE in  ("GCMULPIECE","PGCDETICKET","PGTICKETBULL"'
           + ',"PGCDEXIST","PGVISITEMEDICAL","PGPRESENCEFORM","PGCURSUS","PGCURSUSSESSION","PGMULSALAIREFORM")') ;

Vers_PropTransfEnt := TableToVersion('PROPTRANSFENT');
end;

procedure MajAvant612;
begin
MajGPAOAvant612;
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE in  ("AFMULRESSOURCELOT")') ;
// GI-GA
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("AFMULAPPRECON","AFREVMODIFCOEF","AFSAISIEACTAFF","AFSAISIEACTRESSAV","AFSAISIEACTRESSCT","AFSAISIEACTRESTAV","AFSAISIEACTRESTMP")');
Vers_Activite := TableToVersion('ACTIVITE');
Vers_AfRevision := TableToVersion('AFREVISION'); //nb créer le champ VersAfrevision
//Juridique
if IsMonoOuCommune then ConvertirDateHeureEvenement;
end;

procedure MajAvant613;
begin
if getparamsoc('SO_DEVISEPRINC')='FRF' then ExecuteSQL('UPDATE DEVISE SET D_FONGIBLE="-" WHERE D_DEVISE="FRF"');
MajImoAvantPourSynchro ;
end;

procedure MajAvant614;
begin

ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WORDREGAMME","PGLANCEBULLSANS","PGPLAFONDFORM")');
//GPAO
if TableExiste('WORDREBES') then
  begin
  AddChamp('WORDREBES', 'WOB_NUMORDRE', 'INTEGER', 'WOB', False, 0, 'Numéro ordre', '', False);
  if ChampPhysiqueExiste('WORDREBES','WOB_NUMLIGNE') then ExecuteSQL('UPDATE WORDREBES SET WOB_NUMORDRE=WOB_NUMLIGNE');
  end;
if TableExiste('WORDRELIG') then
  begin
  AddChamp('WORDRELIG', 'WOL_NUMORDRE', 'INTEGER', 'WOL', False, 0, 'Numéro ordre', '', False);
  if ChampPhysiqueExiste('WORDRELIG','WOL_NUMLIGNE') then  ExecuteSQL('UPDATE WORDRELIG SET WOL_NUMORDRE=WOL_NUMLIGNE');
  end;
if TableExiste('WPARC') then
  begin
  AddChamp('WPARC', 'WPC_SERIEINTERNE', 'VARCHAR(35)', 'WPC', False, '', 'Série interne', '', False);
  AddChamp('WPARC', 'WPC_LOTINTERNE', 'VARCHAR(35)', 'WPC', False, '', 'Lot interne', '', False);
  if ChampPhysiqueExiste('WPARC','WPC_IDSERIE') then ExecuteSQL('UPDATE WPARC SET WPC_SERIEINTERNE=WPC_IDSERIE,WPC_LOTINTERNE=WPC_NUMEROLOT');
  end;
end;

procedure MajAvant615;
begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCTARIFART","F2042LISTEZONE","FMULBIENS"'
                   +',"FPDECLARATIONS","FPGP_MUL","FPLELTNAT","FPMULDECLA","FPSELDECLA"'
                   +',"PRT_PGEXERSOCIAL","PRT_PGHORAIREETAB","PRT_PGMOTIFABS","PRT_PGORGANISME"'
                   +',"PRT_PGTAUXAT","PGHISTORETENUE","PGTYPERETENUE"'
                   +',"AFINDICE_MUL","PGMULCONTRAT","PGCONTRATTRAV","RTMULSUPPRTIERS"'
                   +',"GCCONSULTSTOCK","GCCONSULTSTOCKLIG")');
//GI-GA
 Vers_AfRevision := TableToVersion('AFREVISION');
end;

procedure MajAvant616;
begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("CPDPOINTAGE","RTMULACTREPORT","PRT_MULTIEMPLOY","CPCONSECR"'
                   +',"TIERSFRAIS","UORDREBESLASSAIS","UORDREBESLASSTOC","WFAISABILITE","WFORMULE","WGAMMELIG"'
                   +',"WGAMMERES","WGAMMETET","WGAMMETET2","WJOURNALACTION","WNOMELIG","WNOMETET","WORDREBES"'
                   +',"WORDREBES2","WORDREBESL","WORDREGAMME","WORDRELAS","WORDRELIG","WORDRELIG2","WORDREPHASE"'
                   +',"WORDREPHASE2","WORDRERES","WORDRETET","WPARC","WQUANTITE","WREVISION","WREVISIONWGT"'
                   +',"WREVISIONWNT","WTEMPS","WVERSION"'
                   +',"UORDRELIG","UORDREBESLASSAIS","UORDREBESLASSTOC","WGAMMELIG","WNOMELIG","WORDREGAMME","RTTVCHAINAGE")');

end;

procedure MajAvant617;
begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("PRT_LIENAFFTIER","EDILIGNE","WREVISION","RTMULCONTMAILING","RTMULTIERSMAILING")');
end;

procedure MajAvant618;
begin
TestDropTable('INTERCOMPTA'); // recréée proprement dans cette version
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WDEPOTAPP","WORDRELIG2","PGMULBULCOMPL"'
                   +',"RTMULQUALITE","RTMULTIERSSIREN","RTMULTIERSGROUPE","RTMULTIERSPRESCRI","RTMULSUSPQUAL"'
                   +',"WGAMMECIR","GCMULPIECECOURS","GCMULARTCOM","PGREAFFCODEPCS","GCSAISIECC")');
end;

procedure MajAvant619;
Begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WORDRELIG","WORDRELIG2","WORDREBES","WORDREBES2","PGERRPCS")');
End;

procedure MajAvant620;
Begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("YYMULSELDOSS","YYMULANNDOSS"'
                   +',"YYMESSAGES_IN","YYMESSAGES_INGLOB","YYMESSAGES_OUT"'
                   +',"YYMESSAGESDOS_MUL","YYMODELESDOC_MUL","YYAGENDA"'
                   +',"GCHISTORISEPIECE","CPEEXBQ","CPEEXBQ2","PGCOMPLSALARIEMSA"'
                   +',"AFMULPROPOSITION","HLDOSRES3","GCMULPEC","WNOMEDEC","HLMULREGLEMENT","CPAFTIECPTA","MULMTIERS")');
End;

procedure MajAvant621;
Begin
ExecuteSQL('delete from liste where li_liste ="TREQUITRECRITURE"');
End;

procedure MajAvant625;
Begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WORDRECMP","WORDRECMP_VUE","WNOMEDEC","WCBNPROPOSITION"'
                   +',"GCARTTIERS","WCOLORIS","UORDRELIG","WCOMPAREWNL","WCOMPAREWNL"'
                   +',"UORDREBESLASSAIS","UORDREBESLASSTOC","UORDREBES","WORDRELIG2","WORDREPHASE2")');

ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE LIKE "YTARIFS%"');
//GIGA
Vers_AfRevision := TableToVersion('AFREVISION');
End;

procedure MajAvant626;
Begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WNOMELIG","WGAMMELIG","WGAOPERRES"'
                   +',"BTMULDEVIS","BTMULPIECE","BTLIGOUV","WORDREGAMME","WGAMMELIG")');
ExecuteSQL('DELETE FROM PARPIECE WHERE GPP_NATUREPIECEG IN ("CBT","PBT")') ;
end;

procedure MajAvant627;
Begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WNOMELIG")');
end;

procedure MajAvant631;
Begin
//mis un aglnettoieliste ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("AFMULACTIVITECON1")');
End;

procedure MajAvant633;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("QULARTBTNPRO2","QULDATEFAB"'
           +',"QULMACRO001","QULSELCHXRAL1","QULSELECTOF003"'
           +',"GCUNITELOGISTIQUE","WARTICLE","WGAMMECIR","WGAMMELIG","WGAMMETET"'
           +',"RTAFFGENACTION","RTAFFMAILINGCON","RTAFFMAILINGFIC"'
           +',"TREQUITRECRITURE","TRLANCEDOC","TRLISTETRANSAC"'
           +',"WNOMELIG","WNOMETET","WORDREGAMME")');
//TRESO
  if TableToVersion('TRECRITURE') < 15 then TestDropTable('TRECRITURE');
  if TableToVersion('COURTSTERMES') < 10 then TestDropTable('COURTSTERMES') ;
  if TableToVersion('CONDITIONFINPLAC') < 10 then TestDropTable('CONDITIONFINPLAC') ;
  if TableToVersion('CIB') < 5 then TestDropTable('CIB') ;
End;

procedure MajAvant634;
Begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WREVISIONWGT","WREVISIONWNT")');
if TableExiste('WGAMMECIR') then
  begin
  ExecuteSQL('DELETE FROM WGAMMECIR');
  end;
End;

procedure MajAvant635;
Begin
ExecuteSql ('delete from liste where li_liste like "AFSAISIEACT%"');
// mis AglNettoieListes ExecuteSql ('delete from liste where li_liste IN ("AFMULACTIVITECON1")');
  Vers_ArticleCompl := TableToVersion('ARTICLECOMPL');
  Vers_Piece := TableToVersion('PIECE');
  Vers_Activite := TableToVersion('ACTIVITE');
  Vers_AfCumul := TableToVersion('AFCUMUL'); // ajouter Vers_Afcumul en définition
End;

Procedure MajAvant637;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("RTAFFGENACTION","RTAFFMAILINGCON","RTAFFMAILINGFIC"'
                   +',"RTCONTMAILINGFIC","RTMULACTIONSCH","RTMULACTIONSPRO","RTMULCONTACTIONS","RTMULCONTMAILING"'
                   +',"RTMULQUALITECONTA")');
//GIGA
Vers_Activite := TableToVersion('ACTIVITE');
End;

Procedure MajAvant640;
Begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCSTKPHYSIQUE","GCSTKPREVISION","GCSTKPROJETES"'
+',"WORDRETET","WORDRELIG","WORDRELIG2","WORDREPHASE","WORDREPHASE2","WORDREBES","WORDREBES2","WORDREGAMME","WORDREINTERTET"'
+',"WORDREINTERLIG","WGAMMETET","WGAMMELIG","WNOMETET","WNOMELIG","WNOMEDEC","WREVISIONWNT","WREVISIONWGT")');
End;

Procedure MajAvant642;
Begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCSTKBILQUALIFMVT","GCSTKDISPODETAIL"'
     +',"GCSTKPHYSIQUE","GCSTKPREVISION","WORDREINTERLIG","WPARC"'
     +',"UORDRELIG","UORDREBES","AFREVREGULFACTURE","AFREVAPPLIQCOEF")');
TestDropTable('STKMOUVEMENT');
TestDropTable('STKNATURE');
TestDropTable('DISPODETAIL');
End;

Procedure MajAvant644;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCSTKDISPODETAIL","GCSTKPHYSIQUE"'
            +',"GCSTKPREVISION","WORDREPHASE","WORDREPHASE2"'
            +',"PGMULVARIABLE"'
            +',"YTARIFSAQUI","YTARIFSAQUI101","YTARIFSAQUI102","YTARIFSAQUI103"'
            +',"YTARIFSAQUI201","YTARIFSAQUI202","YTARIFSAQUI203","YTARIFSAQUI301"'
            +',"YTARIFSAQUI401","YTARIFSURQUOI","YTARIFSURQUOI101","YTARIFSURQUOI102"'
            +',"YTARIFSURQUOI103","YTARIFSURQUOI201","YTARIFSURQUOI202","YTARIFSURQUOI203"'
            +',"YTARIFSURQUOI301","YTARIFSURQUOI401","YTARIFSTRANCHE","YTARIFSTRANCHE101"'
            +',"YTARIFSTRANCHE102","YTARIFSTRANCHE103","YTARIFSTRANCHE201","YTARIFSTRANCHE202"'
            +',"YTARIFSTRANCHE203","YTARIFSTRANCHE301","YTARIFSTRANCHE401"'
            +')');
  if TableExiste('INSTITUTIONPAYE') then ExecuteSQL('DELETE FROM INSTITUTIONPAYE');
  GIGAAvant644;
  //GPAO
  TestDropTable('YTARIFSFONCTIONNALITE');
  TestDropTable('YTARIFSAQUI');
  TestDropTable('YTARIFSSURQUOI');
  TestDropTable('LIGNETARIF');
  ExecuteSQL('DELETE FROM CHOIXEXT WHERE YX_TYPE IN ("YTL","YTJ","YTS","YTD")');
  ExecuteSQL('DELETE FROM CHOIXCOD WHERE CC_TYPE IN ("YTR","YTN","YTU")');
End;

Procedure MajAvant645;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE in ("GCLISTINVLIGNE"'
            +',"PGMULABSIJSS"'
            +')');
End;

Procedure MajAvant646;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE in ("TRLISTETRANSAC", "TRPREVISIONNELLES","PGMULABSIJSS")');
End;


Procedure MajAvant648;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCSTKDISPO","WORDRELIG3"'
                     +',"YTARIFSMULART101","YTARIFSMULART201","YTARIFSMULART211"'
                     +',"YTARIFSMULART301","YTARIFSMULART401","YTARIFSMULTIE101"'
                     +',"YTARIFSMULTIE201","YTARIFSMULTIE211","YTARIFSMULTIE301"'
                     +',"YTARIFSMULTIE401","YTARIFSBASCULE")');
End;

Procedure MajAvant649;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("PGRHCOMPETRESSOUR")');
End;


Procedure MajAvant650;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WOL_EPR","WOL_RECEPTION","WOB_SPR"'
      +',"WOB_HISTORUP","WOB_CONSOMMATION"'
      +',"GCSTKDISPODETAIL","GCSTKPHYSIQUE","GCSTKPREVISION","GCSTKPROJCON"'
      +',"GCSTKPROJETES","WCBNPREVTET","WCBNPREVLIG"'
      +',"GCSTKVALOPARAM","GCSTKVALOTYPE"'
      +',"WORDRECMP","WORDRECMP_VUE","WORDRELIG","WPDRTET"'
      +',"WNOMELIG","WORDREBES","WOB_RUPTURES","WORDREBES2"'
      +',"WORDREPHASE","WORDREPHASE2","YTARIFSMULTIE101","YTARIFSPARAMETRES"'
      +',"WORDREINTERLIG")');
  if TableExiste('WNOMEDEC') then ExecuteSQL('DELETE FROM WNOMEDEC');
  //GIGA
  if TableExiste('AFORMULEVAR') then
  begin
    if not (ExisteSQL('SELECT AVF_FORMULEVAR FROM AFORMULEVAR')) then TestDropTable('AFORMULEVAR');
    if not (ExisteSQL('SELECT AVD_FORMULEVAR FROM AFORMULEVARDET')) then  TestDropTable('AFORMULEVARDET');
    if not (ExisteSQL('SELECT AVV_ORIGVAR FROM AFORMULEVARQTE')) then  TestDropTable('AFORMULEVARQTE');
  end;
End;


Procedure MajAvant651;
Begin

  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCSTK PROJETES")');
  //GIGA
  Vers_Piece := TableToVersion('PIECE');
  //COMPTA
  ExecuteSQL('delete from modedata where md_cle like "LRLCRCT%" ');
  ExecuteSQL('delete from modeles where mo_type="L" and mo_nature="RLC" and mo_code="RCT" ');
  MoveCommunToChoixcode('AFP', 'AFP');
  if V_PGI.ModePCL='1' then
    ExecuteSQL('update menu set mn_accesgrp="0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" where mn_tag=7602 and mn_1=16');
  //TestDropTable('PSUIVI');       TestDropTable('PCURRICULUM');    TestDropTable('PCANDIDAT');
  //TestDropTable('PANNONCE');     TestDropTable('PCANDIP');        TestDropTable('PCANLANGUES');
  //TestDropTable('PDROITS');      TestDropTable('PRMETIER');       TestDropTable('PSUPPORTANNONCE');
  //TestDropTable('PSUIVITCANDIDAT');

End;

Procedure MajAvant653;
Begin
  if V_PGI.ModePCL='1' then
    ExecuteSQL('update menu set mn_ACCESGRP=(select MAX(mn_accesgrp) from menu MN where MN.mn_tag=Menu.mn_tag and MN.mn_1 in (141,142,143,144,151,152,154)) '
        +' where (mn_1 not in (141,142,143,144,151,152,154)) and mn_tag<>0 '
        +' and (mn_tag  in (select mn_tag from menu M1 where M1.mn_1 in (141,142,143,144,151,152,154) AND M1.mn_tag=Menu.mn_tag ))');

  ExecuteSQL('update article set ga_typenomenc="ASC" where ga_typearticle="NOM" and ga_typenomenc="ASS" and ga_tenuestock="-"');
  //GP
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCSAISIECSP")');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE LIKE "YTARIFSMULTIE%"')
End;


Procedure MajAvant654;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WCC", "WOB_SPR", "WORDRECMP", "WRP", "GCSTKDISPODETAIL", "GCSTKGQDDATE", "GCSTKPHYSIQUE"'
            + ',"WORDREBES3", "WORDRELIG3", "WORDREPHASE3", "WQUANTITE", "WTEMPS"'
            + ',"YTARIFSFSLART211", "YTARIFSFSLTIE211", "YTARIFSMULART211", "YTARIFSMULTIE211", "YTARIFSPARAMETRES"'
            + ',"TBLISTECOMPTA","GCSTKBILQUALIFMVT","WORDREBES2","WORDREBESL","WORDRELIG2","WORDRECMP"'
            + ',"EDIMESSAGES","TRLISTETRANSAC" )');
  // TRESO
  ExecuteSQL('DELETE FROM MODELES WHERE MO_TYPE = "E" AND MO_NATURE = "ECT" AND MO_CODE IN ("ORV", "EIS") ');
  ExecuteSQL('DELETE FROM MODEDATA WHERE MD_CLE LIKE "EECTORV%"');
  ExecuteSQL('DELETE FROM MODEDATA WHERE MD_CLE LIKE "EECTEIS%"');

  ExecuteSQL('DELETE FROM MODELES WHERE MO_TYPE = "T" AND MO_NATURE = "ECT" AND MO_CODE IN ("TOP", "ORD")');
  ExecuteSQL('DELETE FROM MODEDATA WHERE MD_CLE LIKE "TECTTOP%"');
  ExecuteSQL('DELETE FROM MODEDATA WHERE MD_CLE LIKE "TECTORD%"');

  ExecuteSQL('DELETE FROM MODELES WHERE MO_TYPE = "L" AND MO_NATURE = "ECT" AND MO_CODE IN ("TOP", "ORP")');
  ExecuteSQL('DELETE FROM MODEDATA WHERE MD_CLE LIKE "LECTTOP%"');
  ExecuteSQL('DELETE FROM MODEDATA WHERE MD_CLE LIKE "LECTORP%"');
  (* Pour la V6, ce paramsoc doit être de nouveau décoché *)
  ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA = "-" WHERE SOC_NOM = "SO_PREMIERESYNCHRO"');
End;

Procedure MajAvant655;
Begin
   ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCSTKDISPODETAIL","GCSTKBILQUALIFMVT", "GCSTKBILMVT" '
          + ',"TRLISTECIB", "TRLISTECIBACCRO")');
End;


Procedure MajAvant656;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCSTKPHYSIQUE","YTARIFSBASCULE","TRBANQUECP")');
End;

Procedure MajAvant657 ;
var stModules : string;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WORDREINTERTET", "WORDREINTERLIG"'
            +' ,"WORDRETET","WNOMELIG","WCBNPROPOSITION","WCBNPREVTET","WCBNPREVLIG"'
            +',"GCMULARCHIVAGEMVT","GCSTKDISPATCH","GCSTKDISPO","GCSTKDISPODETAIL")');

  stModules:='(30,36,31,32,33,70,92,93,59,160,111,122,260,125,65,60)';
  if not (V_PGI.ModePCL='1') then
     ExecuteSQL('update menu set mn_ACCESGRP=(select MAX(mn_accesgrp) from menu MN '
                                    +' where MN.mn_tag=Menu.mn_tag and MN.mn_1 in '+stModules+') '
                + ' ,MN_VERSIONDEV=(select MAX(MN_VERSIONDEV) from menu MN '
                                    +' where MN.mn_tag=Menu.mn_tag and MN.mn_1 in '+stModules+') '
                +' where (mn_1 not in '+stModules+') and mn_tag<>0 '
                +' and (mn_tag  in (select mn_tag from menu M1 where M1.mn_1 in '+stModules+' AND M1.mn_tag=Menu.mn_tag ))');

End;

Procedure MajAvant658 ;
Begin
  //GIGA
  if V_PGI.ModePCL='1' then
    GIGAAvant658;
  // GPAO
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WCBNPROPOSITION","WNOMEDEC")');
  // Maj pour les traifs GA
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("AFTARIFCA","AFTARIFCT","AFTARIFPRIX","AFTARIFQTEPRIX","AFTARIFREM")');
  // Maj pour les traifs GC
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCTARIFCA","GCTARIFCT","GCTARIFCTPRIX","GCTARIFCTQTEPRIX"'
      + ',"GCTARIFPRIX","GCTARIFQTEPRIX","GCTARIFREM","GCTARIFSR","GCTARIFCON" )');
End;

Procedure MajAvant659 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WORDREINTERLIG","GCSTKDISPATCH"'
       +',"GCSTKGQDDATE","GCSTKPHYSIQUE","GCTARIFSR","WGAMMECIR"'
       +',"AFTARIFCON","AFTARIFCTPRIX","AFTARIFCTQTEPRIX","AFTARIFQTEPRIX"'
       +',"AFTARIFFOUCON","GCTARIFFOUCON","GCTARIFFOUMAJ","GCTARIFMAJ"'
       +')');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE LIKE "YTARIFSMUL%"');
  if TableExiste('WGAMMECIR') then ExecuteSQL('DELETE FROM WGAMMECIR');  // on vide cette table temporaire avant modif d'index
End;

Procedure MajAvant660 ;
Begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WCBNPREVTET","WCBNPREVLIG","GCSTKDISPO","GCSTKDISPODETAIL","WCBNBESOIN")');
End;

Procedure MajAvant661 ;
Begin
  if TableExiste('JQ6_fonction') then ExecuteSQL('delete from JQ6_fonction');
	if TableExiste('jq1_formejur') then ExecuteSQL('delete from jq1_formejur');
	if TableExiste('jq2_typeper') then ExecuteSQL('delete from jq2_typeper');
	if TableExiste('jq3_typecivil') then ExecuteSQL('delete from jq3_typecivil');
	if TableExiste('jq4_situfam') then ExecuteSQL('delete from jq4_situfam');
	if TableExiste('jq5_activite') then ExecuteSQL('delete from jq5_activite');
	if TableExiste('jq7_famper') then ExecuteSQL('delete from jq7_famper');
	if TableExiste('jq8_regfiscal') then ExecuteSQL('delete from jq8_regfiscal');
	if TableExiste('jqf_fields') then ExecuteSQL('delete from jqf_fields');
	if TableExiste('jqt_tables') then ExecuteSQL('delete from jqt_tables');


  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE in ("GCTARIFSR","TRLISTETRANSAC","CPCONSGENE")');

  ExecuteSQL('DELETE FROM MODELES WHERE MO_TYPE = "E" AND MO_NATURE = "ECT" AND MO_CODE IN ("TOP", "ORV", "EIS") AND MO_LANGUE= "FRA"');
  ExecuteSQL('DELETE FROM MODELES WHERE MO_TYPE = "T" AND MO_NATURE = "ECT" AND MO_CODE IN ("TOP", "ORP") AND MO_LANGUE= "FRA"');
  ExecuteSQL('DELETE FROM MODELES WHERE MO_TYPE = "L" AND MO_NATURE = "ECT" AND MO_CODE IN ("TOP", "ORD") AND MO_LANGUE= "FRA"');

End;


Procedure MajAvant662 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE in ("GCSTKDISPATCH","GCSTKGQDDATE"'
                     +',"GCSTKPHYSIQUE","QULSITE","AFMULFRAISCOMPTA"'
                     +',"TRRUBRIQUETRESO","CPCONSGENE","TRECRITURE", "TRLISTECIBACCRO"'
                     +')');
End;

Procedure MajAvant663 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WORDRECMP","WGAMMELIG"'
      + ',"GCSTKPHYSIQUE","GCSTKPREVISION","WNOMELIG","GCSTKVALOTYPE"'
      + ',"WORDREGAMME","WGAMMETET","WCBNPROPOSITION","UORDREBESLASSTOC"'
      + ',"UORDREBESLASSAIS","WORDREBESL","WORDRERES","WORDREINTERTET","PRT_TRAGENCE"'
      + ',"YTARIFSAQUI","WCBNPROPOSITION","WORDREBES2CPD","WORDRELIG2"'
	    + ',"WORDREPHASE2","WORDREBES","WORDRECMP","YTARIFSFOURCH211","YTARIFSFSLART211","YTARIFSFSLTIE211"'
      + ',"GCLISTINVLIGNE","WORDREBESL","GCTRACABILITELIG","CPAFTIECPTA"'
      + ')');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE LIKE "YTARIFSURQUOI%"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE LIKE "YTARIFSTRANCHE%"');

  //Compta
  ExecuteSQL('delete from modedata where md_cle like "EJACJCE%"');
  ExecuteSQL('delete from modeles where mo_type="E" and mo_nature="JAC" and mo_code="JCE"');
  if not ChampPhysiqueExiste('BANQUES', 'PQ_ETABBQ') then
    ExecuteSQL('DELETE FROM DECHAMPS where DH_NOMCHAMP="PQ_ETABBQ"');
  if not ChampPhysiqueExiste('BANQUES', 'PQ_ABREGE') then
    ExecuteSQL('DELETE FROM DECHAMPS where DH_NOMCHAMP="PQ_ABREGE"');
  // correction des mn_versiondev à F ou T
  ExecuteSQL('update menu set mn_versiondev="X" where mn_versiondev="T"');
  ExecuteSQL('update menu set mn_versiondev="-" where mn_versiondev="F"');
  //TRESO
  ExecuteSQL('update menu set mn_versiondev="-" where mn_1=135 and mn_2=1');
  //DP
  ExecuteSQL('update menu set mn_versiondev="-" where mn_1 in (172,173,174,175,176,187) and mn_versiondev="X" ');
End;

Procedure MajAvant667 ;
Begin
	//GIGA .. ces listes ne sont plus utilisées..
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("AFTARIFFOUMAJ","AFTARIFMAJ","AFTACHELIG"'
       +',"AFTACHE_SS_RESS","WORDRECMP_VUE","WRECALCULQTERESTE"'
       +',"TRCONDITIONFIN", "TRLISTEFLUXTRESO", "TRLISTEVIR"'
       +',"TRLISTETRANSAC", "TRECRITURERAPPRO", "TRECRITURE", "PRT_TRAGENCE", "TRPREVISIONNELLES"'
       +')');
End;

Procedure MajAvant668 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("UORDRELASST")');
End;

Procedure MajAvant669 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WPARC","WORDREINTERLIG","TRLISTECIB")');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE LIKE "YTARIFSMUL%"');
End;

Procedure MajAvant670 ;
Begin
  //Compta
  MoveCommunToChoixcode('QME', 'QME');  // TTQUALUNITMESURE
  //GP
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCSTKDISPO","GCSTKMANQUE","GCSTKPREVISION")');
End;


Procedure MajAvant671 ;
Begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "WORDRECMP"');
End;


Procedure MajAvant673 ;
Begin
  ExecuteSQL('UPDATE PARPIECE SET GPP_NATURESUIVANTE=GPP_NATURESUIVANTE||"PRF;" WHERE GPP_NATUREPIECEG="CF" AND GPP_NATURESUIVANTE not like "%PRF%"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WLIGNESCSP","GCTRANSINVLIGNE","WGAMMELIG")');
End;

Procedure MajAvant674 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCMULMODIFLOT","GCMULLIGNE")');
End;

Procedure MajAvant675 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCTRANSINVENT","WGAMMELIG","WGAOPER","WGAOPERRES","WPDRTET")');
End;

Procedure MajAvant676 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WNOMEDEC","WNOMECAS","WORDREINTERTET","TRECRITURERAPPRO")');
  //TRESO
  if TableExiste('FLUXTRESO') and (tableToVersion('FLUXTRESO') < 7) then
  begin
    AddChamp('FLUXTRESO', 'TFT_CLASSEFLUX', 'COMBO', 'TFT', False, '', '', 		'', False);
    if ChampPhysiqueExiste('FLUXTRESO', 'TFT_CLASSEFLUX') then
    begin
      if ChampPhysiqueExiste('FLUXTRESO', 'TFT_PREVISIONNEL') then
      begin
        ExecuteSQL('UPDATE FLUXTRESO SET TFT_CLASSEFLUX = "PRE" WHERE 		  		TFT_PREVISIONNEL = "X"');
        ExecuteSQL('UPDATE FLUXTRESO SET TFT_CLASSEFLUX = "FIN" WHERE 			TFT_PREVISIONNEL = "-" OR TFT_PREVISIONNEL = "" OR 				TFT_PREVISIONNEL IS NULL');
      end;
	  end;
  end;
End;

Procedure MajAvant680 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCMULARTINFOS","WCOMPAREWGL")');
End;

Procedure MajAvant681 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCSTKMANQUE","GCSTKPREVISION","TRCOTATIONTAUX")');
End;

Procedure MajAvant682 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("UORDRELIG","UORDREBES","UORDREBESLASSAIS","UORDREBESLASSTOC")');
End;

Procedure MajAvant683 ;
Begin
  // GRC
  SupprimeEtat('E' ,'RPS','RP1', True);
  // GPAO
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="WCOMPAREWPL"');

End;


Procedure MajAvant684 ;
Begin
  // PCL pour réparer les dossiers mal réparés par l'utilitaire de marc
  if (not IsMonoOuCommune) and (V_PGI.Driver = dbMSSQL) and ( Not ExisteSQL('select 1 from sysindexes where name="DH_CLE1"') ) then //
  begin
    try ExecuteSQL('create unique clustered index DH_CLE1 on DECHAMPS(DH_PREFIXE, DH_NUMCHAMP)'); except; end;
    try ExecuteSQL('create unique index DH_CLE2 on DECHAMPS(DH_NOMCHAMP)'); except; end;
    try ExecuteSQL('create unique clustered index DT_CLE1 on DETABLES(DT_NOMTABLE)'); except; end;
    try ExecuteSQL('create unique clustered index DL_CLE1 on DELIENS(DL_TABLE1, DL_TABLE2)'); except; end;
    try ExecuteSQL('create unique clustered index DV_CLE1 on DEVUES(DV_NOMVUE)'); except; end;
    try ExecuteSQL('create unique clustered index DV_CLE1 on DECOMBOS(DO_COMBO)'); except; end;
    try ExecuteSQL('create unique clustered index DFM_CLE1 on FORMES(DFM_TYPEFORME, DFM_FORME)'); except; end;
    try ExecuteSQL('create unique clustered index PO_CLE1 on PROCEDUR(PO_PROCEDURE)'); except; end;
    try ExecuteSQL('create unique clustered index SRC_CLE1 on SCRIPTS(SRC_TYPESCRIPT, SRC_SCRIPT)'); except; end;
  end;

  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("YYMESSAGES_IN","YYMESSAGES_INGLOB","YYMESSAGES_OUT","YYMESSAGESDOS_MUL","TBLISTEPAIE")');

End;



Procedure MajAvant685 ;
Begin
   ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WORDREAUTO","WORDREPHASE2","YYMULJNALEVENT","TRLISTEFLUXTRESO")');
End;


Procedure MajAvant686 ;
Begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE in ("WTRAITEMENT")');
End;

Procedure MajAvant687 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCUNITELOGISTIQUE","GCSTKPICKING","CPCONSGENE","CPEXPORT")');
End;

Procedure MajAvant688 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCLISTEINVLIGNE","GCSTKDISPATCH","GCSTKDISPODETAIL","GCSTKPICKING","CPCONSGENE")');
End;

Procedure MajAvant689 ;
Begin
   ExecuteSql('DELETE FROM LISTE WHERE LI_LISTE="WORDRELIG2"');
End;

Procedure MajAvant690 ;
var
  sql: string;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("AFMULAFFAIREEXPO1","CPEXPORT"'
                     +',"WSSORDREPARAM","WORDREAUTO","WPF_AUTOWOL","WORDRECMP")');
  //COMPTA
  SupprimeEtat('E' ,'GFF','QUE', True);
  // GPAO

  if TableExiste('WORDREAUTO') then
  begin
  { Mémorisation des infos de WORDREAUTO et WDEPOTAPP avant destruction de la table }
  TobWPF := Tob.Create('WPF', nil, -1);
  Sql := ' SELECT'
       + ' "AUTOWOL" AS WPF_CODEFONCTION, PGIGUID AS WPF_GUID, "" AS WPF_ETABLISSEMENT, WOA_DEPOT AS WPF_DEPOT,'
       + ' "' + UsDateTime(iDate1900) + '" AS WPF_DATECREATION, "' + UsDateTime(iDate1900) + '" AS WPF_DATEMODIF, WOA_CREATEUR AS WPF_CREATEUR, WOA_UTILISATEUR AS WPF_UTILISATEUR,'
       + ' WOA_FAMILLENIV1 AS WPF_FAMILLENIV1, WOA_FAMILLENIV2 AS WPF_FAMILLENIV2, WOA_FAMILLENIV3 AS WPF_FAMILLENIV3, "" AS WPF_COLLECTION, "" AS WPF_FAMILLEVALO, WOA_ARTICLE AS WPF_ARTICLE, WOA_CODEARTICLE AS WPF_CODEARTICLE,'
       + ' WOA_NATURETRAVAIL AS WPF_NATURETRAVAIL, WOA_TYPEORDRE AS WPF_TYPEORDRE,'
       + ' "" AS WPF_QUALIFMVT, "" AS WPF_STKFLUX,'
       + ' WOA_BOOLLIBRE1 AS WPF_BOOLEAN01, WOA_BOOLLIBRE2 AS WPF_BOOLEAN02, WOA_BOOLLIBRE3 AS WPF_BOOLEAN03, WOA_BOOLLIBRE4 AS WPF_BOOLEAN04, WOA_BOOLLIBRE5 AS WPF_BOOLEAN05,'
       + ' WOA_BOOLLIBRE6 AS WPF_BOOLEAN06, WOA_BOOLLIBRE7 AS WPF_BOOLEAN07, WOA_BOOLLIBRE8 AS WPF_BOOLEAN08, "-" AS WPF_BOOLEAN09, "-" AS WPF_BOOLEAN10,'
       + ' 0 AS WPF_INTEGER01, 0 AS WPF_INTEGER02, 0 AS WPF_INTEGER03, 0 AS WPF_INTEGER04, 0 AS WPF_INTEGER05, 0 AS WPF_INTEGER06, 0 AS WPF_INTEGER07, 0 AS WPF_INTEGER08, 0 AS WPF_INTEGER09, 0 AS WPF_INTEGER10,'
       + ' 0 AS WPF_DOUBLE01, 0 AS WPF_DOUBLE02, 0 AS WPF_DOUBLE03, 0 AS WPF_DOUBLE04, 0 AS WPF_DOUBLE05, 0 AS WPF_DOUBLE06, 0 AS WPF_DOUBLE07, 0 AS WPF_DOUBLE08, 0 AS WPF_DOUBLE09, 0 AS WPF_DOUBLE10,'
       + ' "' + UsDateTime(iDate1900) + '" AS WPF_DATE01, "' + UsDateTime(iDate1900) + '" AS WPF_DATE02, "' + UsDateTime(iDate1900) + '" AS WPF_DATE03, "' + UsDateTime(iDate1900) + '" AS WPF_DATE04, "' + UsDateTime(iDate1900) + '" AS WPF_DATE05, "' + UsDateTime(iDate1900) + '" AS WPF_DATE06, "' + UsDateTime(iDate1900) + '" AS WPF_DATE07, "' + UsDateTime(iDate1900) + '" AS WPF_DATE08, "' + UsDateTime(iDate1900) + '" AS WPF_DATE09, "' + UsDateTime(iDate1900) + '" AS WPF_DATE10,'
       + ' "" AS WPF_COMBO01,  "" AS WPF_COMBO02, "" AS WPF_COMBO03, "" AS WPF_COMBO04, "" AS WPF_COMBO05, "" AS WPF_COMBO06, "" AS WPF_COMBO07, "" AS WPF_COMBO08, "" AS WPF_COMBO09, "" AS WPF_COMBO10,'
       + ' "" AS WPF_VARCHAR01,  "" AS WPF_VARCHAR02, "" AS WPF_VARCHAR03, "" AS WPF_VARCHAR04, "" AS WPF_VARCHAR05, "" AS WPF_VARCHAR06, "" AS WPF_VARCHAR07, "" AS WPF_VARCHAR08, "" AS WPF_VARCHAR09, "" AS WPF_VARCHAR10,'
       + ' "-" AS WPF_WBMEMO'
       + ' FROM WORDREAUTO'
       ;
  TobWPF.LoadDetailDBFromSql('WPF', Sql);

  Sql := 'SELECT'
       + ' "DEPOTAPP" AS WPF_CODEFONCTION, PGIGUID AS WPF_GUID, "" AS WPF_ETABLISSEMENT, WEA_DEPOTENC AS WPF_DEPOT,'
       + ' "' + UsDateTime(iDate1900) + '" AS WPF_DATECREATION, "' + UsDateTime(iDate1900) + '" AS WPF_DATEMODIF, "" AS WPF_CREATEUR, "" AS WPF_UTILISATEUR,'
       + ' WEA_FAMILLENIV1 AS WPF_FAMILLENIV1, WEA_FAMILLENIV2 AS WPF_FAMILLENIV2, WEA_FAMILLENIV3 AS WPF_FAMILLENIV3, "" AS WPF_COLLECTION, "" AS WPF_FAMILLEVALO, WEA_ARTICLE AS WPF_ARTICLE, WEA_CODEARTICLE AS WPF_CODEARTICLE,'
       + ' "" AS WPF_NATURETRAVAIL, "" AS WPF_TYPEORDRE,'
       + ' "" AS WPF_QUALIFMVT, "" AS WPF_STKFLUX,'
       + ' "-" AS WPF_BOOLEAN01, "-" AS WPF_BOOLEAN02, "-" AS WPF_BOOLEAN03, "-" AS WPF_BOOLEAN04, "-" AS WPF_BOOLEAN05, "-" AS WPF_BOOLEAN06, "-" AS WPF_BOOLEAN07, "-" AS WPF_BOOLEAN08, "-" AS WPF_BOOLEAN09, "-" AS WPF_BOOLEAN10,'
       + ' 0 AS WPF_INTEGER01, 0 AS WPF_INTEGER02, 0 AS WPF_INTEGER03, 0 AS WPF_INTEGER04, 0 AS WPF_INTEGER05, 0 AS WPF_INTEGER06, 0 AS WPF_INTEGER07, 0 AS WPF_INTEGER08, 0 AS WPF_INTEGER09, 0 AS WPF_INTEGER10,'
       + ' 0 AS WPF_DOUBLE01, 0 AS WPF_DOUBLE02, 0 AS WPF_DOUBLE03, 0 AS WPF_DOUBLE04, 0 AS WPF_DOUBLE05, 0 AS WPF_DOUBLE06, 0 AS WPF_DOUBLE07, 0 AS WPF_DOUBLE08, 0 AS WPF_DOUBLE09, 0 AS WPF_DOUBLE10,'
       + ' "' + UsDateTime(iDate1900) + '" AS WPF_DATE01, "' + UsDateTime(iDate1900) + '" AS WPF_DATE02, "' + UsDateTime(iDate1900) + '" AS WPF_DATE03, "' + UsDateTime(iDate1900) + '" AS WPF_DATE04, "' + UsDateTime(iDate1900) + '" AS WPF_DATE05, "' + UsDateTime(iDate1900) + '" AS WPF_DATE06, "' + UsDateTime(iDate1900) + '" AS WPF_DATE07, "' + UsDateTime(iDate1900) + '" AS WPF_DATE08, "' + UsDateTime(iDate1900) + '" AS WPF_DATE09, "' + UsDateTime(iDate1900) + '" AS WPF_DATE10,'
       + ' WEA_DEPOTAPP AS WPF_COMBO01,  "" AS WPF_COMBO02, "" AS WPF_COMBO03, "" AS WPF_COMBO04, "" AS WPF_COMBO05, "" AS WPF_COMBO06, "" AS WPF_COMBO07, "" AS WPF_COMBO08, "" AS WPF_COMBO09, "" AS WPF_COMBO10,'
       + ' "" AS WPF_VARCHAR01,  "" AS WPF_VARCHAR02, "" AS WPF_VARCHAR03, "" AS WPF_VARCHAR04, "" AS WPF_VARCHAR05, "" AS WPF_VARCHAR06, "" AS WPF_VARCHAR07, "" AS WPF_VARCHAR08, "" AS WPF_VARCHAR09, "" AS WPF_VARCHAR10,'
       + ' "-" AS WPF_WBMEMO'
       + ' FROM WDEPOTAPP'
       ;
  TobWPF.LoadDetailDBFromSql('WPF', Sql, True);
  end;
End;

Procedure MajAvant691 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WGAMMECIR","WCOMPAREWGLWOG"'
             +',"WPDRTET","GCSTKVALOTYPE","CPMULCUTOFF","CPCONSAUXI")');
End;

Procedure MajAvant692 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WSOUSORDREVISU","WPF_EXTFORMS","CPCONSGENE","CPCONSAUXI")');
    // IMMO
  TestDropTable('IMOIMP');   // problème de version. (Table vide)
End;

Procedure MajAvant694 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "WGAMMECIR"');
End;

Procedure MajAvant695 ;
Begin
  TestDropTable('STKNATURE');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WORDRELIG2","REIMPUTANAL")');
  //PAIE    ajout 07092005
  // ExecuteSql('UPDATE histoanalpaie SET PHA_COTREGUL= "..."');
  // ExecuteSql('UPDATE histobulletin SET PHB_COTREGUL= "..."');
  if TableExiste('HISTOANALPAIE') then AddChamp('HISTOANALPAIE', 'PHA_COTREGUL', 'COMBO', 'PHA', False, '...', 'Cotisation de régularisation', 'LDC', False);
  if TableExiste('HISTOBULLETIN') then AddChamp('HISTOBULLETIN', 'PHB_COTREGUL', 'COMBO', 'PHB', False, '...', 'Régularisation', 'LDC', False);
End;



Procedure MajAvant696 ;
Begin
   ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WGAMMECIR","WPF_DEPOTAPP","UCOMPFOURNI")');
End;

Procedure MajAvant697 ;
Begin
ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("QULBPARBRE","QULBPARBRECA1","QULBPARBRECA2"'
                   +',"QULBPARBRECA3","QULBPARBRECA4","QULBPARBRECA5","QULBPARBRECA6"'
                   +',"QULBPARBRED","QULBPARBREQTE","QULBPCHOIXCOD","QULBPCHOIXEXT","QULBPLOIPGI002"'
                   +',"QULBPSESSIONPGI07","GCSTKDEPLACEMENT","GCSTKDISPATCH","GCSTKDISPODETAIL"'
                   +',"GCSTKPICKING","WPF_DELDISPO","WPF_DUREEVIE","WPF_RGESTION"'
                   +',"AFMULPLANNING","AFTACHE_MUL","AFFTIERS","AFMULLIGNERECH","WRESSOURCETV")');

End;

Procedure MajAvant698 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("TRECRITURE","GCSTKDISPATCH","CPMULCUTOFF"'
      +',"GCSTKFICHELOT","GCSTKFICHESERIE","GCSTKNOMEGST","GCSTKNOMELOT","GCSTKNOMESERIE","QULBPSESSIONBPENT")');
End;

Procedure MajAvant699 ;
Begin
  if (TableToVersion('PARCAISSE') = 116) and (ChampPhysiqueExiste('PARCAISSE', 'GPK_FAMRES')) and (NOT ChampPhysiqueExiste('PARCAISSE', 'GPK_LFAMRES')) then
  begin
    AddChamp('PARCAISSE', 'GPK_LFAMRES', 'VARCHAR(200)', 'GPK', False, '', 'Liste des familles de ressources', 'LDZ', False);
    ExecuteSQL ('UPDATE PARCAISSE SET GPK_LFAMRES=GPK_FAMRES WHERE GPK_FAMRES IS NOT NULL')
  end;

  // Pour corriger un vieux bug de la fisca perso
  if tableExiste( 'FPBIENS_VE') and ChampLogiqueExiste('FPBIENS_VE', 'FVE_N03705') and (not ChampPhysiqueExiste('FPBIENS_VE', 'FVE_N03705')) then
      ExecuteSQL ('DELETE From DECHAMPS Where DH_NOMCHAMP="FVE_N03705"');
  if tableExiste( 'FPBIENS') and ChampLogiqueExiste('FPBIENS', 'FPB_N01281') and (not ChampPhysiqueExiste('FPBIENS', 'FPB_N01281')) then
    ExecuteSQL ('DELETE From DECHAMPS Where DH_NOMCHAMP="FPB_N01281"');
   //GIGA
   ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE = "AFREVREGULFACTURE"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("YTARIFSFOURCH501","YTARIFSFSLART501"'
      +',"YTARIFSFSLTIE501","YTARIFSPARAMETRES","WUNITES","AFREVREGULFACTURE"'
      +',"QULBPARBREQTE","QULBPARBRECA1","QULBPARBRECA2","QULBPARBRECA3"'
      +',"QULBPARBRECA4","QULBPARBRECA5","QULBPARBRECA6","GCSTKPREVISION","GCSTKVALOTYPE","GCSTKFICHELOT")');

    TestDropTable('BILANSOCIAL');
    TestDropTable('PBSINDDETSEL');
    TestDropTable('PBSINDSELECT');
    TestDropTable('PBSINDICATEURS');
End;


Procedure MajAvant700 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("YTARIFSMULTIE501","YTARIFSMULART501"'
        +',"YTARIFSFSLART501","YTARIFSFSLART601","YTARIFSFSLTIE501","YTARIFSFSLTIE601"'
        +',"YTARIFSFOURCH501","YTARIFSFOURCH601","CPEXPORTVCOM","WPARC","WVERSION")');
  //SAV
  if TableToVersion('WPARC') < 3 then
    TestDropTable('WPARC');
  if TableToVersion('WPARCNOME') < 9 then
    TestDropTable('WPARCNOME');
  if TableToVersion('WVERSION') < 5 then
    TestDropTable('WVERSION');
  //TRESO
  SupprimeEtat('E' ,'GFF','QUE', True);
  SupprimeEtat('E' ,'GFF','QUI', True);

  SupprimeEtat('F' ,'FIC','GE1', True);
  SupprimeEtat('F' ,'FIC','GE2', True);
  SupprimeEtat('F' ,'FIC','GE3', True);
End;

Procedure MajAvant701 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCPIEDPORT","GCSTKNOMELOT"'
                     +',"GCSTKNOMESERIE","AFMULARTICLE","AFMULRECHARTICLE","AFMULMODIFLOT")');
End;

Procedure MajAvant702 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WPF_PREPAEXPEXT", "WPF_PREPAEXPGEN"'
     +', "GCPREPAEXPEXT","GCPREPAEXPEXT", "GCPREPAEXPAFF", "GCPREPAEXPAFC", "GCPREPAEXPE" '
     +')');
End;


Procedure MajAvant703 ;
Begin
  //-- SUPPRESSION DES LISTES  --//
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN'
  +'("QULGRPOFCAR","QULGRPCOMP","QULHISTOAFF","QULHISTOAFF2","QULDETPROF","QULDETOFD","QULETATPLAN","QULDETSCENA",'
  +'"QULPHASEITI","QULPHASECAR","QULPLANGP","QULPLANGP1","QULMACROGAM","QULITI","QULNOMENCLA","QULMENUREGRP003",'
  +'"QULARTBTNPRO","QULAPPROSSTKCOMP","QULARTTECH","QULARTCHCAR","QULAFFECTOF001","QULAFFECTOF",'
  +'"QULAPPROSSTKART","QULAPPROSCOMP","QULCHGTGAM","QULCARCOUL2","QULDETCIRC","QULCOMPOST",'
  +'"QULBTNCHOIXPROD","QULARTTECH2","QULCARCOUL","QULBTNCHXPRODDET","QULPROINDUS","QULMENUREGRP004",'
  +'"QULGROUPAGECAR","QULSELECTLANCEMEN","QULSATTRIB22","QULSATTRIB2","QULTRANDETDDELPHI","QULSTKDISPO","QULRECALAGEPLAN",'
  +'"QULRECALAGPLAN001","QULREPARTFAB","QULVUEREGRPTR004","QULVUEREGRPTR","GCSTKFORMULE","GCSTKVALOTYPE"'
  +',"PGBSINDSELECT","PGBILANSOCIAL","PGINDICATEURBS")');
End;


Procedure MajAvant704 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCPREPAEXPEXT", "GCPREPAEXPAFF"'
                     +',"GCSTKDISPATCH","WGAMMELIG","WORDREBESCPD","WCOMPAREWGLWOG","WCOMPAREWNLWOB"'
                     +',"HLTARIF","HLDEBITEURS","AFMULFACTIERSAFF"'
                     +',"TRLISTESYNCHRO","TRSUPPRECRITURE","REIMPUTATION")');

  ExecuteSQL('DELETE FROM GRAPHS WHERE  GR_GRAPHE IN ("TWCBNVIEWER1","TWCBNVIEWER2","TWCBNVIEWER3","TWCBNVIEWER4") AND GR_LANGUE="---"');
End;

Procedure MajAvant705 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCSTKDEPLACEMENT","UORDREBES"'
        +',"UORDRELIG","GCPREPAEXPEXT", "GCPREPAEXPAFF", "GCPREPAEXPAFC"'
        +',"WPARCELEM","WVERSION","WVERSIONPARC","CPCONSAUXI","CPCONSGENE"'
        +')');
End;

Procedure MajAvant706 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCPREPAEXPEXT", "GCPREPAEXPAFF"'
        +', "GCPREPAEXPAFC", "GCPREPAEXP","WCOMPAREWPL","WCOMPAREWPLBTH","WCOMPAREWPLORD"'
        +',"YTARIFSMULART101","YTARIFSMULART201","YTARIFSMULART211","YTARIFSMULART301"'
        +',"YTARIFSMULART401","YTARIFSMULART501","YTARIFSMULART601","WORDREBES2CPD","WORDREBESL"'
        +',"PRT_AXE","CPCONSGENE"'
        +')');
End;


Procedure MajAvant707 ;
Begin
  //IMMO
  SupprimeEtat('E' ,'I24','090');
  SupprimeEtat('E' ,'I24','091');
  SupprimeEtat('E' ,'I24','092');
  SupprimeEtat('E' ,'I24','093');
  SupprimeEtat('E' ,'I15','020');
  SupprimeEtat('E' ,'I15','021');
  SupprimeEtat('E' ,'I15','030');
  // GPAO
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCSTKMANQUE","GCSTKDISPATCH","WARTPARC_VERSIONS")');
End;

Procedure MajAvant708 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE in ("WPARC","WORDREBESL","UORDREBESLASSAIS"'
              +',"UORDREBESLASSTOC","WUNITE","CPCONSGENE","CPCONSAUXI","CPMULCUTOFF")');

End;

Procedure MajAvant709 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WORDREBESL","GCSTKGQDATE","GCSTKGQDDATE"'
         +',"GCMULLIGNE","GCMULLIGNEACH","GCMULCONTREAVTGD","GCMULCONTREAVTGT","CPCONSAUXI"'
         +',"GCDUPLICPIECE","GCMULMODIFDOCVEN","GCMULPIECE","GCMULPIECEVISA","GCMULTRANSFDOCVEN")');
End;

Procedure MajAvant710 ;
Begin
   ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("CONTROLETVAS","WORDREBESL","GCGROUPEPIECE")');
   ExecuteSQL('DELETE FROM GRAPHS WHERE  GR_GRAPHE IN ("TWCBNCBB_CUBE","TWPDRVIEWER1") AND GR_LANGUE="---"');

   ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="QULPHASEITIP001"');
End;

Procedure MajAvant711 ;
Begin
  ExecuteSQL('DELETE FROM GRAPHS WHERE GR_GRAPHE IN ("TGCSTKDISPO_CUB","TWORDREGAMME_CUB","TGCDISPODETAIL_CU","TGCSTKMOUVEMENT_C") AND GR_LANGUE="---"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE="GCPREPAEXPE"');
  //PAIE
 SupprimeEtat ('E', 'PCS', 'PRE', TRUE); // Suppression etat modele paie
 SupprimeEtat ('E', 'PCS', 'PDC', TRUE); // Suppression etat modele paie
End;

Procedure MajAvant712 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("SUPPRECRS", "MULVMVTS", "MULMMVTS", "SSAISODA","WCBNPROPOSITION")');
  ExecuteSQL('UPDATE CHOIXCOD SET CC_LIBELLE="Des heures" WHERE CC_TYPE="YTN" AND CC_CODE="H"');
{$IFDEF MAJBOB}
  {$IFDEF MAJPCL}
    if IsdossierPCL then
    begin
      UpdateDomaine ('RESSOURCE','Y');
    end;
  {$ENDIF}
{$ENDIF}
End;

Procedure MajAvant713 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE in ("AFALIGNAFF","GCSTKMANQUE","AFSAISIEACTAFFTT","AFSAISIEACTRESTTS","AFSAISIEACTRESTTT")');
  //GC
  ExecuteSQL('DELETE FROM GRAPHS WHERE GR_GRAPHE = "TGCCLASSEABC_GA"') ;
  //CHR
  ExecuteSQL('UPDATE MENU SET MN_VERSIONDEV="-" WHERE MN_1 IN (161, 162, 163, 164, 165,166,167,286,287,288,289) '
      + ' AND MN_VERSIONDEV="X" AND MN_ACCESGRP<>"----------------------------------------------------------------------------------------------------"');
End;

Procedure MajAvant723 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("CPRELANCEDIV", "CPRELANCECLIENT","GCMULPORT")');
  ExecuteSQL('DELETE FROM GRAPHS WHERE GR_GRAPHE IN ("TWPDRVIEWER1","TWPDRVIEWER2")  AND GR_LANGUE="---"');
End;

Procedure MajAvant724 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("CPDECCHEQUEEDTNC","CPENCTRAITEEDTNC" , "CPENCPRELEVEMENT"'
     + ',"YTARIFSMULART101","YTARIFSMULART201","YTARIFSMULART211","YTARIFSMULART301" '
     + ',"YTARIFSMULART401","EDITIERS","YTARIFSMULART501","YTARIFSMULART601","YTARIFSMULTIE101" '
     + ',"YTARIFSMULTIE201","YTARIFSMULTIE211","YTARIFSMULTIE301","YTARIFSMULTIE401"'
     + ',"YTARIFSMULTIE501","YTARIFSMULTIE601")');
End;

Procedure MajAvant725 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WVERSIONPARC","MULVECRBUD","SUPPRBUDE","MULVALBUD","QULPHASEITIP001","RTMULTITIERSRECH")');
  ExecuteSQL('DELETE FROM GRAPHS WHERE GR_GRAPHE = ("TGCSTKDISPO_CUB") AND GR_LANGUE="---"');
End;

Procedure MajAvant726 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("JUCONTACT_MUL","PRT_ANNUINTERLOC")');
End;

Procedure MajAvant727 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WGAMMELIGMODIFLOT")');
  Vers_Afftiers:=TableToVersion ('AFFTIERS'); // crer le champ vers_afftiers
End;

Procedure MajAvant729 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE="GCLISTINVLIGNE"');
  //GIGA
  Vers_Tache := TableToVersion('TACHE');
  Vers_Afplanning:= TableToVersion('AFPLANNING'); //crer champs vers_aafplanning
End;

Procedure MajAvant730 ;
Begin
   ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="WORDREBESL"');
End;


Procedure MajAvant731 ;
Begin
  ExecuteSQL('DELETE FROM GRAPHS WHERE GR_GRAPHE = "TWCBNVIEWER6" AND GR_LANGUE="---"');
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE="WMARCHE"');
End;

Procedure MajAvant734 ;
Begin
   ExecuteSQL('DELETE FROM GRAPHS WHERE GR_GRAPHE = "TWCBNVIEWER6" AND GR_LANGUE="---"');
   ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE="WMARCHE"');

   // Traitement des tables de la GED
   TraitementTablesGed () ;
End;

Procedure MajAvant735 ;
Begin
   ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WPF_CBNQUOTAFO","CPCECRANA","CPCSECTANA")');
End;

Procedure MajAvant736 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("CPSUIVACC", "CPMODIFECHEMP","QULITI0"'
       +',"JUACTIONNAIRES","JUACTIONNAIRES2","JUALERTES_MUL","JUANNFUSION_LST","JUANNULIEN_LST","JUANNULIEN_MUL"'
       +',"JUANNULIEN_SEL","JUBAILINFO_MUL","JUBIBACTION_LST","JUBIBACTION_MUL","JUBIBMODULE_MUL","JUBIBRUBRIQUE_MUL"'
       +',"JUCONFIDENTIEL","JUDOSFUSION_LST","JUDOSOPACT","JUDOSOPER_MUL","JUDOSOPRUB_MUL","JUDOSSIER_MUL"'
       +',"JUEVEN_DOC_MUL","JUEVEN_REU_MUL","JUEVEN_TAC_MUL","JUEVENEMENT_MUL","JUFONCTION_MUL","JUHISTOREM_MUL"'
       +',"JUHISTOTITRES_MUL","JUMANDAT_LST","JUMODULE_LST","JUMVTREM_MUL","JUMVTTITRES_MUL","JUPERSONNE_MUL"'
       +',"JUPERSONNE_SEL","JUSTIFSLDA","JUTIFSLDA","JUTIFSLDG","JUTIFSLDGG","JUTYPEEX_MUL","JUTYPEFONCT_MUL"'
       +',"JUTYPEPER_MUL","YYMULANNDOSS","DPANNULIEN","DPANNULIEN2","DPANNULIEN3","DPANNUPERS","DPANNUPERS2"'
       +',"DPCREAT_TIERS","DPDOCUMENTS_EWS","DPDOCUMENTS_MUL","DPDOSSIEROBLIG","DPDOSSIEROBLIG2","DPENTGRPE"'
       +',"DPEVENEMENT","DPHISTCAPI","DPLIENFIL","DPLINTVJURI","DPLJURIACTI","DPMULANU","DPMULANU2","DPMULANU3"'
       +',"DPMULCTRFISC","DPMULNODOSS","YYDOSSIER_SEL","YYMESSAGES_IN","YYMESSAGES_INGLOB","YYMESSAGES_OUT"'
       +',"YYMESSAGESDOS_MUL","YYMODELESDOC_MUL","YYMULSELDOSS"'
       +')');
   //GPAO
   If IsMonoOuCommune then wResetWJT;
End;


Procedure MajAvant738 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("WPF_CBNGRPROP","WPF_DEPOTAPP"'
            +',"CPLISTECIRCUITBAP", "CPLISTETYPEVISA", "CPLISTEECRTTC", "CPLISTERELANCEBAP"'
            +',"CPLISTEPURGEBAP", "CPLISTEMODIFBAP")');

  // Pour init du traitement du DICO cas Entreprises et PCL : par défaut, pas de traitement du DICO
  ForceEmptyParamSoc('SO_INTEGREDICO', '-');
  ForceEmptyParamSoc('SO_SUPPRIMEDICO', 'X');
End;

Procedure MajAvant739 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("YYAGENDA","YYEVEN_REU_MUL"'
            +',"YYEVEN_TAC_MUL", "JUEVENEMENT_MUL"'
            +',"CPLISTEMAILBAP", "CPLISTESUIVIBAP","CPLISTERELANCEBAP"'
            +',"YYMESSAGES_OUT", "CPLISTEPURGEBAP", "CPLISTEMODIFBAP"'
            +',"GCMULPIECEACH","GCMULTRANSFDOCACH","GCMULMODIFDOCACH","WORDREBESL","AFDPTIERSMAILING")');
End;

Procedure MajAvant740 ;
Begin
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("YYREPERTPERSO_MUL"'
            +',"YYMULSELDOSS","WPDRTET","RQTLIENSART","RQTLIENSCAT","RQTLIENSTIE"'
            +',"YYLIENTABLETTE","AFMULAFFAIREEDIT","AFMULAVTAFFAIRE"'
            +',"AFMULPROPOSITION","AFMULRECHAFFAIRE","AFPLANNINGREVAL"'
            +')');
End;

// =====   MAJAVANT ================================================

function MAJHalleyAvant(VSoc: Integer; MajLab1, MajLab2: TLabel; MajJauge: TProgressBar): boolean;
begin
  L1 := MAJLab1;
  L2 := MAjLab2;
  J := MajJauge;
  L1.Caption:=TraduireMemoire('Préparation mise à jour');
  L2.Caption:='';
  // not IsDossierPCL:=not((V_PGI.ModePCL='1') and  (not ExisteSQL('select 1 from article'))) ;// si au moins un article alors c'est pas un dossier PCL ...
{$IFDEF DOSSIERPCL}
  //IsDossierPCL:=True;    // abandonné le 11/10/2004
{$ELSE}
  //IsDossierPCL:=False;    // abandonné le 11/10/2004
{$ENDIF}



{$IFDEF MAJBOB}
  {$IFDEF MAJPCL}
  IsDossierPCL:=V_PGI.DossierPCL;
  {$ELSE}
  IsDossierPCL:=(V_PGI.BobDeMiseAJour<>'');
  {$ENDIF}
{$ELSE}
  {$IFDEF MAJPCL}
  IsDossierPCL:=V_PGI.DossierPCL;
  {$ELSE}
  IsDossierPCL:=False;
  {$ENDIF}
{$ENDIF}


  if V_PGI.SAV then
  begin
    LogAGL('');
    LogAGL('**----------------- '+  DateTimeToStr(Now) + ' -----------------------');
    if V_PGI.ModePCL='1' then
      LogAGL('Traitement d''un dossier en environnement PCL')
    else
      LogAGL('Traitement d''un dossier en environnement Entreprise');
    if IsMonoOuCommune then
      LogAGL('    ==> base commune ou mono')
    else
      LogAGL('    ==> base dossier PCL') ;
    LogAGL('Version de la base   : '+IntToStr(VSoc));
    LogAGL('Version de la Socref : '+IntToStr(V_PGI.NumVersionBase));
    LogAGL('Début MajAvant ' + DateTimeToStr(Now));
  end;

  if VSoc < 545 then MajAvant545;
  if VSoc < 547 then MajAvant547;
  if VSoc < 550 then MajAvant550;
  if VSoc < 560 then MajAvant560;
  if VSoc < 561 then MajAvant561;
  if VSoc < 562 then MajAvant562;
  if VSoc < 565 then MajAvant565;
  if VSoc < 566 then MajAvant566;
  if VSoc < 569 then MajAvant569;
  if VSoc < 582 then MajAvant582;
  if VSoc < 585 then MajAvant585;
  if VSoc < 590 then MajAvant590;
  if VSoc < 591 then MajAvant591;
  if VSoc < 595 then MajAvant595;
  if VSoc < 602 then MajAvant602;
  if VSoc < 604 then MajAvant604;
  if VSoc < 605 then MajAvant605;
  if VSoc < 606 then MajAvant606;
  if VSoc < 607 then MajAvant607;
  if VSoc < 608 then MajAvant608;
  if VSoc < 609 then MajAvant609;
  if VSoc < 610 then MajAvant610;
  if VSoc < 612 then MajAvant612;
  if VSoc < 613 then MajAvant613;
  if VSoc < 614 then MajAvant614;
  if VSoc < 615 then MajAvant615;
  if VSoc < 616 then MajAvant616;
  if VSoc < 617 then MajAvant617;
  if VSoc < 618 then MajAvant618;
  if VSoc < 619 then MajAvant619;
  if VSoc < 620 then MajAvant620; //diffusion 5.0
  if VSoc < 621 then MajAvant621; //Correctif 5.05
  if VSoc < 625 then MajAvant625;
  if VSoc < 626 then MajAvant626;
  if VSoc < 627 then MajAvant627;
  if VSoc < 631 then MajAvant631;
  if VSoc < 633 then MajAvant633;
  if VSoc < 634 then MajAvant634;
  if VSoc < 635 then MajAvant635;
  if VSoc < 637 then MajAvant637;
  if VSoc < 640 then MajAvant640;
  if VSoc < 642 then MajAvant642;
  if VSoc < 644 then MajAvant644;
  if VSoc < 645 then MajAvant645;
  if VSoc < 646 then MajAvant646;
  if VSoc < 648 then MajAvant648;
  if VSoc < 649 then MajAvant649;
  if VSoc < 650 then MajAvant650;
  if VSoc < 651 then MajAvant651;
  if VSoc < 653 then MajAvant653;
  if VSoc < 654 then MajAvant654;
  if VSoc < 655 then MajAvant655;
  if VSoc < 656 then MajAvant656;
  if VSoc < 657 then MajAvant657;
  if VSoc < 658 then MajAvant658;
  if VSoc < 659 then MajAvant659;
  if VSoc < 660 then MajAvant660;
  if VSoc < 661 then MajAvant661;
  if VSoc < 662 then MajAvant662;
  if VSoc < 663 then MajAvant663;  // V6.00
  if VSoc < 667 then MajAvant667;
  if VSoc < 668 then MajAvant668;
  if VSoc < 669 then MajAvant669;
  if VSoc < 670 then MajAvant670;
  if VSoc < 671 then MajAvant671;
  if VSoc < 673 then MajAvant673;
  if VSoc < 674 then MajAvant674;
  if VSoc < 675 then MajAvant675;
  if VSoc < 676 then MajAvant676;
  if VSoc < 680 then MajAvant680;
  if VSoc < 681 then MajAvant681;
  if VSoc < 682 then MajAvant682;
  if VSoc < 683 then MajAvant683;
  if VSoc < 684 then MajAvant684;
  if VSoc < 685 then MajAvant685;
  if VSoc < 686 then MajAvant686;
  if VSoc < 687 then MajAvant687;
  if VSoc < 688 then MajAvant688;
  if VSoc < 689 then MajAvant689;
  if VSoc < 690 then MajAvant690;
  if VSoc < 691 then MajAvant691;
  if VSoc < 692 then MajAvant692; //6.30
  if VSoc < 694 then MajAvant694;
  if VSoc < 695 then MajAvant695;
  if VSoc < 696 then MajAvant696;
  if VSoc < 697 then MajAvant697;
  if VSoc < 698 then MajAvant698;
  if VSoc < 699 then MajAvant699;
  if VSoc < 700 then MajAvant700;
  if VSoc < 701 then MajAvant701;
  if VSoc < 702 then MajAvant702;
  if VSoc < 703 then MajAvant703;
  if VSoc < 704 then MajAvant704;
  if VSoc < 705 then MajAvant705;
  if VSoc < 706 then MajAvant706;
  if VSoc < 707 then MajAvant707;
  if VSoc < 708 then MajAvant708;
  if VSoc < 709 then MajAvant709;
  if VSoc < 710 then MajAvant710;
  if VSoc < 711 then MajAvant711;
  if VSoc < 712 then MajAvant712;
  if VSoc < 713 then MajAvant713;  // diffusion 714
  if VSoc < 723 then MajAvant723;
  if VSoc < 724 then MajAvant724;
  if VSoc < 725 then MajAvant725;
  if VSoc < 726 then MajAvant726;
  if VSoc < 727 then MajAvant727;
  if VSoc < 729 then MajAvant729;
  if VSoc < 730 then MajAvant730;
  if VSoc < 731 then MajAvant731;
  if VSoc < 734 then MajAvant734 ;
  if VSoc < 735 then MajAvant735 ;
  if VSoc < 736 then MajAvant736 ;
  if VSoc < 738 then MajAvant738 ;
  if VSoc < 739 then MajAvant739 ;
  if VSoc < 740 then MajAvant740 ;

  Result := True;
end;

{==============================================================================}
{=========================== PROCEDURES MAJ APRES =============================}
{==============================================================================}

procedure UpDateDecoupeEcr(SetSQL: string);
var DMin, DMax, DD1, DD2: TDateTime;
  ListeJ: TStrings;
  Q, QExo: TQuery;
  i, iper, Delta: integer;
begin
  // Lecture des journaux
  ListeJ := TStringList.Create;
  Q := OpenSQL('Select J_JOURNAL from JOURNAL', True);
  while not Q.EOF do
  begin
    ListeJ.Add(Q.Fields[0].AsString);
    Q.Next;
  end;
  Ferme(Q);
  // Balayage des écritures avec découpe
  for i := 0 to ListeJ.Count - 1 do
  begin
    QExo := OpenSQl('Select EX_EXERCICE, EX_DATEDEBUT,EX_DATEFIN from EXERCICE', True);
    while not QExo.EOF do
    begin
      DMin := QExo.Fields[1].AsDateTime;
      DMax := QExo.Fields[2].AsDateTime;
      Delta := Round((DMax - DMin) / 10);
      for iper := 1 to 10 do
      begin
        if iper < 10 then
        begin
          DD1 := DMin + (iper - 1) * Delta;
          DD2 := DD1 + Delta;
        end
        else
        begin
          DD1 := DMin + (iper - 1) * Delta;
          DD2 := DMax;
        end;
        BeginTrans;
        ExecuteSQL('UPDATE ECRITURE SET ' + SetSQL + ' WHERE E_JOURNAL="' + ListeJ[i] + '" AND E_EXERCICE="' + QExo.Fields[0].AsString +
          '" AND E_DATECOMPTABLE>="' + UsDateTime(DD1) + '" AND E_DATECOMPTABLE<="' + UsDateTime(DD2) + '"');
        CommitTrans;
      end;
      QExo.Next;
    end;
    Ferme(QExo);
  end;
  ListeJ.Clear;
  ListeJ.Free;
end;

procedure UpDateDecoupeAna(SetSQL: string);
var DMin, DMax, DD1, DD2: TDateTime;
  ListeJ: TStrings;
  Q, QExo: TQuery;
  i, iper, Delta: integer;
begin
  // Lecture des journaux
  ListeJ := TStringList.Create;
  Q := OpenSQL('Select J_JOURNAL from JOURNAL', True);
  while not Q.EOF do
  begin
    ListeJ.Add(Q.Fields[0].AsString);
    Q.Next;
  end;
  Ferme(Q);
  // Balayage des écritures avec découpe
  for i := 0 to ListeJ.Count - 1 do
  begin
    QExo := OpenSQl('Select EX_EXERCICE, EX_DATEDEBUT,EX_DATEFIN from EXERCICE', True);
    while not QExo.EOF do
    begin
      DMin := QExo.Fields[1].AsDateTime;
      DMax := QExo.Fields[2].AsDateTime;
      Delta := Round((DMax - DMin) / 10);
      for iper := 1 to 10 do
      begin
        if iper < 10 then
        begin
          DD1 := DMin + (iper - 1) * Delta;
          DD2 := DD1 + Delta;
        end
        else
        begin
          DD1 := DMin + (iper - 1) * Delta;
          DD2 := DMax;
        end;
        BeginTrans;
        ExecuteSQL('UPDATE ANALYTIQ SET ' + SetSQL + ' WHERE Y_JOURNAL="' + ListeJ[i] + '" AND Y_EXERCICE="' + QExo.Fields[0].AsString +
          '" AND Y_DATECOMPTABLE>="' + UsDateTime(DD1) + '" AND Y_DATECOMPTABLE<="' + UsDateTime(DD2) + '"');
        CommitTrans;
      end;
      QExo.Next;
    end;
    Ferme(QExo);
  end;
  ListeJ.Clear;
  ListeJ.Free;
end;

procedure UpDateDecoupeLigne(SetSQL: string; SetCondition : string ='');
var NMin, NMax, NCount, NN1, NN2, NbStep: integer;
  ListeNat, ListeSouche: TStrings;
  Q: TQuery;
  QMax: TQuery;
  i, k: integer;
  NbDebutTranche: integer;
  NbFinTranche: integer;
  //PremierNumero: integer;
  DernierNumero: integer;
  //SQL : string ;
begin
  // Lecture des natures
  ListeNat := TStringList.Create;
  ListeSouche := TStringList.Create;
  Q := OpenSQL('Select DISTINCT GL_NATUREPIECEG, GL_SOUCHE FROM LIGNE', True);
  while not Q.EOF do
  begin
    ListeNat.Add(Q.Fields[0].AsString);
    ListeSouche.Add(Q.Fields[1].AsString);
    Q.Next;
  end;
  Ferme(Q);
  // Balayage des pièces avec découpe
  for i := 0 to ListeNat.Count - 1 do
  begin
    // Récupération du plus grand numéros possible
    QMax := OpenSQL('SELECT MIN(GL_NUMERO), MAX(GL_NUMERO),COUNT(*) FROM LIGNE WHERE GL_NATUREPIECEG="' + ListeNat[i] + '" AND GL_SOUCHE="' + ListeSouche[i] +
                    '"', True);
    //PremierNumero := QMax.Fields[0].AsInteger;
    DernierNumero := QMax.Fields[1].AsInteger;
    NCount := QMax.Fields[2].AsInteger;
    Ferme(QMax);

    if NCount > 50000 then
    begin
      // Première tranche traitée : 0 à 1000000
      NbDebutTranche := 0;
      NbFinTranche := 1000000;

      while (NbDebutTranche < DernierNumero) do
      begin
        // Récupère le premier et le dernier numéro à traiter dans la tranche
        QMax := OpenSQL('SELECT MIN(GL_NUMERO), MAX(GL_NUMERO),COUNT(*) FROM LIGNE WHERE GL_NATUREPIECEG="' + ListeNat[i] + '" AND GL_SOUCHE="' +
          ListeSouche[i] + '" AND GL_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND GL_NUMERO<' + IntToStr(NbFinTranche) + '', True);
        NMin := QMax.Fields[0].AsInteger;
        NMax := QMax.Fields[1].AsInteger;
        NCount := QMax.Fields[2].AsInteger;
        Ferme(QMax);

        if NCount > 0 then
        begin
          NbStep := (NMax - NMin) div 25000;
          for k := 0 to NbStep do
          begin
            NN1 := NMin + k * 25000;
            if k < NbStep then NN2 := NN1 + 25000 else NN2 := NMax;
            BeginTrans;
            ExecuteSQL('UPDATE LIGNE SET ' + SetSQL + ' WHERE GL_NATUREPIECEG="' + ListeNat[i] + '" AND GL_SOUCHE="' + ListeSouche[i] + '" AND GL_NUMERO>=' +
                        IntToStr(NN1) + ' AND GL_NUMERO<=' + IntToStr(NN2) + ' ' + SetCondition );
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
      // sinon trou ds la numéraotation, update pour la souche
    end else ExecuteSQL('UPDATE LIGNE SET ' + SetSQL + ' WHERE GL_NATUREPIECEG="' + ListeNat[i] + '" AND GL_SOUCHE="' + ListeSouche[i] + '"' + ' ' + SetCondition);
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;

procedure UpDateDecoupePiece(SetSQL: string);
var NMin, NMax, NCount, NN1, NN2, NbStep: integer;
  ListeNat, ListeSouche: TStrings;
  Q: TQuery;
  QMax: TQuery;
  i, k: integer;
  NbDebutTranche: integer;
  NbFinTranche: integer;
  //PremierNumero: integer;
  DernierNumero: integer;
begin
  // Lecture des natures
  ListeNat := TStringList.Create;
  ListeSouche := TStringList.Create;
  Q := OpenSQL('Select DISTINCT GP_NATUREPIECEG, GP_SOUCHE FROM PIECE', True);
  while not Q.EOF do
  begin
    ListeNat.Add(Q.Fields[0].AsString);
    ListeSouche.Add(Q.Fields[1].AsString);
    Q.Next;
  end;
  Ferme(Q);

  // Balayage des pièces avec découpe
  for i := 0 to ListeNat.Count - 1 do
  begin
    // Récupération du plus grand numéros possible
    QMax := OpenSQL('SELECT MIN(GP_NUMERO), MAX(GP_NUMERO),COUNT(*) FROM PIECE WHERE GP_NATUREPIECEG="' + ListeNat[i] + '" AND GP_SOUCHE="' + ListeSouche[i] +
      '"', True);
    //PremierNumero := QMax.Fields[0].AsInteger;
    DernierNumero := QMax.Fields[1].AsInteger;
    NCount := QMax.Fields[2].AsInteger;
    Ferme(QMax);

    // si plus de 50000, update par tranche
    if NCount > 50000 then
    begin
      // Première tranche traitée : 0 à 1000000
      NbDebutTranche := 0;
      NbFinTranche := 1000000;

      while (NbDebutTranche < DernierNumero) do
      begin
        // Récupère le premier et le dernier numéro à traiter dans la tranche
        QMax := OpenSQL('SELECT MIN(GP_NUMERO), MAX(GP_NUMERO),COUNT(*) FROM PIECE WHERE GP_NATUREPIECEG="' + ListeNat[i] + '" AND GP_SOUCHE="' +
          ListeSouche[i] + '" AND GP_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND GP_NUMERO<' + IntToStr(NbFinTranche) + '', True);
        NMin := QMax.Fields[0].AsInteger;
        NMax := QMax.Fields[1].AsInteger;
        NCount := QMax.Fields[2].AsInteger;
        Ferme(QMax);

        if NCount > 0 then
        begin
          NbStep := (NMax - NMin) div 25000;
          for k := 0 to NbStep do
          begin
            NN1 := NMin + k * 25000;
            if k < NbStep then NN2 := NN1 + 25000 else NN2 := NMax;
            BeginTrans;
            ExecuteSQL('UPDATE PIECE SET ' + SetSQL + ' WHERE GP_NATUREPIECEG="' + ListeNat[i] + '" AND GP_SOUCHE="' + ListeSouche[i] + '" AND GP_NUMERO>=' +
              IntToStr(NN1) + ' AND GP_NUMERO<=' + IntToStr(NN2));
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
    end
      // sinon trou ds la numérotation, update pour la souche
    else
    begin
      ExecuteSQL('UPDATE PIECE SET ' + SetSQL
        + ' WHERE GP_NATUREPIECEG="' + ListeNat[i] + '" AND GP_SOUCHE="' + ListeSouche[i] + '"');
    end;
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;

procedure UpDateDecoupePiedEche(SetSQL: string);
var NMin, NMax, NCount, NN1, NN2, NbStep: integer;
  ListeNat, ListeSouche: TStrings;
  Q: TQuery;
  QMax: TQuery;
  i, k: integer;
  NbDebutTranche: integer;
  NbFinTranche: integer;
  //PremierNumero: integer;
  DernierNumero: integer;
begin
  // Lecture des natures
  ListeNat := TStringList.Create;
  ListeSouche := TStringList.Create;
  Q := OpenSQL('Select DISTINCT GPE_NATUREPIECEG, GPE_SOUCHE FROM PIEDECHE', True);
  while not Q.EOF do
  begin
    ListeNat.Add(Q.Fields[0].AsString);
    ListeSouche.Add(Q.Fields[1].AsString);
    Q.Next;
  end;
  Ferme(Q);

  // Balayage des pièces avec découpe
  for i := 0 to ListeNat.Count - 1 do
  begin
    // Récupération du plus grand numéros possible
    QMax := OpenSQL('SELECT MIN(GPE_NUMERO), MAX(GPE_NUMERO),COUNT(*) FROM PIEDECHE WHERE GPE_NATUREPIECEG="' + ListeNat[i]
     + '" AND GPE_SOUCHE="' + ListeSouche[i] + '"', True);
    //PremierNumero := QMax.Fields[0].AsInteger;
    DernierNumero := QMax.Fields[1].AsInteger;
    NCount := QMax.Fields[2].AsInteger;
    Ferme(QMax);

    // si plus de 50000, update par tranche
    if NCount > 50000 then
    begin
      // Première tranche traitée : 0 à 1000000
      NbDebutTranche := 0;
      NbFinTranche := 1000000;

      while (NbDebutTranche < DernierNumero) do
      begin
        // Récupère le premier et le dernier numéro à traiter dans la tranche
        QMax := OpenSQL('SELECT MIN(GPE_NUMERO), MAX(GPE_NUMERO),COUNT(*) FROM PIEDECHE WHERE GPE_NATUREPIECEG="' + ListeNat[i] + '" AND GPE_SOUCHE="' +
          ListeSouche[i] + '" AND GPE_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND GPE_NUMERO<' + IntToStr(NbFinTranche) + '', True);
        NMin := QMax.Fields[0].AsInteger;
        NMax := QMax.Fields[1].AsInteger;
        NCount := QMax.Fields[2].AsInteger;
        Ferme(QMax);

        if NCount > 0 then
        begin
          NbStep := (NMax - NMin) div 25000;
          for k := 0 to NbStep do
          begin
            NN1 := NMin + k * 25000;
            if k < NbStep then NN2 := NN1 + 25000 else NN2 := NMax;
            BeginTrans;
            ExecuteSQL('UPDATE PIEDECHE SET ' + SetSQL + ' WHERE GPE_NATUREPIECEG="' + ListeNat[i] + '" AND GPE_SOUCHE="' + ListeSouche[i] + '" AND GPE_NUMERO>=' +
              IntToStr(NN1) + ' AND GPE_NUMERO<=' + IntToStr(NN2));
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
    end
      // sinon trou ds la numérotation, update pour la souche
    else
    begin
      ExecuteSQL('UPDATE PIEDECHE SET ' + SetSQL
        + ' WHERE GPE_NATUREPIECEG="' + ListeNat[i] + '" AND GPE_SOUCHE="' + ListeSouche[i] + '"');
    end;
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;

procedure UpDateDecoupeAcomptes(SetSQL: string);
var NMin, NMax, NCount, NN1, NN2, NbStep: integer;
  ListeNat, ListeSouche: TStrings;
  Q: TQuery;
  QMax: TQuery;
  i, k: integer;
  NbDebutTranche: integer;
  NbFinTranche: integer;
  //PremierNumero: integer;
  DernierNumero: integer;
begin
  // Lecture des natures
  ListeNat := TStringList.Create;
  ListeSouche := TStringList.Create;
  Q := OpenSQL('Select DISTINCT GAC_NATUREPIECEG, GAC_SOUCHE FROM ACOMPTES', True);
  while not Q.EOF do
  begin
    ListeNat.Add(Q.Fields[0].AsString);
    ListeSouche.Add(Q.Fields[1].AsString);
    Q.Next;
  end;
  Ferme(Q);

  // Balayage des pièces avec découpe
  for i := 0 to ListeNat.Count - 1 do
  begin
    // Récupération du plus grand numéros possible
    QMax := OpenSQL('SELECT MIN(GAC_NUMERO), MAX(GAC_NUMERO),COUNT(*) FROM ACOMPTES WHERE GAC_NATUREPIECEG="' + ListeNat[i]
     + '" AND GAC_SOUCHE="' + ListeSouche[i] + '"', True);
    //PremierNumero := QMax.Fields[0].AsInteger;
    DernierNumero := QMax.Fields[1].AsInteger;
    NCount := QMax.Fields[2].AsInteger;
    Ferme(QMax);

    // si plus de 50000, update par tranche
    if NCount > 50000 then
    begin
      // Première tranche traitée : 0 à 1000000
      NbDebutTranche := 0;
      NbFinTranche := 1000000;

      while (NbDebutTranche < DernierNumero) do
      begin
        // Récupère le premier et le dernier numéro à traiter dans la tranche
        QMax := OpenSQL('SELECT MIN(GAC_NUMERO), MAX(GAC_NUMERO),COUNT(*) FROM ACOMPTES WHERE GAC_NATUREPIECEG="' + ListeNat[i] + '" AND GAC_SOUCHE="' +
          ListeSouche[i] + '" AND GAC_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND GAC_NUMERO<' + IntToStr(NbFinTranche) + '', True);
        NMin := QMax.Fields[0].AsInteger;
        NMax := QMax.Fields[1].AsInteger;
        NCount := QMax.Fields[2].AsInteger;
        Ferme(QMax);

        if NCount > 0 then
        begin
          NbStep := (NMax - NMin) div 25000;
          for k := 0 to NbStep do
          begin
            NN1 := NMin + k * 25000;
            if k < NbStep then NN2 := NN1 + 25000 else NN2 := NMax;
            BeginTrans;
            ExecuteSQL('UPDATE ACOMPTES SET ' + SetSQL + ' WHERE GAC_NATUREPIECEG="' + ListeNat[i] + '" AND GAC_SOUCHE="' + ListeSouche[i] + '" AND GAC_NUMERO>=' +
              IntToStr(NN1) + ' AND GAC_NUMERO<=' + IntToStr(NN2));
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
    end
      // sinon trou ds la numérotation, update pour la souche
    else
    begin
      ExecuteSQL('UPDATE ACOMPTES SET ' + SetSQL
        + ' WHERE GAC_NATUREPIECEG="' + ListeNat[i] + '" AND GAC_SOUCHE="' + ListeSouche[i] + '"');
    end;
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;

procedure MAJVER541;
begin
  BeginTrans;
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  ExecuteSQL('UPDATE ARTICLE SET GA_NATUREPRES="",GA_CALCPRIXPR=""');
  ExecuteSQL('UPDATE NOMENENT SET GNE_QTEDUDETAIL=1');
  ExecuteSQL('UPDATE NOMENLIG SET GNL_PRIXFIXE=0');
  ExecuteSQL('UPDATE PROFILART SET GPF_CALCPRIXPR="", GPF_COEFCALCPR=0, GPF_CALCAUTOPR="-"');

  CommitTrans;
  UpdateDecoupeLigne('GL_QTESIT=0, GL_NIVEAUIMBRIC=0');
end;

procedure MAJVER542;
begin
  BeginTrans;
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  ExecuteSQL('UPDATE CHAMPSPRO SET RCL_NOMONGLET=""');
  CommitTrans;
end;

procedure MAJVER543;
begin
  BeginTrans;
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  ExecuteSQL('UPDATE ARTICLE SET GA_DESIGNATION1="", GA_DESIGNATION2="", GA_TYPEEMPLACE=""');
  CommitTrans;
end;

procedure MAJVER544;
var SQL, SS: string;
begin
  SS := UsDateTime(IDate1900);
  BeginTrans;
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SQL :=
    'UPDATE PROSPECTCOMPL SET RPC_RPCLIBTEXTE3="",RPC_RPCLIBTEXTE4="",RPC_RPCLIBTEXTE5="",RPC_RPCLIBTEXTE6="",RPC_RPCLIBTEXTE7="",RPC_RPCLIBTEXTE8="",RPC_RPCLIBTEXTE9="",RPC_RPCLIBTEXTEA="",RPC_RPCLIBTEXTEB=""';
  SQL := SQL +
    ' ,RPC_RPCLIBVAL3=0,RPC_RPCLIBVAL4=0,RPC_RPCLIBVAL5=0,RPC_RPCLIBVAL6=0,RPC_RPCLIBVAL7=0,RPC_RPCLIBVAL8=0,RPC_RPCLIBVAL9=0,RPC_RPCLIBVALA=0,RPC_RPCLIBVALB=0';
  SQL := SQL +
    ' ,RPC_RPCLIBTABLE3="",RPC_RPCLIBTABLE4="",RPC_RPCLIBTABLE5="",RPC_RPCLIBTABLE6="",RPC_RPCLIBTABLE7="",RPC_RPCLIBTABLE8="",RPC_RPCLIBTABLE9="",RPC_RPCLIBTABLEA="",RPC_RPCLIBTABLEB="" ';
  SQL := SQL +
    ' ,RPC_RPCLIBBOOL3="-",RPC_RPCLIBBOOL4="-",RPC_RPCLIBBOOL5="-",RPC_RPCLIBBOOL6="-",RPC_RPCLIBBOOL7="-",RPC_RPCLIBBOOL8="-",RPC_RPCLIBBOOL9="-",RPC_RPCLIBBOOLA="-",RPC_RPCLIBBOOLB="-" ';
  SQL := SQL + ' ,RPC_RPCLIBDATE3="' + SS + '",RPC_RPCLIBDATE4="' + SS + '",RPC_RPCLIBDATE5="' + SS + '",RPC_RPCLIBDATE6="' + SS + '",RPC_RPCLIBDATE7="' + SS +
    '",RPC_RPCLIBDATE8="' + SS + '",RPC_RPCLIBDATE9="' + SS + '",RPC_RPCLIBDATEA="' + SS + '",RPC_RPCLIBDATEB="' + SS + '"';
  ExecuteSQL(SQL);
  ExecuteSQL('UPDATE EMPLACEMENT SET GEM_MONOARTICLE="-", GEM_EMPLACEOCCUPE="-"');
  ExecuteSQL('UPDATE ACTIVITE SET ACT_NUMAPPREC=0, ACT_ETATVISA="VIS",ACT_ETATVISAFAC="VIS",ACT_VISEUR="",ACT_VISEURFAC="",ACT_DATEVISA="' + SS +
    '",ACT_DATEVISAFAC="' + SS + '"');
  SQL := 'UPDATE RESSOURCE SET ARS_PVHT2=0,ARS_PVTTC2=0,ARS_PVHTCALCUL2=0,ARS_DATEPRIX2="' + SS + '",ARS_TAUXREVIENTUN2=0';
  SQL := SQL + ' ,ARS_PVHT3=0,ARS_PVTTC3=0,ARS_PVHTCALCUL3=0,ARS_DATEPRIX3="' + SS + '",ARS_TAUXREVIENTUN3=0';
  SQL := SQL + ' ,ARS_PVHT4=0,ARS_PVTTC4=0,ARS_PVHTCALCUL4=0,ARS_DATEPRIX4="' + SS + '",ARS_TAUXREVIENTUN4=0';
  SQL := SQL + ' ,ARS_CALCULPV="-",ARS_PVHTCALCUL=ARS_PVHT,ARS_DATEPRIX="' + SS + '"';
  ExecuteSQL(SQL);
  ExecuteSQL('UPDATE RESSOURCEPR SET ARP_TYPEVALO="R"');
  ExecuteSQL('UPDATE AFTABLEAUBORD SET ATB_NATUREPIECEG="",ATB_SOUCHE="",ATB_NUMEROACHAT=0,ATB_INDICEG=0');
  ExecuteSQL('UPDATE FACTAFF SET AFA_TYPECHE="NOR"');

  CommitTrans;
end;

procedure MAJVER545;
var SQL: string;
begin
  BeginTrans;
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  (* mcd 032003SQL := 'UPDATE FACTAFF SET AFA_PROFILGENER="",AFA_GENERAUTO="",AFA_JUSTIFBONI2="",AFA_BONIMALIQTE=0';
  SQL := SQL+ ' , AFA_BM1FO=0, AFA_BM1FOCON=0, AFA_BM1FODEV=0, AFA_BM1FR=0, AFA_BM1FRCON=0, AFA_BM1FRDEV=0';
  SQL := SQL+ ' , AFA_BM2PRQTE=0, AFA_BM2PR=0, AFA_BM2PRDEV=0, AFA_BM2PRCON=0, AFA_BM2FO=0, AFA_BM2FODEV=0, AFA_BM2FOCON=0';
  SQL := SQL+ ' , AFA_BM2FR=0, AFA_BM2FRDEV=0, AFA_BM2FRCON=0, AFA_AFACTOT=0, AFA_AFACTOTDEV=0, AFA_AFACTOTCON=0';
  SQL := SQL+ ' , AFA_AFACTURERQTE=0, AFA_AFACTURER=0, AFA_AFACTURERDEV=0, AFA_AFACTURERCON=0, AFA_AFACTFR=0, AFA_AFACTFRDEV=0';
  SQL := SQL+ ' , AFA_AFACTFRCON=0, AFA_AFACTFO=0, AFA_AFACTFODEV=0, AFA_AFACTFOCON=0'; {, AFA_TAFTFAQTE=0'; // supprimé de la table en 560 }
  SQL := SQL+ ' , AFA_LIBELLE1="",AFA_LIBELLE2="",AFA_LIBELLE3=""'; *)
  //mcd 032003
  SQL := 'UPDATE FACTAFF SET AFA_PROFILGENER="",AFA_GENERAUTO="",AFA_JUSTIFBONI2="",AFA_BONIMALIQTE=0';
  SQL := SQL + ' , AFA_BM1FO=0,  AFA_BM1FODEV=0, AFA_BM1FR=0, AFA_BM1FRDEV=0';
  SQL := SQL + ' , AFA_BM2PRQTE=0, AFA_BM2PR=0, AFA_BM2PRDEV=0,  AFA_BM2FO=0, AFA_BM2FODEV=0';
  SQL := SQL + ' , AFA_BM2FR=0, AFA_BM2FRDEV=0,  AFA_AFACTOT=0, AFA_AFACTOTDEV=0';
  SQL := SQL + ' , AFA_AFACTURERQTE=0, AFA_AFACTURER=0, AFA_AFACTURERDEV=0, AFA_AFACTFR=0, AFA_AFACTFRDEV=0';
  SQL := SQL + ' ,  AFA_AFACTFO=0, AFA_AFACTFODEV=0';
  SQL := SQL + ' , AFA_LIBELLE1="",AFA_LIBELLE2="",AFA_LIBELLE3=""';

  ExecuteSQL(SQL);
  ExecuteSQL('UPDATE AFFAIRE SET AFF_NUMEROCONTACT=0');
  CommitTrans;
end;

procedure MAJVER546;
var SS: string;
  HH: TDateTime;
  StH: string;
begin
  SS := UsDateTime(IDate1900);
  HH := 0;
  StH := USTIME(HH);
  BeginTrans;
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  ExecuteSQL('UPDATE PARPIECE SET GPP_PROCLI="-"');
  ExecuteSQL('UPDATE HREVENEMENT SET HEV_COULEUR=""');
  ExecuteSQL('UPDATE HRPARAMPLANNING SET HPP_AFFICHEICONE="-",HPP_HEUREDEBUT="' + StH + '",HPP_HEUREFIN="' + StH + '"');
  ExecuteSQL('UPDATE TIERS SET T_REPRESENTANT=""');
  ExecuteSQL('UPDATE ABSENCESALARIE SET PCN_SAISIEDEPORTEE="-",PCN_VALIDSALARIE="ATT",PCN_VALIDRESP="ATT",PCN_EXPORTOK="-",PCN_LIBCOMPL1="",PCN_LIBCOMPL2="",PCN_DEBUTDJ="MAT",PCN_FINDJ="PAM"');
  ExecuteSQL('UPDATE COTISATION SET PCT_BASECRDS="-",PCT_BRUTSS="-",PCT_PLAFONDSS="-",PCT_BRUTFISC="-",PCT_NETIMPO="-"');
  ExecuteSQL('UPDATE CONTRATTRAVAIL SET PCI_CONDEMPLOI="W"');
  ExecuteSQL('UPDATE MOTIFABSENCE SET PMA_MOTIFEAGL="-"');
  ExecuteSQL('UPDATE HISTOSALARIE SET PHS_DADSPROF="",PHS_BDADSPROF="-",PHS_DADSCAT="",PHS_BDADSCAT="-"');
  ExecuteSQL('UPDATE SALARIES SET PSA_DADSPROF="",PSA_DADSCAT="",PSA_TAUXPARTSS=0');
  CommitTrans;
  UpdateDecoupePiece('GP_MAJLIBRETIERS="", GP_GENERAUTO="", GP_FACREPRISE=0');
  UpdateDecoupeLigne('GL_PUHTORIGINE=0,GL_PUHTORIGINEDEV=0'); // ,GL_PUHTORIGINECON=0 supprimé en 604
  if TableExiste('PARACTIONS') then RTMoveTypeAction;
end;

procedure MAJVER547;
begin
  BeginTrans;
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  ExecuteSQL('UPDATE MODEREGL SET MR_EINTEGREAUTO="-"');
  CommitTrans;
end;

procedure MAJVER548;
begin
  BeginTrans;
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  // ExecuteSQL('UPDATE ACTIVITE SET ACT_TRAITE="-",ACT_AVANTVENTE="-"');
  ExecuteSQL('UPDATE ACTIVITE SET ACT_AVANTVENTE="-"'); // modif MCD 23032004 version 637
  ExecuteSQL('UPDATE AFTABLEAUBORD SET ATB_BONIQTE=0,ATB_BONIQTEUNITERE=0,ATB_BONIVENTE=0,ATB_BONIPR=0'); {,ATB_CUTOFF=0 supprimé en 562 }
  CommitTrans;
end;

{***********A.G.L.***********************************************
Auteur  ...... : JCF
Créé le ...... : 02/08/2001
Modifié le ... :   /  /
Description .. : Resynchronisation de la MODE qui s'est arrêtée sur une
Suite ........ : base 530 entre février et août 2001.
Suite ........ : Les mises à jour de structure ont été effectuées sur le
Suite ........ : numéro de la table uniquement.
Suite ........ : Modification de cette méthode possible dès migration des
Suite ........ : clients de la MODE au 01-08-01 :
Suite ........ : QUIK (10 sites) - LMV (22 sites) - OL (2 sites) - MDF (2 sites)
Suite ........ : - MONDEX (2 sites)
Mots clefs ... : MODE;AREVOIR;MAJVER
*****************************************************************}

procedure MAJVER549;
var SS: string;
begin
  SS := UsDateTime(IDate1900);
  BeginTrans;
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if Vers_Dispo < 120 then ExecuteSQL('UPDATE DISPO SET GQ_ECARTINV=0');
  if Vers_ClavierEcran < 101 then ExecuteSQL('UPDATE CLAVIERECRAN SET CE_ALLERA=-1');
  if Vers_JoursCaisse < 101 then ExecuteSQL('UPDATE JOURSCAISSE SET GJC_VENDOUV="",GJC_VENDFERME=""');
  if Vers_JoursEtabEvt < 101 then ExecuteSQL('UPDATE JOURSETABEVT SET GET_ETABLISSEMENT="' + GetParamSoc('SO_ETABLISDEFAUT') + '"');
  if Vers_ListeInvent < 103 then ExecuteSQL('UPDATE LISTEINVENT SET GIE_STOCKCLOS="-"');
  // GPK_PRELEVEMENT="", supprimé en 582
  if Vers_ParCaisse < 103 then
    ExecuteSQL('UPDATE PARCAISSE SET GPK_REMARRONDI="",GPK_MDPFDCAIS="",GPK_MDPCTRLCAIS="",GPK_FDCAISSE="",GPK_TIROIRPIN5="-",GPK_NBEXEMPBON=1');
  if Vers_ParCaisse < 104 then ExecuteSQL('UPDATE PARCAISSE SET GPK_TOXAPPEL1="",GPK_TOXAPPEL2="",GPK_IMPMODFDC=""');
  if Vers_ParCaisse < 105 then ExecuteSQL('update PARCAISSE set GPK_REMPOURMAX2=0, GPK_REMPOURMAX3=0, GPK_PWDREM2="", GPK_PWDREM3="", GPK_REMPIEDLIG="X"');
  if Vers_ParPiece < 146 then ExecuteSQL('UPDATE PARPIECE SET GPP_TYPEFACT=""');
  if Vers_TarifMode < 105 then ExecuteSQL('UPDATE TARIFMODE SET GFM_CASCADE="-"');
  if Vers_TarifMode < 106 then ExecuteSQL('UPDATE TARIFMODE SET GFM_PROMO="-"');
  if Vers_TarifPer < 101 then ExecuteSQL('UPDATE TARIFPER SET GFP_ETABLISSEMENT="..."');
  if Vers_TarifPer < 103 then ExecuteSQL('UPDATE TARIFPER SET GFP_CASCADE="-"');
  if Vers_TarifPer < 104 then ExecuteSQL('UPDATE TARIFPER SET GFP_PROMO="-"');
  if TableToVersion('TYPEMASQUE') <= 100 then RecupTabletteChoixCodeGCTYPEMASQUE;
  if Vers_Etabliss < 18 then ExecuteSQL('UPDATE ETABLISS SET ET_TYPETARIF="",ET_DEVISE=""');

  ExecuteSQL('UPDATE F2072_JOUISSEUR SET J72_CONSISTANCE=""');
  ExecuteSQL('UPDATE JUTYPEFONCT SET JTF_DIRIGEANT="-",JTF_TITRE="-"');
  ExecuteSQL('UPDATE IMMOLOG SET IL_BASETAXEPRO=0');
  ExecuteSQL('UPDATE AFTABLEAUBORD  SET ATB_TYPEENR="NOR"');
  ExecuteSQL('UPDATE PROFILGENER  SET APG_LIBSUP=""');
  ExecuteSQL('UPDATE CALENDRIERREGLE  SET ACG_NBDJMAXANL=0,ACG_NBHMAXJL=0,ACG_NBHMAXSMNL=0,ACG_NBMAXHANALERTL=0,ACG_NBMAXHANBLOCL=0');
  CommitTrans;
end;

procedure MAJVER550;
begin
  BeginTrans;
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  ExecuteSQL('UPDATE PERSPECTIVES  SET RPE_REPRESENTANT=""');
  ExecuteSQL('UPDATE PERSPHISTO  SET RPH_REPRESENTANT=""');
  CommitTrans;
end;

procedure MAJVER560;
var SS: string;
begin
  SS := UsDateTime(IDate1900);
  BeginTrans;
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  ExecuteSQL('UPDATE ETABLISS SET ET_SURSITEDISTANT="-"');
  ExecuteSQL('UPDATE DEPOTS SET GDE_SURSITEDISTANT="-"');
  ExecuteSQL('UPDATE ARTICLELIE SET GAL_QTEREF="-"');
  ExecuteSQL('UPDATE DONNEURORDRE SET PDO_PGMODEREGLE=PDO_MODEREGLE');
  ExecuteSQL('UPDATE ETABCOMPL SET ETB_PGMODEREGLE=ETB_MODEREGLE, ETB_MEDTRAV=0,ETB_CODEDDTEFP=0');
  ExecuteSQL('UPDATE PAIEENCOURS SET PPU_PGMODEREGLE=PPU_MODEREGLE');
  ExecuteSQL('UPDATE DADSDETAIL SET PDS_DONNEEAFFICH = ""');
  ExecuteSQL('UPDATE COTISATION SET PCT_DADSTOTIMPTSS="-",PCT_DADSMONTTSS="-"');
  ExecuteSQL('UPDATE REMUNERATION SET  PRM_HCHOMPAR="-"');
  ExecuteSQL('UPDATE HISTOSALARIE SET PHS_TAUXPARTIEL="",PHS_BTAUXPARTIEL="-"');
  ExecuteSQL('UPDATE SALARIES SET PSA_TAUXPARTIEL=0,PSA_PGMODEREGLE=PSA_MODEREGLE,PSA_DADSDATE="' + SS + '"');
  //mcd 032003 ExecuteSQL('UPDATE FACTAFF SET AFA_ACPTEQTE=0,AFA_ACPTEPR=0,AFA_ACPTEPRCON=0,AFA_ACPTEPRDEV=0,AFA_ACPTEFR=0,AFA_ACPTEFRCON=0,AFA_ACPTEFRDEV=0,AFA_ACPTEFO=0,AFA_ACPTEFOCON=0,AFA_ACPTEFODEV=0');
  ExecuteSQL('UPDATE FACTAFF SET AFA_ACPTEQTE=0,AFA_ACPTEPR=0,AFA_ACPTEPRDEV=0,AFA_ACPTEFR=0,AFA_ACPTEFRDEV=0,AFA_ACPTEFO=0,AFA_ACPTEFODEV=0'); //mcd 032003
  ExecuteSQL('UPDATE AFFAIRE SET AFF_TOTALPR=0');
  ExecuteSQL('UPDATE REMUNERATION SET PRM_BASEMTQTE="004" WHERE PRM_BASEMTQTE="-"');
  ExecuteSQL('UPDATE REMUNERATION SET PRM_BASEMTQTE="001" WHERE PRM_BASEMTQTE="X"');
  ExecuteSQL('UPDATE DPFISCAL SET DFI_REDEVABILITE=""');
  ExecuteSQL('UPDATE AFCUMUL SET ACU_NUMPIECE=""'); {ACU_CUTOFFCALC=0, supprimé en 562 }
  CommitTrans;
  // AddNewNaturesGC ;   voir version MAJVER568
  UpdateDecoupePiece('GP_RESSOURCE=""');
  InsertChoixCode('ARF', 'ARE', 'Sur affaire de référence', 'Aff. référence', 'C');
  InsertChoixCode('ARF', 'AUC', 'Pas de regroupement', 'Aucun', 'A');
  InsertChoixCode('ARF', 'RE1', 'Regroupement 1 ', 'Regroupement 1', 'B');
  InsertChoixCode('ARF', 'RE2', 'Regroupement 2', 'Regroupement 2', 'B');
  InsertChoixCode('ARF', 'RE3', 'Regroupement 3', 'Regroupement 3', 'B');
end;

procedure MAJVER561;
var SS: string;
  i: integer;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SS := UsDateTime(IDate1900);
  BeginTrans;
  ExecuteSQL('UPDATE DOMAINEPIECE SET GDP_SOUCHE=""');
  ExecuteSQL('UPDATE AFCUMUL SET ACU_FAE=0,ACU_AAE=0,ACU_PCA=0');
  ExecuteSQL('UPDATE AFTABLEAUBORD SET ATB_FAE=0,ATB_AAE=0,ATB_PCA=0');
  ExecuteSQL('UPDATE PARPIECE SET GPP_NATPIECEANNUL=""');
  for i := 0 to 9 do InsertChoixCode('RTC', 'BL' + intTostr(i), 'Booléen libre ' + intTostr(i + 1), '', '');
  InsertChoixCode('RTC', 'BLA', 'Booléen libre 11', '', '');
  InsertChoixCode('RTC', 'BLB', 'Booléen libre 12', '', '');
  for i := 0 to 9 do InsertChoixCode('RTC', 'DL' + intTostr(i), 'Date libre ' + intTostr(i + 1), '', '');
  InsertChoixCode('RTC', 'DLA', 'Date libre 11', '', '');
  InsertChoixCode('RTC', 'DLB', 'Date libre 12', '', '');
  for i := 0 to 9 do InsertChoixCode('RTC', 'TL' + intTostr(i), 'Table libre ' + intTostr(i + 1), '', '');
  InsertChoixCode('RTC', 'TLA', 'Table libre 11', '', '');
  InsertChoixCode('RTC', 'TLB', 'Table libre 12', '', '');
  for i := 0 to 9 do InsertChoixCode('RTC', 'TX' + intTostr(i), 'Texte libre ' + intTostr(i + 1), '', '');
  InsertChoixCode('RTC', 'TXA', 'Texte libre 11', '', '');
  InsertChoixCode('RTC', 'TXB', 'Texte libre 12', '', '');
  for i := 0 to 9 do InsertChoixCode('RTC', 'VL' + intTostr(i), 'Valeur libre ' + intTostr(i + 1), '', '');
  InsertChoixCode('RTC', 'VLA', 'Valeur libre 11', '', '');
  InsertChoixCode('RTC', 'VLB', 'Valeur libre 12', '', '');
  CommitTrans;
end;

procedure MAJVER562;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  BeginTrans;
  ExecuteSQL('UPDATE ABSENCESALARIE SET PCN_FINDJ="PAM" WHERE PCN_FINDJ = "AMP"');
  ExecuteSQL('UPDATE ABSENCESALARIE SET PCN_DEBUTDJ="PAM" WHERE PCN_DEBUTDJ="AMP"');
  InstalleLesCoefficientsDegressifs;
  CommitTrans;
end;

procedure MAJVER564;
var SS: string;
begin
  SS := UsDateTime(IDate1900);
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  BeginTrans;
  ExecuteSQL('UPDATE PARPIECE SET GPP_MASQUERNATURE="-"');
  ExecuteSQL('UPDATE TIERSCOMPL SET YTC_FAMREG=""');
  CommitTrans;
end;

procedure MAJVER565;
var SQL, SS: string;
begin
  SS := UsDateTime(IDate1900);
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  BeginTrans;
  if Vers_TarifMode < 108 then // 09-11-2001
  begin
    SQL := 'UPDATE TARIFMODE SET GFM_COEF=GFM_COEFFICIENT';
    ExecuteSQL(SQL);
  end;
  if Vers_TarifTypMode < 106 then // 09-11-2001
  begin
    SQL := 'UPDATE TARIFTYPMODE SET GFT_COEF=GFT_COEFFICIENT';
    ExecuteSQL(SQL);
  end;
  CommitTrans;
end;

procedure MAJVER567;
var {SQL,} SS: string;
begin
  SS := UsDateTime(IDate1900);
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  BeginTrans;
  ExecuteSQL('UPDATE ARTICLE SET GA_ACTIVITEEFFECT="-" where GA_TYPEARTICLE<>"PRE"');
  ExecuteSQL('UPDATE ARTICLE SET GA_ACTIVITEEFFECT="X" where GA_TYPEARTICLE="PRE"');
  ExecuteSQL('UPDATE PARPIECE SET GPP_ECLATEAFFAIRE="-"');
  (* mcd 032003SQL:='UPDATE AFCUMUL SET ACU_PVPRODCON=0,ACU_PRPRODCON=0,ACU_PVFACTCON=0,ACU_FAECON=0,ACU_AAECON=0,ACU_PCACON=0 ';
  SQL:=SQL+',ACU_PVPRODDEV=0,ACU_PRPRODDEV=0,ACU_PVFACTDEV=0,ACU_FAEDEV=0,ACU_AAEDEV=0,ACU_PCADEV=0 ';
  ExecuteSQL(SQL);*)
  ExecuteSQL('UPDATE ACTIVITE SET ACT_ACTIVITEEFFECT="-",ACT_FOURNISSEUR="" where ACT_TYPEARTICLE<>"PRE"');
  ExecuteSQL('UPDATE ACTIVITE SET ACT_ACTIVITEEFFECT="X",ACT_FOURNISSEUR="" where ACT_TYPEARTICLE="PRE"');
  ExecuteSQL('UPDATE AFTABLEAUBORD SET ATB_FOURNISSEUR=""');
  CommitTrans;

  InsertChoixCode('ZLI', 'PT1', 'Table libre 1', 'Table libre 1', '');
  InsertChoixCode('ZLI', 'PT2', 'Table libre 2', 'Table libre 2', '');
  InsertChoixCode('ZLI', 'PT3', 'Table libre 3', 'Table libre 3', '');
  InsertChoixCode('ZLI', 'PD1', 'Date libre 1', 'Date libre 1', '');
  InsertChoixCode('ZLI', 'PD2', 'Date libre 2', 'Date libre 2', '');
  InsertChoixCode('ZLI', 'PD3', 'Date libre 3', 'Date libre 3', '');
  InsertChoixCode('ATU', '039', 'les tiers', '', 'les clients');
end;

procedure MAJVER568;
var SS: string;
begin
  SS := UsDateTime(IDate1900);
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  BeginTrans;
  ExecuteSQL('UPDATE BASEBRUTESPEC SET PBB_SOMISO=""');
  ExecuteSQL('UPDATE ETABCOMPL SET ETB_PRUDH="04",ETB_PRUDHCOLL="1",ETB_PRUDHSECT="4",ETB_PRUDHVOTE="1",ETB_PROFILREM="",ETB_BCMOISPAIEMENT="",ETB_BCMODEREGLE="",ETB_BCJOURPAIEMENT=1,ETB_CODESECTION="1"');
  ExecuteSQL('UPDATE HISTOSALARIE SET PHS_CONDEMPLOI="",PHS_PROFILREM="",PHS_BCONDEMPLOI="-",PHS_BPROFILREM="-"');
  ExecuteSQL('UPDATE REMUNERATION SET PRM_INDEMINTEMP="-"');
  ExecuteSQL('UPDATE SALARIES SET PSA_TYPDADSFRAC="ETB",PSA_ANAPERSO="-",PSA_TYPPROFILREM="ETB",PSA_PROFILREM="",PSA_REGIMESS="200",PSA_DADSFRACTION="1"');
  ExecuteSQL('UPDATE SALARIES SET PSA_TYPPRUDH="ETB" WHERE PSA_PRUDHCOLL="" AND PSA_PRUDHSECT="" AND PSA_PRUDHVOTE=""');
  ExecuteSQL('UPDATE SALARIES SET PSA_TYPPRUDH="PER" WHERE PSA_PRUDHCOLL<>"" OR PSA_PRUDHSECT<>"" OR PSA_PRUDHVOTE<>""');
  ExecuteSQL('UPDATE SALARIES SET PSA_CATDADS = "001" WHERE PSA_CATDADS="D"');
  ExecuteSQL('UPDATE SALARIES SET PSA_CATDADS = "002" WHERE PSA_CATDADS="C"');
  ExecuteSQL('UPDATE SALARIES SET PSA_CATDADS = "003" WHERE PSA_CATDADS="A"');
  ExecuteSQL('UPDATE PAIEENCOURS SET PPU_BULCOMPL="-",PPU_PROFILPART="",PPU_EDTDEBUT=PPU_DATEDEBUT,PPU_EDTFIN=PPU_DATEFIN');
  ExecuteSQL('UPDATE VARIABLEPAIE SET PVA_NBREMOISGLISS=0');
  ExecuteSQL('UPDATE HANDICAPE SET PGH_DATEFINCOT="' + SS + '"');
  ExecuteSql('UPDATE ABSENCESALARIE SET PCN_CODETAPE="..." WHERE PCN_CODETAPE=""');
  ExecuteSql('UPDATE COTISATION SET PCT_DADSEXOBASE="",PCT_DADSEXOCOT="" WHERE PCT_PREDEFINI<>"CEG"');
  CommitTrans;
  MajAnnulienLoiNRE;
  // AddNewNaturesGC ;   majver590
end;

procedure MAJVER569;
var i: integer;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  for i := 0 to 9 do InsertChoixCode('RSZ', 'CL' + intTostr(i), 'Table libre ' + intTostr(i + 1), '', '');
  for i := 0 to 2 do InsertChoixCode('RSZ', 'BL' + intTostr(i), 'Booléen libre ' + intTostr(i + 1), '', '');
  for i := 0 to 2 do InsertChoixCode('RSZ', 'DL' + intTostr(i), 'Date libre ' + intTostr(i + 1), '', '');
  for i := 0 to 2 do InsertChoixCode('RSZ', 'VL' + intTostr(i), 'Valeur libre ' + intTostr(i + 1), '', '');
  for i := 0 to 2 do InsertChoixCode('RSZ', 'TX' + intTostr(i), 'Texte libre ' + intTostr(i + 1), '', '');
end;

procedure MAJVER570;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  BeginTrans;
  ExecuteSQL('UPDATE CHAMPSPRO SET RCL_TYPETEXTE=""');
  CommitTrans;
  MajApresForTheTox;
end;

procedure MAJVER574;
var SS: string;
begin
  SS := UsDateTime(IDate1900);
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  BeginTrans;
  ExecuteSQL('UPDATE AFFAIRE SET AFF_GROUPECONF=""');
  //ExecuteSQL('UPDATE RESSOURCE SET ARS_WGENERIQUE="-",ARS_WRESSOURCELIE=""') ;
  ExecuteSQL('UPDATE JURIDIQUE set JUR_DATESTAT="' + SS + '"');
  CommitTrans;
end;

procedure MAJVER580;
var SQL: string;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  BeginTrans;
  ExecuteSQL('UPDATE CODECPTA SET GCP_CPTEGENESTOCK="", GCP_CPTEGENEVARSTK=""');
  SQL := 'UPDATE PARPIECE SET GPP_TYPEECRSTOCK="RIE", GPP_COMPSTOCKLIGNE="SAN", GPP_COMPSTOCKPIED="SAN"';
  SQL := SQL + ',GPP_TYPEPRESENT=0,GPP_RECUPPRE="PRE",GPP_TYPEPRESDOC="DEF",GPP_APPLICRG="-"';
  SQL := SQL + ',GPP_CONTEXTES="",GPP_NUMEROSERIE="-" ';
  ExecuteSQL(SQL);
  ExecuteSQL('UPDATE ARTICLE SET GA_REGROUPELIGNE="",GA_LISTEREGROUPE="",GA_PRINCIPALEXTRA=""');
  ExecuteSQL('UPDATE PIEDECHE SET GPE_NOFOLIO=0');
  CommitTrans;
  UpdateDecoupePiece('GP_REFCPTASTOCK="",GP_AFFAIREDEVIS=""');
  UpdateDecoupeLigne('GL_TYPEPRESENT=0, GL_ACTIONLIGNE="",GL_NOFOLIO=0');

end;

procedure MAJVER581(VSoc: Integer);
var SQL, SS: string;
  Q: TQuery;
  i: integer;
begin
  SS := UsDateTime(IDate1900);
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  BeginTrans;
  if Vers_TarifMode < 108 then ExecuteSQL('UPDATE TARIFMODE SET GFM_COEF=GFM_COEFFICIENT'); //09-11-2001
  if Vers_TarifMode < 109 then ExecuteSQL('UPDATE TARIFMODE SET GFM_NATURETYPE="VTE" where GFM_NATURETYPE<>"VTE" and GFM_NATURETYPE<>"ACH"'); //09-11-2001
  if Vers_TarifTypMode < 106 then ExecuteSQL('UPDATE TARIFTYPMODE SET GFT_COEF=GFT_COEFFICIENT'); //09-11-2001
  if Vers_TarifTypMode < 108 then ExecuteSQL('UPDATE TARIFTYPMODE SET GFT_NATURETYPE="VTE"');
  if Vers_ParPieceCompl < 105 then ExecuteSQL('UPDATE PARPIECECOMPL SET GPC_TYPEPASSCPTA="",GPC_TYPEPASSACC="",GPC_TYPEECRCPTA=""'); //05/12/01
  if Vers_ProfilArt < 111 then ExecuteSQL('UPDATE PROFILART SET GPF_ARRONDIPRIX="", GPF_ARRONDIPRIXTTC="", GPF_PRIXUNIQUE="X"'); // 09-01-2002
  // if Vers_ParCaisse     < 105 then ExecuteSQL('update PARCAISSE set GPK_REGFACSAISIE="-"') ;     //30-01-2002

  ExecuteSQL('UPDATE PARPIECE SET GPP_APERCUAVIMP="X" WHERE (GPP_APERCUAVIMP<>"-") OR (GPP_APERCUAVIMP IS NULL)');
  ExecuteSQL('UPDATE PARPIECE SET GPP_APERCUAVETIQ="X" WHERE (GPP_APERCUAVETIQ<>"-") OR (GPP_APERCUAVETIQ IS NULL)');
  ExecuteSQL('UPDATE PARPIECE SET GPP_VALMODELE="-" WHERE (GPP_VALMODELE<>"X") OR (GPP_VALMODELE IS NULL)');
  ExecuteSQL('UPDATE PARPIECECOMPL SET GPC_TYPEPASSCPTA="", GPC_TYPEPASSACC="", GPC_TYPEECRCPTA=""');
  ExecuteSQL('UPDATE PARPIECE SET GPP_MAJPRIXVALO="",GPP_ARTSTOCK="X" WHERE (GPP_NATUREPIECEG="INV") OR (GPP_NATUREPIECEG="TEM") OR (GPP_NATUREPIECEG="EEX") OR (GPP_NATUREPIECEG="SEX")');
  ExecuteSQL('UPDATE PARPIECE SET GPP_MAJPRIXVALO="PPA;",GPP_ARTSTOCK="X" WHERE GPP_NATUREPIECEG="TRE"');
  ExecuteSQL('UPDATE PARPIECE SET GPP_NBEXEMPLAIRE=1 WHERE GPP_NBEXEMPLAIRE=0');
  ExecuteSQL('UPDATE PARPIECE SET GPP_LIBELLE="Annonce de transfert" WHERE GPP_NATUREPIECEG="TRV"');
  ExecuteSQL('UPDATE PARPIECE SET GPP_REGROUPCPTA="AUC" WHERE GPP_REGROUPCPTA IS NULL');
  ExecuteSQL('UPDATE PARPIECE SET GPP_CPTCENTRAL="-" WHERE GPP_CPTCENTRAL IS NULL');

  if Vers_JoursEtab < 101 then ExecuteSQL('UPDATE JOURSETAB SET GJE_DATEMODIF=GJE_JOURNEE'); //15-01-2002
  if Vers_JoursCaisse < 102 then ExecuteSQL('UPDATE JOURSCAISSE SET GJC_DATEMODIF=GJC_DATEOUV,GJC_HEUREOUV="",GJC_HEUREFERME=""');
  if Vers_CtrlCaisMt < 101 then ExecuteSQL('UPDATE CTRLCAISMT SET GJM_DATEMODIF="' + SS + '"');

  // Modif pour saisie par codes barre
  ExecuteSQL('UPDATE LIGNE SET GL_TYPEDIM="NOR" WHERE GL_TYPEDIM="UNI"');
  ExecuteSQL('UPDATE LIGNE SET GL_REFARTSAISIE=GL_CODESDIM WHERE GL_TYPEDIM="GEN" AND GL_CODESDIM<>""');

  ExecuteSQL('UPDATE ARTICLE SET GA_INTERDITVENTE="-" WHERE GA_INTERDITVENTE<>"X"');
  SetParamSoc('SO_GCTOUTEURO', True);
  SetParamSoc('SO_GCFOACTIVECOMPTA', '-');

  ExecuteSQL('UPDATE PIECE SET GP_ETATEXPORT="ATT",GP_DATEEXPORT="' + SS + '" WHERE GP_ETATEXPORT<>"EXP"');
  ExecuteSQL('UPDATE PIECE SET GP_SUPPRIME="-" WHERE GP_SUPPRIME IS NULL');

  ExecuteSQL('UPDATE TIERS SET T_CLETELEPHONE=""');
  ExecuteSQL('UPDATE REPRISECHGPAO SET GRC_FORMULE="",GRC_GERETABLECORR="-"');

  //CHR
  ExecuteSQL('UPDATE ARTICLE SET GA_PRIXPASMODIF="-",GA_HRSTAT="",GA_QTEDEFAUT=""');
  //BTP
  ExecuteSQL('UPDATE TXCPTTVA SET TV_CPTVTERG="",TV_CPTACHRG=""');
  //Affaire
  ExecuteSQL('UPDATE AFFAIRE SET AFF_COEFREVALOCUM=1,AFF_MULTIECHE="-"');
  ExecuteSQL('Update factaff set afa_generauto=(select aff_generauto from affaire where aff_affaire=factaff.afa_affaire)  where afa_generauto="" ');
  ExecuteSQL('UPDATE EAFFAIRE SET EAF_DATEDEBEXER="' + SS + '",EAF_DATEFINEXER="' + SS + '"');
  ExecuteSQL('UPDATE ACTIVITE SET ACT_NUMPIECEACH=""');
  //GRC
  if vsoc <> 574 then ExecuteSQL('UPDATE SUSPECTS SET RSU_CONTACTTELEPH="",RSU_CONTACTRVA=""'); // protection pour ne pas reinitialiser la base SIC
  ExecuteSQL('UPDATE ACTIONS SET RAC_NUMCHAINAGE=0');
  RT_InsertLibelleInfoComplAction;
  //JURI
  ExecuteSQL('UPDATE ANNULIEN SET ANL_ACTDATE="' + SS + '" , ANL_ACTNBRE=0, ANL_ACTMONT=0, ANL_CONVLIB="-", ANL_CONVSUITE="-"');
  // Compta
  ExecuteSQL('UPDATE JOURNAL SET J_CHOIXDATE=""');
  ExecuteSQL('UPDATE JALREF SET JR_MULTIDEVISE="-" WHERE JR_NUMPLAN>20');
  //PAIE
  SQL := 'update histosalarie set PHS_HORAIREMOIS=PHS_PGHHORAIREMOIS,PHS_BHORAIREMOIS=PHS_BPGHHORMOIS,';
  SQL := SQL + 'PHS_BTAUXHORAIRE=PHS_BPGHTAUXHOR,PHS_BPERIODBUL=PHS_BPERIODEBUL,PHS_TTAUXPARTIEL=0,PHS_BTTAUXPARTIEL=PHS_BTAUXPARTIEL,';
  SQL := SQL +
    'PHS_SALAIRETHEO=0,PHS_BSALAIRETHEO="-",PHS_CHARLIBRE1="",PHS_BCHARLIBRE1="-",PHS_CHARLIBRE2="",PHS_BCHARLIBRE2="-",PHS_CHARLIBRE3="",PHS_BCHARLIBRE3="-",PHS_CHARLIBRE4="",PHS_BCHARLIBRE4="-"';
  ExecuteSQL(SQL);
  ExecuteSQL('update PAIEENCOURS SET PPU_REGLTMOD="-",PPU_CIVILITE=(SELECT PSA_CIVILITE FROM SALARIES WHERE PSA_SALARIE=PAIEENCOURS.PPU_SALARIE)');
  ExecuteSQL('update DUCSENTETE SET PDU_EMETTSOC=" "'); // PDU_DUCSDOSSIER="-", 26/9/2002 supprime à la demande de P.Dumet
  ExecuteSQL('update ORGANISMEPAIE SET POG_IDENTQUAL=" ",POG_IDENTEMET=" ",POG_IDENTDEST=" ",POG_ADHERCONTACT=" ",POG_PAIEGROUPE="-",POG_PAIEMODE=" ",POG_IDENTOPS=" ",POG_NOCONTEMET=" "');
  ExecuteSQL('update DUCSPARAM SET PDP_COTISQUAL=" "');
  SQL := 'update DEPORTSAL set PSE_BTP="-",PSE_ORGANIGRAMME="X",PSE_INTERMITTENT="-",PSE_MSA="-",PSE_ISNUMIDENT="",PSE_ISCONGSPEC="",';
  SQL := SQL +
    'PSE_ISRETRAITE="",PSE_ISCATEG="",PSE_ISNUMASSEDIC="",PSE_MSALIEUTRAV="",PSE_MSAINFOSCOMPL="",PSE_MSATYPEACT="",PSE_MSAACTIVITE="",PSE_BTPADHESION="",';
  SQL := SQL + 'PSE_BTPANCIENNETE="' + SS + '",PSE_BTPASSEDIC="",PSE_BTPSALMOYEN="",PSE_BTPHORAIRE="",PSE_EMAILPROF=""';
  ExecuteSQL(SQL);
  ExecuteSQL('update REMUNERATION set PRM_BTPSALAIRE="-",PRM_BTPARRET=""');
  ExecuteSQL('update ETABCOMPL set ETB_ISLICSPEC="-",ETB_ISNUMLIC="",ETB_ISOCCAS="-",ETB_ISLABELP="-",ETB_ISNUMLAB="",ETB_ISNUMCPAY="",ETB_ISNUMRECOUV="",ETB_MSAACTIVITE="",ETB_MSASECTEUR="",ETB_PROFILRET=""');
  ExecuteSQL('Update VISITEMEDTRAV Set PVM_MTENVOIMAIL="-",PVM_MTDATEMAIL="' + SS + '",PVM_EDITCONVOC="-",PVM_DATECONVOC="' + SS + '"');
  ExecuteSQL('Update SALARIES set PSA_PROFILRET="",PSA_TYPPROFILRET="ETB"');
  ExecuteSQL('UPDATE CONTRATTRAVAIL set PCI_ISCACHETS="-",PCI_ISHEURES="-",PCI_ISNBEFFECTUE=0');
  //FP
  ExecuteSQL('UPDATE FPBIENS SET FPB_N01265=0,FPB_N01266=0,FPB_N01267=0');

  CommitTrans;

  UpdateDecoupeLigne('GL_INDICESERIE=0, GL_NUMEROSERIE="-"');

  for i := 1 to 2 do InsertChoixCode('ZLI', 'AS' + intTostr(i), 'Statistique libre ' + intTostr(i), 'Stat ' + intTostr(i), '');
  InsertChoixCode('GAF', 'COM', 'Comptabilisé', 'Comptabilisé', '');
  InsertChoixCode('GAF', 'VIS', 'Visa', 'Visa', '');
  MajApresForTheTox2;

  if Vers_ParCaisse < 106 then
    ExecuteSQL('UPDATE PARCAISSE SET GPK_MODIFFDCAIS="-",GPK_REPRISEFDCAIS="-",GPK_CTRLPIECBIL="-",GPK_GEREREMISEBNQ="-",GPK_OPREMISEBNQ="",GPK_GERETICKETATT="-"');
  if Vers_CtrlCaisMt < 102 then ExecuteSQL('UPDATE CTRLCAISMT SET GJM_FDCAISOUV=0,GJM_PIECBILOUV="",GJM_PIECBILTOT=""');
  ExecuteSQL('UPDATE MODEPAIE SET MP_CLIOBLIGFO="-"');
  ExecuteSQL('UPDATE YMYBOBS SET YB_SOCREF=0');
  UpdateDecoupePiece('GP_ETATCOMPTA="ATT",GP_DATECOMPTA="' + SS + '"');

  Q := Opensql('Select * from CHOIXCOD where CC_TYPE="RSZ" and  CC_CODE like "TL%"', TRUE);
  if Q.Eof then
  begin
    Ferme(Q);
    for i := 0 to 2 do ExecuteSQL('Update CHOIXCOD set CC_CODE="TL' + intTostr(i) + '" where CC_TYPE="RSZ" and CC_CODE="TX' + intTostr(i) + '"');
  end else
    Ferme(Q);
end;

procedure MAJVER589;
var Creat, Usr: string;

  procedure MajChpsSyst(Table, Racine: string);
  begin
    ExecuteSQL('update ' + Table + ' set ' + Racine + '_DATECREATION="' + UsTime(NowH) + '", '
      + Racine + '_DATEMODIF  = "' + UsDateTime(Encodedate(1900, 01, 02)) + '", '
      + Racine + '_CREATEUR   ="' + Creat + '", '
      + Racine + '_UTILISATEUR="' + Usr + '"');
  end;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //Maj des dates et usr
  Creat := 'CEG';
  Usr := 'CEG';
  MajChpsSyst('ADRESSES', 'ADR');
  MajChpsSyst('TIERSCOMPL', 'YTC');
  MajChpsSyst('CONTACT', 'C');
  MajChpsSyst('TIERSPIECE', 'GTP');
  MajChpsSyst('RIB', 'R');
  MajChpsSyst('LIENSOLE', 'LO');
  MajChpsSyst('ARTICLELIE', 'GAL');
  MajChpsSyst('ARTICLEPIECE', 'GAP');
  MajChpsSyst('CATALOGU', 'GCA');
  MajChpsSyst('CONDITIONNEMENT', 'GCO');
  MajChpsSyst('ARTICLETIERS', 'GAT');
  MajChpsSyst('GTRADARTICLE', 'GTA');
  MajChpsSyst('NOMENLIG', 'GNL');

  // MODE
  BeginTrans;
  // voir MAJVER594 AddNewNaturesGC ;             // Mise à jour des natures de PARPIECE - les supprimer dans MAJAVANT
  InsertChoixCode('GTM', 'ALF', 'Annonce de livraison', 'Annonce livraison', '');
  InsertChoixCode('ZLI', 'AS0', 'Collection', 'Collection', '');
  //ExecuteSQL ('delete from graphs where (gr_graphe="TGCGRAPHVIEPROD" or gr_graphe="TGCGRAPHSYNM" or gr_graphe="TGCGRAPHSYNA")');
  ExecuteSQL('update DIMMASQUE set GDM_FERMER="-" where (GDM_FERMER<>"X") OR (GDM_FERMER IS NULL)');
  ExecuteSQL('update ARTICLE set GA_FAMILLETAXE1="" where GA_TYPEARTICLE="FI"');
  ExecuteSQL('update PROPTRANSFLIG set GTL_STOCKALF=0');
  ExecuteSQL('update PARPIECE set GPP_CPTCENTRAL="-" where GPP_CPTCENTRAL<>"X"');
  ExecuteSql('update PARPIECE set GPP_TRSFACHAT="-", GPP_TRSFVENTE="-"');
  ExecuteSQL('update MODEPAIE set MP_CONDITION="-" where MP_CONDITION<>"X"');
  SetParamSoc('SO_GCCREERTARIFBASE', '-');
  SetParamSoc('SO_GCPROFILART', '');
  SetParamSoc('SO_GCAXEANALYTIQUE', '-');
  SetParamSoc('SO_GCTOXEPURE', 7);

  ExecuteSQL('update STOXSITES set SSI_CURRENTSITE="X" where SSI_STYPESITE="002"');
  ExecuteSQL('update STOXSITES set SSI_CDEPARTFTP=SSI_CDEPART where SSI_CDEPARTFTP=""');
  ExecuteSQL('update STOXSITES set SSI_CREJETFTP=SSI_CREJET where SSI_CREJETFTP=""');
  ExecuteSQL('update STOXSITES set SSI_CTRAITEFTP=SSI_CTRAITE where SSI_CTRAITEFTP=""');
  ExecuteSQL('update STOXSITES set SSI_CENVOYEFTP=SSI_CENVOYE where SSI_CENVOYEFTP=""');
  ExecuteSQL('update STOXSITES set SSI_CARRIVEFTP=SSI_CARRIVE where SSI_CARRIVEFTP=""');
  CommitTrans;

  // MODE-PCP
  ExecuteSQL('update COMMERCIAL set GCL_PREFIXETIERS="" where GCL_PREFIXETIERS is null');
end;

procedure MAJVER590;
var SS, SQL: string;
  Creat, Usr: string;

  procedure MajChpsSyst(Table, Racine: string);
  begin
    ExecuteSQL('update ' + Table + ' set ' + Racine + '_DATECREATION="' + UsTime(NowH) + '", '
      + Racine + '_DATEMODIF  ="'+UsDateTime(Encodedate(1900,01,02))+'", '
      + Racine + '_CREATEUR   ="' + Creat + '", '
      + Racine + '_UTILISATEUR="' + Usr + '"');
  end;
begin
  SS := UsDateTime(IDate1900);
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  Creat := 'CEG';
  Usr := 'CEG';

  //CHR
  ExecuteSql('update HRDOSSIER set HDC_HRTARIF=""');
  ExecuteSql('update HRETAT set HES_HRNIVEAUTARIF=0');
  ExecuteSql('update HRPREFACT set HPF_HRTARIF="", HPF_QTESIT=0');
  ExecuteSql('update TARIF set GF_TYPRES="", GF_REGROUPELIGNE="", GF_HRPREFACT="-", GF_NOFOLIO=0');
  //affaire
  if Vers_Tache < 107 then
  begin
    BeginTrans;
    ExecuteSQL('UPDATE ACTIVITE SET ACT_NUMEROLIGNE=0,ACT_ACTORIGINE="",ACT_NUMEROTACHE=0');

    if not IsDossierPCL then
    begin
      ExecuteSQL('UPDATE EACTIVITE SET EAC_ACTORIGINE=""');
    end;  // not IsDossierPCL
    InsertChoixCode('ZLI', 'TM1', 'Montant libre 1', 'Montant libre 1', '');
    InsertChoixCode('ZLI', 'TM2', 'Montant libre 2', 'Montant libre 2', '');
    InsertChoixCode('ZLI', 'TM3', 'Montant libre 3', 'Montant libre 3', '');
    InsertChoixCode('ZLI', 'TT4', 'Table libre 4', 'Table libre 4', '');
    InsertChoixCode('ZLI', 'TT5', 'Table libre 5', 'Table libre 5', '');
    InsertChoixCode('ZLI', 'TT6', 'Table libre 6', 'Table libre 6', '');
    InsertChoixCode('ZLI', 'TT7', 'Table libre 7', 'Table libre 7', '');
    InsertChoixCode('ZLI', 'TT8', 'Table libre 8', 'Table libre 8', '');
    InsertChoixCode('ZLI', 'TT9', 'Table libre 9', 'Table libre 9', '');
    InsertChoixCode('ZLI', 'TTA', 'Table libre 10', 'Table libre 10', '');

    //tablette AFTACTIVITEORIGINE
    InsertChoixCode('AAO', 'ACH', 'Achats', '', '');
    InsertChoixCode('AAO', 'BL', 'Bon de livraison', '', '');
    InsertChoixCode('AAO', 'EAC', 'E-activité', '', '');
    InsertChoixCode('AAO', 'PLA', 'Planning', '', '');
    InsertChoixCode('AAO', 'REP', 'Reprise', '', '');
    InsertChoixCode('AAO', 'SAI', 'Saisie', '', '');
    InsertChoixCode('AAO', 'SDE', 'Saisie décentralisée', '', '');

    //tablette AFTETAT
    // InsertChoixCode ('AFE','COM','Complet','','');
    // InsertChoixCode ('AFE','ENC','En Cours','','');
    // InsertChoixCode ('AFE','RIE','Rien','','');
    // InsertChoixCode ('AFE','TER','Terminé','','');
    CommitTrans;
    BeginTrans;
    ExecuteSQL('UPDATE AFFAIRE SET AFF_DATECUTOFF="' + SS + '"');
    // mcd 644 ExecuteSQL('UPDATE TACHE SET ATA_MODESAISIEPDC = "QUA"');
    CommitTrans;
  end;
  InsertChoixCode('DET', 'EXC', 'Document Excel', 'Excel', '');
  InsertChoixCode('DET', 'PPT', 'Document PowerPoint', 'Power', '');
  InsertChoixCode('DET', 'WOR', 'Document Word', 'WOrd', '');
  InsertChoixCode('DEN', 'LMI', 'Lettre de mission', 'LM', '');
  InsertChoixCode('DEN', 'PMI', 'Proposition de mission', 'Propo', '');

  ExecuteSQL('UPDATE RESSOURCE SET ARS_GENERIQUE="-",ARS_RESSOURCELIE="",ARS_TAUXSIMUL=0,ARS_SECTIONPDR="",ARS_RUBRIQUEPDR=""');
  ExecuteSQL('UPDATE AFBUDGET SET ABU_PERIODE="",ABU_SEMAINE=""');
  ExecuteSQL('UPDATE AFTABLEAUBORD SET ATB_PRRESS=0,ATB_PVRESS=0,ATB_PRART=0,ATB_PVART=0,ATB_ENGAGEPV=0,ATB_ENGAGEPR=0,ATB_ENGAGEQTE=0,ATB_FAR=0,ATB_AAR=0,ATB_CCA=0');
  {$IFDEF CEGID}
  ExecuteSQL('UPDATE AFFAIRE SET AFF_DATFACECLATMOD="' + SS + '"');
  {$ELSE}
  ExecuteSQL('UPDATE AFFAIRE SET AFF_DATFACECLATMOD="' + SS + '",AFF_REGSURCAF="-"');
  {$ENDIF}

  // fin affaire

  // GRC
  MajChpsSyst('PROSPECTS', 'RPR');
  MajChpsSyst('PROSPECTCOMPL', 'RPC');
  MajChpsSyst('RTINFOSDATA', 'RDA');
  {$IFNDEF CEGID}
  ExecuteSql('update PERSPECTIVES set RPE_COMMENTPERTE="",RPE_CREATEUR=RPE_UTILISATEUR');
  ExecuteSql('update PERSPHISTO set RPH_COMMENTPERTE="",RPH_CREATEUR=RPH_UTILISATEUR');
  ExecuteSql('update PARACTIONS set RPA_GESTRAPPEL="-", RPA_DELAIRAPPEL="",RPA_OUTLOOK="-",RPA_ENVOIMAIL="-",RPA_MAILRESP="-"');
  ExecuteSql('update ACTIONS set RAC_GESTRAPPEL="-", RAC_DELAIRAPPEL="",RAC_DATERAPPEL="' + SS + '",RAC_CREATEUR=RAC_UTILISATEUR');
  {$ENDIF}
  ExecuteSql('update actions set RAC_DESTMAIL = ";"||RAC_DESTMAIL where RAC_DESTMAIL <> "" and RAC_DESTMAIL not like ";%"');
  ExecuteSql('update ACTIONSGENERIQUES set RAG_CREATEUR="' + Creat + '",RAG_UTILISATEUR="' + Usr + '"');
  ExecuteSql('update SUSPECTS set RSU_CREATEUR="' + Creat + '",RSU_UTILISATEUR="' + Usr + '"');
  ExecuteSql('update OPERATIONS set ROP_CREATEUR=ROP_UTILISATEUR');
  ExecuteSql('update PROJETS set RPJ_CREATEUR=RPJ_UTILISATEUR');
  MajChpsSyst('SUSPECTSCOMPL', 'RSC');
  {$IFNDEF CEGID}
  SQL := 'update ACTIONSCHAINEES set RCH_DATEDEBUT="' + SS + '",RCH_DATEFIN="' + SS + '",RCH_CREATEUR="' + Creat;
  SQL := SQL + '",RCH_UTILISATEUR="' + Usr + '",RCH_TERMINE="-",RCH_INTERVENANT="",RCH_DATECREATION="' + UsTime(NowH);
  SQL := SQL + '",RCH_DATEMODIF="'+UsDateTime(Encodedate(1900,01,02))+'",RCH_AUXILIAIRE=(select T_Auxiliaire from Tiers where T_Tiers=ACTIONSCHAINEES.RCH_TIERS)';
  ExecuteSQL(SQL);
  {$ENDIF}
  ExecuteSql('update RTINFOSDESC set RDE_FILTRE=""');
  ExecuteSql('update CHAMPSPRO set RCL_FILTRE=""');
  ExecuteSql('update LISTEINVENT set GIE_CONTREMARQUE="-"');

  // COMPTA
  ExecuteSql('update RELANCE set RR_ENJOURS="-"');
  ExecuteSql('update GENERAUX set G_COMPENS="-"');
  ExecuteSQL('UPDATE GUIDE SET GU_TRESORERIE="-"');
  ExecuteSQL('UPDATE GUIDEREF SET GDR_TRESORERIE="-"');
  ExecuteSQL('UPDATE JOURNAL SET J_CHOIXDATE="DATEOPERATION",J_INCREF="-", J_NATDEFAUT="",J_NATCOMPL="-",J_EQAUTO="-",J_INCNUM="-"');
  ExecuteSQL('UPDATE JALREF SET JR_CHOIXDATE="DATEOPERATION",JR_INCREF="-", JR_NATDEFAUT="",JR_NATCOMPL="-",JR_EQAUTO="-",JR_INCNUM="-"');
  ExecuteSQL('UPDATE IMMO SET I_SURAMORT="-"');

  ExecuteSql('UPDATE ECRGUI SET EG_DEBITDEV="" WHERE EG_DEBITDEV IS NULL ');
  ExecuteSql('UPDATE ECRGUI SET EG_CREDITDEV="" WHERE EG_CREDITDEV IS NULL ');
  // ExecuteSQL('UPDATE BANQUECP SET BQ_EBAN=""') ;

  //PAIE

  ExecuteSQL('Update CONTRATTRAVAIL set PCI_DEBPREAVIS="' + SS + '",PCI_FINPREAVIS="' + SS + '",PCI_MOTIFCTINT="",PCI_SALAIREMOIS5=0,PCI_SALAIRANN5=0 ');
  ExecuteSQL('Update EXERSOCIAL set PEX_DEBUTSAISVAR="' + SS + '",PEX_FINSAISVAR="' + SS + '"');
  ExecuteSQL('Update DROITACCES set YDA_MASQUEVAR=""');
  ExecuteSQL('Update SALARIES set PSA_ANCIENPOSTE=PSA_DATEANCIENNETE,PSA_CATBILAN="",PSA_EDITORG="",PSA_MOTIFSUSPPAIE="",PSA_MVALOMS="",PSA_SALAIREMOIS5=0,PSA_SALAIRANN5=0,PSA_TYPEDITORG="ETB",PSA_TYPNBACQUISCP="ETB",PSA_VALODXMN=0,PSA_CPTYPEVALO="ETB"');
  ExecuteSQL('Update HISTOSALARIE set PHS_BSALAIREANN5="-",PHS_BSALAIREMOIS5="-",PHS_DATERETRO=PHS_DATEEVENEMENT,PHS_SALAIREMOIS5=0,PHS_SALAIREANN5=0');
  ExecuteSQL('UPDATE MOTIFABSENCE SET PMA_TYPERTT="-", PMA_OKSAISIERESP="-" , PMA_CALENDCIVIL="-"');
  ExecuteSQL('UPDATE ABSENCESALARIE SET PCN_DATECREATION="' + SS + '", PCN_VALIDABSENCE="", PCN_OKFRACTION="OUI"');
  ExecuteSQL('UPDATE DEPORTSAL SET PSE_EMAILENVOYE="-", PSE_EMAILDATE="' + SS + '"');
  ExecuteSQL('UPDATE RECAPSALARIES SET PRS_PRIVALIDE=0, PRS_PRIATTENTE=0, PRS_CUMRTTREST=0');
  ExecuteSQL('update REMUNERATION set PRM_INDEMCOMPCP="-", PRM_INDEMCONVENT="-", PRM_SOMISOL=""');
  ExecuteSQL('update ATTESTATIONS set PAS_INDCONV2="0", PAS_INDTRANS2="0"');
  ExecuteSQL('update EMETTEURSOCIAL set PET_PAYS="FRA"');
  ExecuteSQL('update ETABCOMPL SET ETB_RIBDUCSEDI="",ETB_TYPNBACQUISCP="VAL",ETB_EDITORG="001"');
  ExecuteSQL('update ORGANISMEPAIE SET POG_RIBDUCSEDI="",POG_CODAPPLI="013",POG_SERVUNIQ="-",POG_LGOPTIQUE=""');
  ExecuteSQL('update ENVOISOCIAL SET PES_CODAPPLI="013",PES_SERVUNIQ="-"  WHERE PES_TYPEMESSAGE="DUC"');
  ExecuteSQL('update ENVOISOCIAL SET PES_CODAPPLI="",PES_SERVUNIQ="" WHERE PES_TYPEMESSAGE<>"DUC"');
  ExecuteSQL('update ENVOISOCIAL SET PES_MTPAYER=0,PES_NEANT="-",PES_PAIEMODE=""');
  ExecuteSQL('update DUCSENTETE SET PDU_DUCSDOSSIER="-" WHERE PDU_DUCSDOSSIER=""');
  ExecuteSQL('update absencesalarie set pcn_codetape="..." where pcn_typemvt="CPA" and pcn_typeconge="PRI" and (pcn_codetape is null or pcn_codetape="")');
  ExecuteSQL('update absencesalarie set pcn_validresp="VAL" where pcn_typemvt="CPA" and pcn_typeconge="PRI" and (pcn_validresp is null or pcn_validresp="")');

  // GPAO
  BeginTrans;
  ExecuteSql('UPDATE DEVISE SET D_PARITEEUROFIXING = D_PARITEEURO');
  ExecuteSql('UPDATE PARPIECE SET GPP_PIECEPILOTE="-"');
  SQL := 'UPDATE ARTICLE SET GA_QECOACH = 0,GA_QPCBACH = 0,GA_DELAIACH = 0,GA_METHACH = "CDE",GA_PRIXSIMULACH = 0';
//  SQL := SQL + ' ,GA_UNITEQTEACH = GA_QUALIFUNITESTO,GA_QECOPROD = 0,GA_QPCBPROD = 0,GA_DELAIPROD = 0';  DELAIPROD supprimé
  SQL := SQL + ' ,GA_UNITEQTEACH = GA_QUALIFUNITESTO,GA_QECOPROD = 0,GA_QPCBPROD = 0';
  SQL := SQL + ' ,GA_METHPROD = "CDE",GA_UNITEPROD = GA_QUALIFUNITESTO,GA_COEFPROD = 1,GA_QECOVTE = 0';
  SQL := SQL + ' ,GA_UNITEQTEVTE = GA_QUALIFUNITESTO,GA_CODEFORME = "",GA_TYPECOMPOSANT = "AUTRE",GA_VALOQTEECOPDR = "-"';
  SQL := SQL + ' ,GA_SECTIONPDR = "",GA_RUBRIQUEPDR = "",GA_UNITECONSO = GA_QUALIFUNITESTO,GA_MODECONSO = "LAN"';
  SQL := SQL + ' ,GA_PERTEPROP = 0,GA_FANTOME = "-",GA_PROFILARTICLE = "",GA_ESTPROFIL = "-"';
  ExecuteSQL(SQL);
  SQL := 'UPDATE PROFILART SET GPF_TYPEPROFIL = "SIM",GPF_MODEREMPLACE = "",GPF_ARTICLE = "",GPF_MODECOPIE = ""';
  SQL := SQL + ' ,GPF_LISTECHAMPS1 = "",GPF_LISTECHAMPS2 = "",GPF_LISTECHAMPS3 = "",GPF_LISTECHAMPS4 = "",GPF_LISTECHAMPS5 = ""';
//  SQL := SQL + ' ,GPF_MAJFICHEARTICLE = "",GPF_CODEFORME = "",GPF_QECOPROD = 0,GPF_QPCBPROD = 0,GPF_DELAIPROD = 0';
// DELAIPROD supprimé
  SQL := SQL + ' ,GPF_MAJFICHEARTICLE = "",GPF_CODEFORME = "",GPF_QECOPROD = 0,GPF_QPCBPROD = 0';
  SQL := SQL + ' ,GPF_METHPROD = "",GPF_UNITECONSO = "",GPF_MODECONSO = "",GPF_PERTEPROP = 0,GPF_UNITEPROD = ""';
  ExecuteSQL(SQL);
  CommitTrans;

  //JURI
  ExecuteSQL('update ANNULIEN set ANL_RACINE="ASS" where ANL_FONCTION="MGI" ');
  ExecuteSQL('update ANNULIEN set ANL_RACINE="ADM" where ANL_FONCTION="AGI" ');
  ExecuteSQL('update ANNULIEN set ANL_RACINE="PCA" where ANL_FONCTION="PGI" ');
  CorrectionSiren;

  // GC
  // voir MAJVER594 AddNewNaturesGC ;
  UpdateDecoupePiece('GP_CODEORDRE = 0');
  //  champs supprimes en 606
  // UpdateDecoupePiece('GP_NATURETRAVAIL = "",GP_CODEORDRE = 0');
  //UpdateDecoupeLigne('GL_NATURETRAVAIL = "",GL_CODEORDRE = 0,GL_LIGNEORDRE = 0');
end;

procedure MAJVER594;
var Creat, stInsertDepot: string;
  SS, SQL, StH, St: string;
  i: integer;
  HH: TDateTime;
  (*
  procedure UpdateNUM2(Nomtable, Prefixe: string);
  begin
    ExecuteSQL('UPDATE ' + NomTable + ' SET ' + Prefixe + '_NUM2=0 ');
  end;
  *)
begin
  SS := UsDateTime(IDate1900);
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //GC
  ExecuteSql('UPDATE ARTICLE SET GA_PRIXPOURQTEAC=1');
  ExecuteSql('UPDATE CATALOGU SET GCA_PRIXPOURQTEAC=1,GCA_CODEBARRE="",GCA_QUALIFCODEBARRE=""');

  // GRC
  SQL := 'update ACTIONSCHAINEES set RCH_TABLELIBRECH1="",RCH_TABLELIBRECH2="",RCH_TABLELIBRECH3=""';
  SQL := SQL + ',RCH_GESTRAPPEL="-",RCH_DELAIRAPPEL="",RCH_DATERAPPEL="' + SS + '"';
  ExecuteSQL(SQL);
  RTMoveConcurrent;
  RT_InsertLibelleInfoComplOperationetChain;
  RTMoveInfosData;

  // GA-GI
  ExecuteSql('update aftableaubord set atb_ecart3=0,atb_ecart4=0,atb_ecart5=0,atb_rafpr=0,atb_rafpv=0,atb_rafqte=0,atb_libelleart=""');
  ExecuteSql('update afplanning set apl_lignegeneree="-",apl_qteplanifuref=0,apl_qterealuref=0');
  ExecuteSql('update activite set act_qteuniteref=0');
  SQL := 'update tacheressource set ';  //mcd 644
  SQL := SQL + 'atr_qteinituref=0,atr_statutres="ACT"';   //mcd 644
  ExecuteSql(SQL);
  SQL := 'update tache set ata_modegene=1,ata_qteintervent=0,ata_dateannuelle="' + SS + '",ata_jourinterval=1';
  SQL := SQL +
    ',ata_semaineinterv=1,ata_moismethode=1,ata_moisjourfixe=1,ata_moisfixe=1,ata_moissemaine=1,ata_moisjourlib="LUN",ata_gestiondatefin="X",ata_qteinituref=0';
  SQL := SQL + ',  ata_lastdategene="' + SS +  '"';      //mcd 644
    
  ExecuteSql(SQL);
  ExecuteSql('update Ressource set ars_esthumain="X" where ars_typeressource in ("SAL","ST","INT")');
  ExecuteSql('update Ressource set ars_esthumain="-" where ars_typeressource not in ("SAL","ST","INT")');
  InsertChoixCode('DEN', 'ESO', 'Etat de synthèse OLE', 'OLE', '');
  GIGAChangeUnite;
  GiGaBasculeAna(False);
  //BTP

  ExecuteSQL('UPDATE AFFAIRE SET AFF_ACOMPTE=0,AFF_ACOMPTEREND=0');
  ExecuteSQL('UPDATE BTPARDOC SET BPD_SAUTAPRTXTDEB="-",BPD_SAUTAVTTXTFIN="-",BPD_IMPRECPAR="-",BPD_IMPRECSIT="-",BPD_IMPDESCRIPTIF="I"');
  ExecuteSQL('UPDATE BSITUATIONS SET BST_MONTANTACOMPTE=0,BST_MONTANTREGL=0');
  ExecuteSQL('UPDATE LIGNEOUV SET BLO_REFARTSAISIE=BLO_CODEARTICLE ');

  // CHR
  ExecuteSql('UPDATE HRDOSRES set HDR_DOSSIERALLOT="NON",HDR_NBRESALLOT=0,HDR_NBRESALLOTATT=0,HDR_NBRESALLOTRET=0,HDR_EXCEPTION="-"');
  ExecuteSql('UPDATE HRDOSSIER set HDC_DOSSIERALLOT="NON", HDC_HRCONTRAT=""');
  ExecuteSQL('UPDATE HREVENEMENT SET HEV_HRTARIF=""');
  ExecuteSQL('UPDATE HRNBPERSONNE SET HNP_ABREGECOURT=""');
  ExecuteSQL('UPDATE HRPARAMPLANNING SET HPP_TAILLECOLENT2=0,HPP_TAILLECOLENT3=0,HPP_COULJFERIE=HPP_COULDIMANCHE, HPP_HAUTLIGNEENT=18, HPP_HAUTLIGNEDATA=18, HPP_VISUTYPERES="-",'
    +
    'HPP_VISULIGNETO="-", HPP_DUREETIMER=0');
  ExecuteSql('UPDATE TARIF SET GF_HRGCOM="-", GF_HRGNUITGRAT="-", GF_NIVEAUYIELD=0');
  HH := 0;
  StH := USTIME(HH);
  ExecuteSQL('UPDATE PIEDECHE SET GPE_CREATEUR="' + Creat + '", GPE_DATECREATION="' + SS + '", GPE_HEURECREATION="' + StH + '"');

  // MODE
  ExecuteSQL('delete from etabliss where et_etablissement is null');
  ExecuteSQL('delete from etabliss where et_etablissement=""');
  if Vers_Etabliss < 20 then ExecuteSQL('update etabliss set et_depotlie=""');
  if Vers_Depots < 103 then
  begin
    if (GetParamSoc('SO_GCMULTIDEPOTS') = False) then
    begin
      // suppression des dépôts ayant le même code qu'un établissement
      ExecuteSQL('delete from depots where gde_depot in (select et_etablissement from etabliss)');
      stInsertDepot := 'insert into depots (gde_depot,gde_libelle,gde_abrege,gde_adresse1, gde_adresse2,gde_adresse3,gde_codepostal, gde_ville, gde_pays,' +
        'gde_telephone, gde_fax, gde_email, gde_utilisateur, gde_datecreation, gde_datemodif, gde_sursite, gde_sursitedistant,' +
        'gde_libredep1, gde_libredep2, gde_libredep3, gde_libredep4, gde_libredep5, gde_libredep6, gde_libredep7, gde_libredep8, gde_libredep9, ' +
        'gde_libredepA, gde_datelibre1, gde_datelibre2, gde_datelibre3, gde_boollibre1, gde_boollibre2, gde_boollibre3, gde_charlibre1,' +
        'gde_charlibre2, gde_charlibre3, gde_vallibre1,gde_vallibre2, gde_vallibre3) select et_etablissement as gde_depot,et_libelle as  ' +
        'gde_libelle,et_abrege as gde_abrege, et_adresse1 as gde_adresse1, et_adresse2 as gde_adresse2, et_adresse3 as ' +
        'gde_adresse3,et_codepostal as gde_codepostal,et_ville as gde_ville,et_pays as gde_pays,et_telephone as gde_telephone,et_fax as ' +
        'gde_fax,et_email as gde_email,et_utilisateur as gde_utilisateur,et_datecreation as ' +
        'gde_datecreation,et_datemodif as gde_datemodif,et_sursite as gde_sursite,et_sursitedistant as gde_sursitedistant,et_libreet1 as ' +
        'gde_libredep1,et_libreet2 as gde_libredep2,et_libreet3 as gde_libredep3,et_libreet4 as gde_libredep4,et_libreet5 as ' +
        'gde_libredep5,et_libreet6 as gde_libredep6,et_libreet7 as gde_libredep7,et_libreet8 as gde_libredep8,et_libreet9 as ' +
        'gde_libredep9,et_libreetA as gde_libredepA,et_datelibre1 as gde_datelibre1,et_datelibre2 as gde_datelibre2,et_datelibre3 as ' +
        'gde_datelibre3,et_boollibre1 as gde_boollibre1,et_boollibre2 as gde_boollibre2,et_boollibre3 as gde_boollibre3,et_charlibre1 as ' +
        'gde_charlibre1,et_charlibre2 as gde_charlibre2,et_charlibre3 as gde_charlibre3,et_vallibre1 as gde_vallibre1,et_vallibre2 as ' +
        'gde_vallibre2,et_vallibre3 as gde_vallibre3 ' +
        'from etabliss';
      ExecuteSQL(stInsertDepot);
      // suppression de toutes les lignes des titres des zones libres du dépôt et insertion à la place la valeur des titres des zones libres de l'établissement
      ExecuteSQL('delete from CHOIXCOD where CC_TYPE="ZLI" and CC_CODE like "D%"');
      ExecuteSQL('insert into CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) select "ZLI" as CC_TYPE,"D"||right(CC_CODE,2) as CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE from CHOIXCOD where CC_TYPE="ZLI" and CC_CODE like "E%"');
      // suppression de toutes les lignes des tables libres du dépôt et insertion à la place la valeur des tables libres de l'établissement
      ExecuteSQL('delete from CHOIXEXT where YX_TYPE like "D%"');
      ExecuteSQL('insert into CHOIXEXT (YX_TYPE,YX_CODE,YX_LIBELLE,YX_ABREGE,YX_LIBRE) select "D"||right(YX_TYPE,2) as YX_TYPE,YX_CODE,YX_LIBELLE,YX_ABREGE,YX_LIBRE from CHOIXEXT where YX_TYPE like "E%"');
    end
    else
    begin
      ExecuteSQL('update depots set gde_libredep1="", gde_libredep2="", gde_libredep3="", gde_libredep4="", gde_libredep5="", gde_libredep6="", gde_libredep7="", gde_libredep8="", gde_libredep9="", '
        +
        'gde_libredepA="", gde_datelibre1="' + SS + '", gde_datelibre2="' + SS + '", gde_datelibre3="' + SS +
          '", gde_boollibre1="-", gde_boollibre2="-", gde_boollibre3="-", gde_charlibre1="",' +
        'gde_charlibre2="", gde_charlibre3="", gde_vallibre1=0,gde_vallibre2=0, gde_vallibre3=0');
    end;
  end;

  if Vers_ParCaisse < 108 then
    ExecuteSQL('UPDATE PARCAISSE SET GPK_IMPMODSTAFAM="FAM",GPK_IMPMODSTAREM="REM",GPK_IMPMODLISART="ART",GPK_AFFICHEUR="X",GPK_OUVTIROIRFIN="-"');
  if Vers_ParCaisse < 109 then ExecuteSQL('UPDATE PARCAISSE SET GPK_CLIREPRISE="X",GPK_SEUILCLIOBLIG=0,GPK_AFFPRIXTIC="002",GPK_TOOLBAR="-"');
  if Vers_Jourscaisse < 103 then
    ExecuteSQL('UPDATE JOURSCAISSE SET GJC_NBTICABANDON=0,GJC_NBTICANNUL=0,GJC_NBTICMISATT=0,GJC_NBTICATTREPRI=0,GJC_NBTICATTSUP=0,GJC_NBLIGANNUL=0,GJC_NBMODIFMDP=0,GJC_NBOUVTIROIR=0,GJC_NBRATCLI=0');

  ExecuteSQL('UPDATE PARPIECE SET GPP_TYPECOMMERCIAL="REP",GPP_FILTRECOMM="" WHERE GPP_NATUREPIECEG NOT IN ("FFO","FFA")');
  ExecuteSQL('UPDATE PARPIECE SET GPP_TYPECOMMERCIAL="VEN",GPP_FILTRECOMM="001" WHERE GPP_NATUREPIECEG IN ("FFO","FFA")');
  ExecuteSQL('UPDATE PARPIECE SET GPP_MAJINFOTIERS="X" WHERE GPP_NATUREPIECEG NOT IN ("EEX","SEX","TEM","TRE","TRV","FFO","FFA")');
  ExecuteSQL('UPDATE PARPIECE SET GPP_MAJINFOTIERS="-" WHERE GPP_NATUREPIECEG IN ("EEX","SEX","TEM","TRE","TRV","FFO","FFA")');

  ExecuteSQL('UPDATE PIECE SET GP_TICKETANNULE="-" WHERE GP_TICKETANNULE <> "X"');
  if Vers_ArticleCompl < 101 then ExecuteSQL('UPDATE ARTICLECOMPL SET GA2_RATTACHECLI="" where GA2_RATTACHECLI is null');
  // MODE Multi-dépôts dans le paramétrage des dimensions
  ExecuteSQL('UPDATE USERPREFDIM SET GUD_CHAMP="GDE_DEPOT" WHERE GUD_CHAMP="ET_ETABLISSEMENT"');
  ExecuteSQL('UPDATE USERPREFDIM SET GUD_CHAMP="GDE_LIBELLE" WHERE GUD_CHAMP="ET_LIBELLE"');
  ExecuteSQL('UPDATE USERPREFDIM SET GUD_CHAMP="GDE_ABREGE" WHERE GUD_CHAMP="ET_ABREGE"');

  // MODE-PCP
  ExecuteSQL('update COMMERCIAL set GCL_PREFIXETIERS="" where GCL_PREFIXETIERS is null');

  (*
  // GP : Remise à Niveau Structure IMMO
  BeginTrans;
  ExecuteSql('update imoref set IRF_Surete="", IRF_SureteL="", IRF_SureteM=0, IRF_SureteD="' + DateToStr(Idate1900) + '"');
  UpdateNUM2('IMOASS', 'IAS');
  UpdateNUM2('IMOD205', 'ID5');
  UpdateNUM2('IMODOT', 'IDA');
  UpdateNUM2('IMOECA', 'ICT');
  UpdateNUM2('IMOFAC', 'IFC');
  UpdateNUM2('IMOFIS', 'IFS');
  UpdateNUM2('IMOLS0', 'ILL');
  UpdateNUM2('IMOMET', 'IMT');
  UpdateNUM2('IMOPER', 'IPR');
  UpdateNUM2('IMOREF', 'IRF');
  UpdateNUM2('IMORVL', 'IRV');
  UpdateNUM2('IMOSO1', 'IS1');
  UpdateNUM2('IMOSO2', 'IS2');
  UpdateNUM2('IMOSUB', 'ISB');
  UpdateNUM2('IMOVEN', 'IVC');
  UpdateNUM2('IMOVI1', 'IV1');
  UpdateNUM2('IMOVI2', 'IV2');
  CommitTrans;
  BeginTrans;
  MajTableDureeTaux;
  CommitTrans;
  *)

  //PAIE
  ExecuteSQL('UPDATE STAGE SET PST_COMPTERENDU="-",PST_ATTESTPRESENC="-",PST_QUESTAPPREC="-",PST_QUESTEVALUAT="-",PST_AUTREEVALUAT="",PST_SUPPCOURS="-",PST_VIDEOPROJ="-",PST_RETROPROJ="-",PST_JEUXROLE="-",PST_ETUDECAS="-",PST_AUTREMOYEN=""');
  //  ExecuteSQL('UPDATE SESSIONSTAGE SET PSS_FORMATION1="",PSS_FORMATION2="",PSS_FORMATION3="",PSS_FORMATION4="",PSS_FORMATION5="",PSS_FORMATION6="",PSS_FORMATION7="",PSS_FORMATION8="",PSS_NUMFACTURE="",PSS_SUPPCOURS="-",PSS_VIDEOPROJ="-",PSS_RETROPROJ="-",PSS_JEUXROLE="-",PSS_ETUDECAS="-",PSS_AUTREMOYEN="",PSS_ENVMAILSAL="-",PSS_ENVMAILRESP="-",PSS_ENVMAILAVER="-",PSS_EDITCONVOC="-",PSS_AVECCLIENT="-",PSS_DUREESTAGE=0,PSS_JOURSTAGE=0,PSS_HEUREDEBUT="'+UsTime(0)+'",PSS_HEUREFIN="'+UsTime(0)+'"');
  St := '';
  for i := 1 to 8 do St := 'PSS_FORMATION' + IntToStr(i) + '="",';
  St := St + 'PSS_NUMFACTURE="",PSS_SUPPCOURS="-",PSS_VIDEOPROJ="-",PSS_RETROPROJ="-",PSS_JEUXROLE="-",PSS_ETUDECAS="-",' +
    'PSS_AUTREMOYEN="",PSS_ENVMAILSAL="-",PSS_ENVMAILRESP="-",PSS_ENVMAILAVER="-",PSS_EDITCONVOC="-",PSS_AVECCLIENT="-",PSS_DUREESTAGE=0,PSS_JOURSTAGE=0,PSS_HEUREDEBUT="' +
      UsTime(0) + '",PSS_HEUREFIN="' + UsTime(0) + '"';
  ExecuteSQL('UPDATE SESSIONSTAGE SET ' + St);
  ExecuteSQL('UPDATE FORMATIONS SET PFO_FORMATION1="",PFO_FORMATION2="",PFO_FORMATION3="",PFO_FORMATION4="",PFO_FORMATION5="",PFO_FORMATION6="",PFO_FORMATION7="",PFO_FORMATION8="",PFO_DATECREATION="'
    + UsDateTime(IDate1900) + '",PFO_HEUREDEBUT="' + UsTime(0) + '",PFO_HEUREFIN="' + UsTime(0) + '"');
  ExecuteSQL('UPDATE INSCFORMATION SET PFI_FORMATION1="",PFI_FORMATION2="",PFI_FORMATION3="",PFI_FORMATION4="",PFI_FORMATION5="",PFI_FORMATION6="",PFI_FORMATION7="",PFI_FORMATION8=""');
  ExecuteSQL('UPDATE LIEUFORMATION SET PLF_PLAFPARIS="-"');
  ExecuteSQL('UPDATE FRAISSALPLAF SET PFP_ORDRE=1,PFP_PROVINCE="X",PFP_PARIS="-"');
  ExecuteSQL('UPDATE ELTNATIONAUX SET PEL_DECALMOIS="X" WHERE PEL_DECALMOIS="" OR PEL_DECALMOIS  IS NULL');
  ExecuteSQL('UPDATE ABSENCESALARIE SET PCN_ETATPOSTPAIE="VAL"');
  ExecuteSQL('UPDATE ETABCOMPL SET ETB_NBACQUISCP=""');
  ExecuteSQL('UPDATE ABSENCESALARIE SET PCN_VALIDSALARIE="SAL" WHERE PCN_VALIDSALARIE<>"RES" AND PCN_VALIDSALARIE<>"ADM"');

  MajJuridiqueTypePer;

  // JURI
  SQL := 'update JURIDIQUE set JUR_CAPEURO="' + SS + '" , JUR_LIQUIDDATE="' + SS + '", JUR_STATDATEENR="' + SS + '" ';
  SQL := SQL + ', JUR_DATEREPENG="' + SS + '", JUR_DATELIBCAP="' + SS + '", JUR_DATECPTES="' + SS + '"';
  SQL := SQL + ', JUR_COMPTESCONSO="-", JUR_CPTESDIV="-", JUR_CPTESBIL=0, JUR_CPTESCA=0, JUR_CPTESEFF=0';
  SQL := SQL + ',  JUR_PRTDATEAG="' + SS + '", JUR_PRTRECONST="' + SS + '", JUR_PRTASUIVRE="-", JUR_NRESTAT="' + SS + '"';
  SQL := SQL + ', JUR_NREDG="' + SS + '", JUR_AGPROPPEE="' + SS + '"';
  ExecuteSQL(SQL);
  ExecuteSQL('UPDATE JUFORMEJUR SET JFJ_PAYS="FRA"');

  // Fiscalité personnelle
  SQL := 'UPDATE F9065_MEMBRE SET B95_N01360=0,B95_N01361=0,B95_N01362=0,B95_N01363=0';
  SQL := SQL + ',B95_N01364=0,B95_N01365=0,B95_N01366=0,B95_C01368=0,B95_C01369=0';
  SQL := SQL + ',B95_C01370=0,B95_C01371=0,B95_N01372=0,B95_N01373=0,B95_N01374=0';
  SQL := SQL + ',B95_N01375=0,B95_N01376=0,B95_C01377=0,B95_C01378=0,B95_C01379=0,B95_C01380=0 ';
  ExecuteSQL(SQL);

  //Commun
  //AddNewNaturesGC;

end;

{$IFDEF CEGID}

procedure MAJVER595;
begin
  ExecuteSQL('UPDATE PARACTIONS SET RPA_MODIFRESP="-"');
  ExecuteSQL('UPDATE TIERS SET T_ENSEIGNE=""');
  if V_PGI.SAV then LogAGL('Fin Mise à jour à ' + DateTimeToStr(Now));
end;
{$ENDIF}

procedure MAJVER601;
var ss, Sf, Sql: string;
begin
  SS := UsDateTime(IDate1900);
  SF := UsDateTime(IDate2099);
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  // GI/GA
  //ExecuteSql('UPDATE PARPIECE SET GPP_VENTEACHAT="AUT" WHERE GPP_NATUREPIECEG="AFF"  OR GPP_NATUREPIECEG="PAF"');
  ExecuteSql('UPDATE RESSOURCE SET ARS_DEPOT="",ARS_SITE="",ARS_DATESORTIE="'
            + SF + '",ARS_MENSUALISE="X"');
  ExecuteSql('UPDATE AFTABLEAUBORD SET ATB_SOCIETEGROUPE="",ATB_REGROUPEFACT="",ATB_DATEDEBUT="'
            + SS + '",ATB_DATEFIN="' + SS + '",ATB_DATELIMITE="' + SS + '",ATB_DATEGARANTIE="' + SS
            + '",ATB_DATEDEBEXER="' + SS + '",ATB_DATEFINEXER="' + SS + '"');
  ExecuteSql('UPDATE AFPLANNING SET APL_CREATEUR="",APL_UTILISATEUR="",APL_ACTIVITEREPRIS="F",APL_DATECREATION="' + SS + '",APL_DATEMODIF="' + ss + '"');
  ExecuteSql('Update TACHE SET  ATA_METHODEDECAL="P",ATA_CREATEUR="",ATA_UTILISATEUR="",ATA_ACTIVITEREPRIS="F",ATA_DATECREATION="' + SS + '",ATA_DATEMODIF="' +
             ss + '"');
  ExecuteSql('update factaff set afa_numechebis=afa_numeche');
  // suppression de la tablette fantome fantome
  ExecuteSQL('DELETE FROM CHOIXCOD WHERE CC_TYPE="AFE"');


// on n'est pas dans la base commune, il faut remonter les info dans la DB0000

// if TableSurAutreBase('JOURFERIE') then by PCS and MD 09112004 DESHARE désactivé dans MajStruct
  if not IsMonoOuCommune then
  begin
    v_pgi.enableDEShare := True;
    Sql := 'INSERT INTO ##DP##.JOURFERIE (AJF_CODEFERIE,AJF_LIBELLE,AJF_JOURFIXE,'
      + 'AJF_JOUR,AJF_MOIS,AJF_ANNEE,AJF_PREDEFINI,AJF_NODOSSIER)'
      + 'SELECT * FROM ##DOSSIER##.JOURFERIE ';
    ExecuteSql(Sql);
    ExecuteSql('DELETE FROM ##DOSSIER##.JOURFERIE');
    v_pgi.enableDEShare := false;
  end;


  // Juri
  ExecuteSQL('update ANNULIEN set ANL_DIVBRUT=0, ANL_DIVNET=0, ANL_DIVAF=0, ANL_DIVANNEE=""');
  ExecuteSQL('update ANNULIEN set ANL_INFO="" where ANL_INFO is null');
  ExecuteSQL('update ANNULIEN set ANL_ACTNAT="" where ANL_ACTNAT is null');
  ExecuteSQL('update JURIDIQUE set JUR_IMPODIR="" where JUR_IMPODIR is null');
  ExecuteSQL('update JURIDIQUE set JUR_STATLIEUENR="" where JUR_STATLIEUENR is null');
  ExecuteSQL('update JURIDIQUE set JUR_STATBORDENR="" where JUR_STATBORDENR is null');
  //GRC
  ExecuteSQL('UPDATE SUSPECTS set RSU_PUBLIPOSTAGE="-"');
  ExecuteSQL ('UPDATE TIERS SET T_ENSEIGNE="" where T_ENSEIGNE is Null');
  ExecuteSQL('UPDATE PARACTIONS SET RPA_MODIFRESP="-" where RPA_MODIFRESP is null');
end;



procedure MAJVER603;
var ss, Sf, Sql: string;
begin
  SS := UsDateTime(IDate1900);
  SF := UsDateTime(IDate2099);
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  // Mode - Initialisation nouveau champ
  if Vers_PropTransfEnt < 103 then
  begin
    ExecuteSQL('update PROPTRANSFENT set GTE_LIMITEART="STD", GTE_ENCOURSCF="-", GTE_COMMANDEFOU="-"');
    ExecuteSQL('update PROPTRANSFLIG set GTL_GENECF="-"');
    ExecuteSQL('update PROPTRANSFLIG set GTL_GENERER="-" Where (GTL_GENERER IS NULL) or (GTL_GENERER="")');
    ExecuteSQL('UPDATE PROPTRANSFENT SET GTE_STKNET="-", GTE_STKTRANSIT="-", GTE_STKNETBTQ="-", GTE_STKTRANSITBTQ="-", GTE_LIMITEENREG="-"');
    ExecuteSQL('UPDATE PROPTRANSFLIG SET GTL_RESERVECLI=0, GTL_RESERVEFOU=0, GTL_PREPACLI=0');
    ExecuteSQL('DELETE FROM COMMUN WHERE CO_TYPE="GTM" and CO_CODE="ALF"');
    ExecuteSQL('DELETE FROM PROPMETETAB WHERE GTQ_CODEMETPAFF IN (SELECT GTM_CODEMETPAFF FROM PROPMETHODE WHERE GTM_TYPEMETPAFF="ALF")');
    ExecuteSQL('DELETE FROM PROPMETHODE WHERE GTM_TYPEMETPAFF="ALF"');
  end;

  if Vers_ExportAscii < 103 then ExecuteSQL('update EXPORTASCII_ENTETE set GAS_FORMATFIXE="X"');

  ExecuteSQL('update ARTICLECOMPL set GA2_COLLECTIONBAS=""');

  ExecuteSQL('UPDATE PARPIECE SET GPP_MAJINFOTIERS="X" WHERE GPP_NATUREPIECEG NOT IN ("EEX","SEX","TEM","TRE","TRV","FFO","FFA") AND GPP_MAJINFOTIERS IS NULL');
  ExecuteSQL('UPDATE PARPIECE SET GPP_MAJINFOTIERS="-" WHERE GPP_NATUREPIECEG IN ("EEX","SEX","TEM","TRE","TRV","FFO","FFA") AND GPP_MAJINFOTIERS IS NULL');
  if Vers_ParCaisse < 110 then ExecuteSQL('UPDATE PARCAISSE SET GPK_LIGNESUIVAUTO="X", GPK_NBJANNULTIC=0');
  if Vers_ParCaisse < 111 then ExecuteSQL('UPDATE PARCAISSE SET GPK_MODEDETAXE="000",GPK_CLIOBLIGDETAXE="-"');
  if Vers_ParCaisse < 112 then ExecuteSQL('UPDATE PARCAISSE SET GPK_TIROIRPORT="",GPK_TIROIRPARAM="",GPK_AFFPORT="",GPK_AFFPARAM=""');
  ExecuteSQL('UPDATE MODEPAIE SET MP_MONTANTMIN=0');
  if Vers_Piece < 158 then UpdateDecoupePiece('GP_DETAXE="-",GP_NUMDETAXE=""') ;
  ExecuteSQL('UPDATE REPRISEGPAO SET GRE_ENTETELIG="X" WHERE GRE_ENTETELIG<>"-"');
  if Vers_TarifMode < 110 then ExecuteSQL('UPDATE TARIFMODE SET GFM_NATURETYPE="VTE"');
	if Vers_Etabliss < 21 then
	begin
	    ExecuteSQL('UPDATE ETABLISS SET ET_TYPETARIFACH="" WHERE ET_TYPETARIFACH IS NULL');
	    ExecuteSQL('UPDATE ETABLISS SET ET_DEVISEACH="" WHERE ET_DEVISEACH IS NULL');
  end;
  if Vers_ParCaisse < 113 then ExecuteSQL('UPDATE PARCAISSE SET GPK_LIAISONREG="-"');
// fidélité
  if Vers_Etabliss < 22 then ExecuteSQL('UPDATE ETABLISS SET ET_PROGRAMME="" WHERE ET_PROGRAMME IS NULL');
  CreationEnrTYPEREMISE;

 // Resynchro GL_VIVANTE par rapport à GP_VIVANTE
 UpdateDecoupeLigne('GL_VIVANTE=(SELECT GP_VIVANTE FROM PIECE WHERE GP_NATUREPIECEG=GL_NATUREPIECEG AND GP_SOUCHE=GL_SOUCHE AND GP_NUMERO=GL_NUMERO AND GP_INDICEG=GL_INDICEG)',
                      ' AND GL_VIVANTE<>(SELECT GP_VIVANTE FROM PIECE WHERE GP_NATUREPIECEG=GL_NATUREPIECEG AND GP_SOUCHE=GL_SOUCHE AND GP_NUMERO=GL_NUMERO AND GP_INDICEG=GL_INDICEG)');

 //UpdateDecoupeLigne('GL_QTERESTE=GL_QTESTOCK', ' AND GL_VIVANTE="X" AND GL_TYPEDIM in ("GEN", "NOR", "DIM") ' );
 //UpdateDecoupeLigne('GL_QTERESTE=GL_QTESTOCK', ' AND GL_NATUREPIECEG="FFO" ') ;

 sql:= ' AND ((GL_VIVANTE="X") OR (GL_NATUREPIECEG IN (select gpp_naturepieceg from PARPIECE where gpp_actionfini in ("ENR","COM") and ((gpp_qteplus like "%PHY%") OR (GPP_QTEMOINS like "%PHY%"))))) AND GL_TYPEDIM in ("GEN", "NOR", "DIM") ';
 UpdateDecoupeLigne('GL_QTERESTE=GL_QTESTOCK', sql );

end;

procedure MAJVER604;
var ss, Sf: string;     
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SS := UsDateTime(IDate1900);
  SF := UsDateTime(IDate2099);
  //Système/Noyau
  ExecuteSQL('UPDATE UTILISAT SET US_EMAIL="", US_EMAILPASSWORD="", US_EMAILPOPSERVER="", US_EMAILSMTPSERVER=""') ;
  if IsMonoOuCommune then
  begin
    ExecuteSQL('UPDATE JUTYPEEVT SET JTE_DOMAINEACT="" WHERE JTE_DOMAINEACT IS NULL');
    // MDesgoutte 21/03/2006  ExecuteSQL('UPDATE JUEVENEMENT SET JEV_OUTLOOK="-", JEV_NODOSSIER="", JEV_DOMAINEACT=""') ;
    ExecuteSQL('UPDATE JUEVENEMENT SET JEV_DOMAINEACT=(SELECT JTE_DOMAINEACT FROM JUTYPEEVT where JTE_CODEEVT=JEV_CODEEVT) '
               +' WHERE exists (SELECT JTE_DOMAINEACT from JUTYPEEVT where JTE_CODEEVT=JEV_CODEEVT and JTE_DOMAINEACT="JUR")');
    ExecuteSQL('UPDATE JUEVENEMENT SET JEV_NODOSSIER=(SELECT DOS_NODOSSIER'
      + ' FROM  DOSSIER WHERE DOS_CODEPER=JEV_CODEPER)'
      + ' WHERE exists (SELECT DOS_NODOSSIER FROM DOSSIER WHERE DOS_CODEPER=JEV_CODEPER AND DOS_NODOSSIER IS NOT NULL)');
  end;

  //GPAO
  MajGPAOApres;

  {-- Mise à jour de PAPIECE --}
  ExecuteSQL('UPDATE PARPIECE SET GPP_TARIFMODULE="201" WHERE (GPP_VENTEACHAT="VEN")');
  ExecuteSQL('UPDATE PARPIECE SET GPP_TARIFMODULE="101" WHERE (GPP_VENTEACHAT="ACH") and (GPP_PILOTEORDRE="-")');
  ExecuteSQL('UPDATE PARPIECE SET GPP_TARIFMODULE="301" WHERE (GPP_VENTEACHAT="ACH") and (GPP_PILOTEORDRE="X")');

  ExecuteSQL('UPDATE PARPIECE SET GPP_INSERTLIG="X"');
  ExecuteSQL('UPDATE PARPIECE SET GPP_INSERTLIG="-" WHERE GPP_NATUREPIECEG="BSA"');
  ExecuteSQL('UPDATE PARPIECE SET GPP_PILOTEORDRE="-"');
  ExecuteSQL('UPDATE PARPIECE SET GPP_PILOTEORDRE="X" WHERE GPP_NATUREPIECEG IN ("BSA","CSA")');

  ExecuteSQL('UPDATE PARPIECE SET GPP_CFGART="-",GPP_CFGARTASSIST="",GPP_IMPAUTOBESCBN="-",GPP_IMPAUTOETATCBN="-",'
	  +' GPP_IMPBESOIN="-",GPP_PIECEEDI="-",GPP_PRIXNULOK="-",GPP_REFEXTCTRL="000",GPP_REFINTCTRL="000",GPP_REFINTEXT="INT"');

end;

Procedure MajVer605;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  InitTablePaie605;
End;


Procedure MajVer606;
var ss : string;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SS := UsDateTime(IDate1900);
  //CHR
  ExecuteSQL('UPDATE HRCALENDRIER SET HCA_HRTARIF="",HCA_COULEUR="",HCA_NIVEAUYIELD=0');
  ExecuteSQL('UPDATE HRDOSRES SET HDR_HRPERIODEDEBUT="",HDR_HRPERIODEFIN="",HDR_NATIONALITE=HDR_PAYS, HDR_PHONETIQUE=""');
  ExecuteSQL('UPDATE HRDOSSIER SET HDC_BOOLLIBRE1="-",HDC_BOOLLIBRE2="-",HDC_BOOLLIBRE3="-", HDC_CHARLIBRE1="", HDC_CHARLIBRE2="",HDC_CHARLIBRE3="",'+
             'HDC_DATELIBRE1="'+SS+'", HDC_DATELIBRE2="'+SS+'", HDC_DATELIBRE3="'+SS+'",'+
             'HDC_LIBREHDC1="", HDC_LIBREHDC2="", HDC_LIBREHDC3="",HDC_LIBREHDC4="", HDC_LIBREHDC5="",HDC_LIBREHDC6="",HDC_LIBREHDC7="",HDC_LIBREHDC8="",'+
             'HDC_LIBREHDC9="", HDC_LIBREHDCA="",HDC_VALLIBRE1=0,HDC_VALLIBRE2=0,HDC_VALLIBRE3=0,'+
             'HDC_NATIONALITE=HDC_PAYS, HDC_PHONETIQUE="", HDC_REGROUPELIGNE=""');
  ExecuteSQL('UPDATE HRPARAMPLANNING SET HPP_MODEPLANNING="", HPP_PLANNINGTYPEET="-", HPP_PLANNINGTYPETD="-",HPP_PLANNINGTYPEGR="-",HPP_PLANNINGTYPEPC="-"');
  ExecuteSQL('UPDATE HRPREFACT SET HPF_NUMLIGNOMEN=0, HPF_TYPEARTICLE=""');
  ExecuteSQL('UPDATE HRTYPDOS SET HTD_FAMRES=""');
  ExecuteSQL('UPDATE HRTYPRES SET HTR_HRETAT="", HTR_HRETATINTERNE=""');
  //Juri
  ExecuteSQL('delete from CHOIXCOD WHERE CC_TYPE="JFE" and CC_CODE="FAC"');
  //GC
  ExecuteSQL('UPDATE ARTICLE SET GA_REFCONSTRUC=""');  //142
  //GI/GA
  ExecuteSQL('UPDATE AFTABLEAUBORD SET ATB_ARTICLE=""');
  InsertChoixCode('ATU', '041', 'méclient', '', 'métiers');
  // DP
  ExecuteSQL('delete FROM CHOIXEXT where YX_TYPE="DCN"');
end;

Procedure MajVer607;
var ss,SQL, Creat, Usr : string;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SS := UsDateTime(IDate1900);
  //GRC
  { nouveaux champs ACTIONS }
  SQL:='UPDATE ACTIONS set RAC_CHAINAGE="---",RAC_NUMLIGNE=0,RAC_MAILENVOYE="-"';
  SQL:=SQL+',RAC_DERNIERMAIL="' + SS + '",RAC_MAILAUTO="-"';
  ExecuteSQL(SQL);
  { nouveaux champs ACTIONSCHAINEES }
  SQL:='UPDATE ACTIONSCHAINEES set RCH_CHAINAGE=""';
  ExecuteSQL(SQL);
  { nouveaux champs PARACTIONS }
  Creat := 'CEG'; Usr := 'CEG';
  SQL:='UPDATE PARACTIONS set RPA_CHAINAGE="---",RPA_ETATACTION="",RPA_HEUREACTION="' + SS + '"';
  SQL:=SQL+',RPA_INTERVENANT="",RPA_MAILANUAUTO="-",RPA_MAILETATANU="-",RPA_MAILETATPRE="-"';
  SQL:=SQL+',RPA_MAILETATREA="-",RPA_MAILETATVAL="-",RPA_MAILPREAUTO="-",RPA_MAILREAAUTO="-"';
  SQL:=SQL+',RPA_MAILRESPANU="",RPA_MAILRESPAUTO="-",RPA_MAILRESPPRE="",RPA_MAILRESPREA=""';
  SQL:=SQL+',RPA_MAILRESPVAL="",RPA_MAILVALAUTO="-",RPA_NBJOURS=0';
  SQL:=SQL+',RPA_NIVIMP="",RPA_NUMLIGNE=0,RPA_REPLICLIB="-",RPA_NONSUPP="-"';
  SQL:=SQL+',RPA_CREATEUR="'+Creat+'",RPA_UTILISATEUR="'+Usr+'"';
  SQL:=SQL+',RPA_DATECREATION="'+UsTime(NowH)+'",RPA_DATEMODIF="'+UsDateTime(Encodedate(1900,01,02))+'"';
  ExecuteSQL(SQL);
  { transformation d'un booléan en Combo avec adaptation des valeurs }
  ExecuteSQL('UPDATE PARACTIONS set RPA_MAILRESP="RSP" WHERE RPA_MAILRESP LIKE "X%"');
  ExecuteSQL('UPDATE PARACTIONS set RPA_MAILRESP="INT" WHERE RPA_MAILRESP LIKE "-%"');
  { correction anomalie : ce code manquait }
  InsertChoixCode ('RLZ','BLO','Bloc-notes ','','');
  { correction anomalie : ces code ne servent plus }
  ExecuteSQL('DELETE FROM CHOIXCOD WHERE CC_TYPE="RTA"');
  ExecuteSQL ('UPDATE CONTACT SET C_EMAILING="-"');
  ExecuteSQL ('UPDATE TIERS SET T_EMAILING="-"');
  ExecuteSQL ('UPDATE OPERATIONS SET ROP_FERME="-"');

  //Compta
  ExecuteSQL ('update eexbq set ee_avancement=""');
  //Mode
  if Vers_Affcdeentete < 102 then ExecuteSQL('UPDATE AFFCDEENTETE SET GEA_VERIFARTFERME="-", GEA_VERIFCLIFERME="-", GEA_ECLATEPARCDE="-", GEA_NATPIECEGRP="PRE"');
  if Vers_Affcdeligne < 102 then ExecuteSQL('UPDATE AFFCDELIGNE SET GEL_QTESTOCK=0') ;
end;


Procedure MajVer608;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSQL ('update EAITRANSENT set EAI_AUTODELAY=0 where EAI_AUTODELAY is null');
  end;  // not IsDossierPCL
        //GI/GA
  InsertChoixCode('ATU', '009', 'cette ressource', '', 'cet assistant');
  ExecuteSql ('Delete from choixcod  where cc_type="ATU" and cc_code="015"');
  If Vers_Activite < 126 then
    ExecuteSQL('UPDATE ACTIVITE SET ACT_NUMORDRE = 0, ACT_MONTANTTVA=0,ACT_QTETARIF = 0, ACT_MNTREMISE = 0, ACT_UNITETARIF = "", ACT_FORMULEVAR = ""');
  If Vers_Affaire < 137 then
    ExecuteSQL('UPDATE AFFAIRE SET AFF_FORCODE1 = "", AFF_FORCODE2 = ""');
  If Vers_ligne < 168 then
     UpdateDecoupeLigne( 'GL_FORMULEVAR = "", GL_GENERAUTO = "MAN", GL_FACTURABLE = "F", GL_LIGNELIEE = 0, GL_QUALIFQTETARIF = "",GL_FORCODE1 = "", GL_FORCODE2 = "", GL_REGULARISABLE = "-", GL_REGULARISE = "-"','');
  If Vers_article < 143 then
    ExecuteSQL('UPDATE ARTICLE SET GA_FORMULEVAR = ""');
  // fin GI/GA
  //GPAO
  UpdateDecoupeLigne('GL_BLOQUETARIF="-",'
    +'GL_DATEDEPARTUSINE=GL_DATELIVRAISON,GL_IDENTIFIANTWOL=0,'
    +'GL_REFEXTERNE="",'   // supprimé en 644 GL_REMISELIBRE1=0,GL_REMISELIBRE2=0,GL_REMISELIBRE3=0,'
    +'GL_TARIFSPECIAL="",GL_TRANSFERT="NON",GL_VALEURFIXEDEV=0', ' AND GL_TRANSFERT IS NULL');
  UpdateDecoupePiece('GP_IDENTIFIANTWOT=0,GP_NUMEROCONTACT=0,GP_REPRESENTANT2="",GP_REPRESENTANT3="",GP_TARIFSPECIAL=""');
  //Mode
  ExecuteSQL('UPDATE PARCAISSE SET GPK_LIAISONREG="-" WHERE GPK_LIAISONREG<>"X"  ');
  //Gescom
  ExecuteSQL('update PARPIECE set GPP_MONTANTMINI=0  ');
  ExecuteSQL('update TIERSPIECE set GTP_MONTANTMINI=0 ');
end;

Procedure MajVer609;
var ss,SQL : string;
    i : integer;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SS := UsDateTime(IDate1900);
  { nouveaux champs CONTACTS}
  SQL:='UPDATE CONTACT SET C_FERME="-",C_DATEFERMETURE="'+SS+'",C_SERVICECODE=""';
  SQL:=SQL+',C_LIBRECONTACT4="",C_LIBRECONTACT5="",C_LIBRECONTACT6=""';
  SQL:=SQL+',C_LIBRECONTACT7="",C_LIBRECONTACT8="",C_LIBRECONTACT9="",C_LIBRECONTACTA=""';
  ExecuteSQL(SQL);
  for i := 4 to 9 do InsertChoixCode('ZLI', 'BT' + intTostr(i), '.-Table libre ' + intTostr(i), '', '');
  InsertChoixCode('ZLI', 'BTA', '.-Table libre 10', '', '');

  // Mode
  ExecuteSQL('UPDATE ARTICLE SET GA_TYPEARTFINAN="ECA" WHERE GA_TYPEARTICLE="FI" AND GA_CODEARTICLE IN (SELECT GPK_CTRLECART FROM PARCAISSE)');

  //CHR
  ExecuteSQL('UPDATE HRPARAMPLANNING SET HPP_AXECONTINGENT1="", HPP_AXECONTINGENT2="", HPP_COULVEILJF=""');
  ExecuteSQL('UPDATE HRDOSSIER SET HDC_INFOSBANQUE="",HDC_REFEXTERNE=""');
  // Paie
  ExecuteSQL('UPDATE DUCSENTETE SET PDU_CENTREPAYEUR=""');
  ExecuteSQL('UPDATE ORGANISMEPAIE SET POG_CPINFO="",POG_CENTREPAYEUR=""');
  ExecuteSQL('UPDATE HISTOBULLETIN SET PHB_ORIGINELIGNE="CPA" WHERE (PHB_RUBRIQUE LIKE "3210.%" OR PHB_RUBRIQUE LIKE "4300.%")');
  ExecuteSQL('update HISTOSALARIE SET PHS_ETABLISSEMENT = (SELECT PPU_ETABLISSEMENT FROM PAIEENCOURS WHERE PPU_DATEFIN=HISTOSALARIE.PHS_DATEAPPLIC AND HISTOSALARIE.PHS_DATEEVENEMENT=PPU_DATEDEBUT AND HISTOSALARIE.PHS_SALARIE=PPU_SALARIE)');
  ExecuteSQL('update HISTOSALARIE SET PHS_BETABLISSEMENT = "-"');

end;

Procedure MajVer610;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  // Paie
  ExecuteSQL('UPDATE ORGANISMEPAIE SET POG_LIBEDITBULL=""');
  ExecuteSQL('UPDATE ECRITODPAIE SET PEC_ALEAT=0');
  //mode
  if Vers_PropTransfEnt < 105 then
  begin
    ExecuteSQL('update PROPTRANSFENT set GTE_LIMITEART="STD", GTE_ENCOURSCF="-", GTE_COMMANDEFOU="-"');
    ExecuteSQL('update PROPTRANSFLIG set GTL_GENECF="-"');
    ExecuteSQL('update PROPTRANSFLIG set GTL_GENERER="-" Where (GTL_GENERER IS NULL) or (GTL_GENERER="")');
    ExecuteSQL('UPDATE PROPTRANSFENT SET GTE_STKNET="-", GTE_STKTRANSIT="-", GTE_STKNETBTQ="-", GTE_STKTRANSITBTQ="-", GTE_LIMITEENREG="-"');
    ExecuteSQL('UPDATE PROPTRANSFLIG SET GTL_RESERVECLI=0, GTL_RESERVEFOU=0, GTL_PREPACLI=0');
    ExecuteSQL('DELETE FROM COMMUN WHERE CO_TYPE="GTM" and CO_CODE="ALF"');
    ExecuteSQL('DELETE FROM PROPMETETAB WHERE GTQ_CODEMETPAFF IN (SELECT GTM_CODEMETPAFF FROM PROPMETHODE WHERE GTM_TYPEMETPAFF="ALF")');
    ExecuteSQL('DELETE FROM PROPMETHODE WHERE GTM_TYPEMETPAFF="ALF"');
  end;

End;

Procedure MajVer612;
var ss : string;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  MajGPAOApres612;
// GI-GA
SS := UsDateTime(IDate1900);
If Vers_Activite < 127 then
  ExecuteSQL('UPDATE ACTIVITE SET ACT_NUMPIECEORIG = ""');

If Vers_AfRevision < 4 then
begin
  ExecuteSQL('UPDATE AFREVISION SET AFR_ValInit1=0,AFR_ValInit2=0,AFR_ValInit3=0,'
    +'AFR_ValInit4=0,AFR_ValInit5=0,AFR_ValInit6=0,AFR_ValInit7=0,AFR_ValInit8=0,'
    +' AFR_ValInit9=0,AFR_ValInit10=0,AFR_DATEINDICE1="'+SS
    +'",AFR_DATEINDICE2="'+SS+'",AFR_DATEINDICE3="'+SS
    +'",AFR_DATEINDICE4="'+SS+'",AFR_DATEINDICE5="'+SS
    +'",AFR_DATEINDICE6="'+SS+'",AFR_DATEINDICE7="'+SS
    +'",AFR_DATEINDICE8="'+SS+'",AFR_DATEINDICE9="'+SS
    +'",AFR_DATEINDICE10="'+SS+'"');
  ExecuteSQL('UPDATE AFORMULEVAR SET AVF_QTEARRONDI="X"');
  ExecuteSQL('UPDATE AFORMULEVARQTE SET AVV_QTECALCUL=0');
end;

end;

Procedure MajVer613;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //Compta
  ExecuteSQL('UPDATE ECRITURE SET E_TRESOSYNCHRO="CRE"');
  //DP
  ExecuteSQL('update YMESSAGES set YMS_USERMAIL="" where YMS_MAILID=0');
  ExecuteSQL('update YMESSAGES set YMS_USERMAIL=YMS_USERDEST where YMS_MAILID>0');
  //Paie
  ExecuteSQL ('DELETE FROM DECOMBOS WHERE DO_COMBO in  ("PGRETENUESAL","PGREPORTFORM","PGREFUSFORM")') ;
  ExecuteSQL ('DELETE FROM COMMUN WHERE CO_TYPE="PAF" AND CO_CODE="ILL"');
  ExecuteSQL ('UPDATE SALARIES SET PSA_CODEPCS82=PSA_CODEEMPLOI');
End;

Procedure MajVer614;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
//GPAO
ExecuteSQL('UPDATE DISPO SET GQ_PRODUCTION=0,GQ_CONSOMMATION=0 WHERE GQ_PRODUCTION IS NULL');
ExecuteSQL('UPDATE WORDREBES SET WOB_QAFFSAIS=0,WOB_QAFFSTOC=0 WHERE WOB_QAFFSAIS IS NULL');
//-- Mise à jour des champs pour les arrondis de prix --//
// ExecuteSQL('UPDATE DEVISE SET D_ARRONDIPRIXACHAT=0.01,D_ARRONDIPRIXVENTE=0.01 WHERE D_DECIMALE<>0');
//DP
ExecuteSQL('UPDATE JUEVENEMENT SET JEV_RESSOURCE="", JEV_AFFAIRE="", JEV_EXTERNE="-", JEV_ABSENCE="-"');
ExecuteSQL('UPDATE JUTYPEEVT SET JTE_CODEARTICLE="", JTE_EXTERNE="-", JTE_ABSENCE="-"');
If Not ExisteSQL('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="JFE" AND CC_CODE="ACT"') then
  ExecuteSQL('INSERT INTO CHOIXCOD(CC_TYPE, CC_CODE, CC_LIBELLE, CC_ABREGE, CC_LIBRE) '
    + 'VALUES ("JFE", "ACT", "Activité de l''agenda", "Agenda", "95")' );
If Not ExisteSQL('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="JFE" AND CC_CODE="TRA"') then
  ExecuteSQL('INSERT INTO CHOIXCOD(CC_TYPE, CC_CODE, CC_LIBELLE, CC_ABREGE, CC_LIBRE) '
    + 'VALUES ("JFE", "TRA", "Travaux", "Travaux", "81")' );
// PAIE
ExecuteSQL('UPDATE CONTRATTRAVAIL SET PCI_MONTANTCT=0');
ExecuteSQL('UPDATE CONTRATTRAVAIL SET PCI_ETABLISSEMENT=(SELECT PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE=CONTRATTRAVAIL.PCI_SALARIE)');
ExecuteSQL('UPDATE ABSENCESALARIE SET PCN_NBREMOIS=0 WHERE PCN_NBREMOIS IS NULL');
ExecuteSQL('UPDATE DEPORTSAL SET PSE_STATUTCOTIS="", PSE_METIERBTP=""');

end;

Procedure MajVer615;
var SS : string;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
SS := UsDateTime(IDate1900); // nb creer SS

  if not IsDossierPCL then
  begin
    ExecuteSQL('UPDATE IMOREF SET IRF_DATEMODIF="'+USTime(NowH)+'" ') ;
    ExecuteSQL('UPDATE IMOLS1 SET ILS_DATEMODIF="'+USTime(NowH)+'" ') ;
  end;

//GPAO

//-- MAJ TABLE DEPOTS --//
ExecuteSQL('UPDATE DEPOTS SET GDE_DEPOTTRANSIT="",GDE_DEPOTLITIGE="", GDE_STOCKBLOQUE="-"  WHERE GDE_DEPOTTRANSIT IS NULL');
//-- MAJ TABLE PROPTRANSFENT --//
ExecuteSQL('UPDATE PROPTRANSFENT SET  GTE_PIECEPRECEDENT="", GTE_LIBRE1="", GTE_DATEEDITION=" ' + SS +' " WHERE GTE_PIECEPRECEDENT IS NULL');
//-- MAJ TABLE PROPTRANSFLIG --//
ExecuteSQL('UPDATE PROPTRANSFLIG SET  GTL_EMPLACEMENT="", GTL_PROPINITIAL=0, GTL_QTESAISIE=0, GTL_QTETRANS=0 WHERE GTL_EMPLACEMENT IS NULL');
//-- MAJ TABLE WCBNEVOLUTION --//
ExecuteSQL('UPDATE WCBNEVOLUTION SET  WEV_LIBRE1=""  WHERE WEV_LIBRE1 IS NULL');
//-- Table PROFILART --//
ExecuteSQL('UPDATE PROFILART SET GPF_TYPEPROFILART="SIM" WHERE GPF_TYPEPROFILART IS NULL OR GPF_TYPEPROFILART=""');
//-- PARPIECE sous traitance --//
ExecuteSQL('UPDATE PARPIECE SET GPP_INSERTLIG="-" WHERE GPP_NATUREPIECEG="BSA"');

//DP
ExecuteSQL('update YREPERTPERSO set YRP_CIVILITE=""');
ExecuteSQL('update UTILISAT set US_EMAILLOGIN=""');
ExecuteSQL('update YDOCUMENTS set YDO_NATDOC=""');
ExecuteSQL('update DOSSIER set DOS_CABINET="-" where DOS_NODOSSIER<>"000000"');
ExecuteSQL('update DOSSIER set DOS_CABINET="X" where DOS_NODOSSIER="000000"');
//GI-GA
If Vers_AfRevision < 5 then
begin
    ExecuteSQL('UPDATE AFREVISION SET AFR_PUBCODE1="",AFR_PUBCODE2="",AFR_PUBCODE3="",'
     +'AFR_PUBCODE4="",AFR_PUBCODE5="",AFR_PUBCODE6="",AFR_PUBCODE7="",AFR_PUBCODE8="",'
     +'AFR_PUBCODE9="",AFR_PUBCODE10="" ,AFR_INDDATEVAL1="'+SS+'" ,AFR_INDDATEVAL2="'+SS
     +'" ,AFR_INDDATEVAL3="'+SS+'" ,AFR_INDDATEVAL4="'+SS+'" ,AFR_INDDATEVAL5="'+SS
     +'" ,AFR_INDDATEVAL6="'+SS+'" ,AFR_INDDATEVAL7="'+SS+'" ,AFR_INDDATEVAL8="'+SS
     +'" ,AFR_INDDATEVAL9="'+SS+'" ,AFR_INDDATEVAL10="'+SS+'"');
end;
ExecuteSql ('update ressource  set ARS_TAUXUNIT2=0,ARS_TAUXUNIT3=0,ARS_TAUXUNIT4=0');
ExecuteSql ('update activite set ACT_PAIENUMFIC=0');
//GRC
ExecuteSQL('UPDATE ACTIONS SET RAC_DUREEACTION=0');
RTMajDureeAction;
ExecuteSQL('UPDATE SUSPECTS set RSU_FORMEJURIDIQUE="",RSU_REGION="",RSU_EMAILING="-",RSU_CONTACTPUBLI="-",RSU_CONTACTEMLG="-"');
// GC
ExecuteSql ('update MODEPAIE set MP_JALCAISSE=""');

ExecuteSql ('update juevenement set JEV_GENEREGI="-"');
ExecuteSql ('update jutypeevt  set JTE_AFFAIRE=""');

End;

Procedure MajVer616;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
ExecuteSql ('update cpprorata set pa_journal=""');
Init_WInitChampLig;   // erreur à reprendre
End;

Procedure MajVer617;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
ExecuteSql ('update NOMENENT  set GNE_IDENTIFIANTWNT=0');
ExecuteSql ('update NOMENLIG  set GNL_IDENTIFIANTWNT=0');
GPAOMajApres_617;
End;

Procedure MajVer618;
var SS : string;
Begin
  SS := UsDateTime(IDate1900); // nb creer SS
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  ExecuteSql ('update LIGNECOMPL  set GLC_IDENTIFIANTWNT=0');
  //GIGA
  ExecuteSQL('UPDATE PARAMOBLIG SET  GOB_CODE="B01" WHERE GOB_CODE ="A04"');
  ExecuteSQL('UPDATE PARAMOBLIG SET  GOB_CODE="B02" WHERE GOB_CODE ="A05"');
  ExecuteSQL('UPDATE PARAMOBLIG SET  GOB_CODE="B03" WHERE GOB_CODE ="A06"');
  ExecuteSQL('UPDATE PARAMOBLIG SET  GOB_CODE="B04" WHERE GOB_CODE ="A07"');
  ExecuteSQL('UPDATE PARAMOBLIG SET  GOB_CODE="B05" WHERE GOB_CODE ="A03"');
  //GP
  ExecuteSQL('UPDATE WFORMULECHAMP'
          +' SET WDX_CHAMPCALC="GLC_QACCSAIS", WDX_UNITECHAMPCALC="GLC_UNITEACC"'
          +' WHERE WDX_CONTEXTE="GL" AND WDX_QUALIFIANT="QAC"');
  ExecuteSQL('DELETE FROM WFORMULECHAMPDEF'
          +' WHERE (WDY_CONTEXTE="GL" AND WDY_NOMCHAMP="GL_QACCSAIS")'
          +' OR (WDY_CONTEXTE="GL" AND WDY_NOMCHAMP="GL_UNITEACC")');
  if Not ExisteSQL('SELECT WDY_NOMCHAMP FROM WFORMULECHAMPDEF WHERE WDY_CONTEXTE="GL" AND WDY_NOMCHAMP="GLC_QACCSAIS"') then
    ExecuteSQL('INSERT INTO WFORMULECHAMPDEF'
          +' VALUES ("GL", "GLC_QACCSAIS", "QTE", "' + USDATETIME(iDate1900) + '", "' + USDATETIME(iDate1900) + '", "", "")');

  if Not ExisteSQL('SELECT WDY_NOMCHAMP FROM WFORMULECHAMPDEF WHERE WDY_CONTEXTE="GL" AND WDY_NOMCHAMP="GLC_UNITEACC"') then
    ExecuteSQL('INSERT INTO WFORMULECHAMPDEF'
          +' VALUES ("GL", "GLC_UNITEACC", "UNI", "' + USDATETIME(iDate1900) + '", "' + USDATETIME(iDate1900) + '", "", "")');
  //Mode
  ExecuteSQL('UPDATE PARCAISSE SET GPK_LIGNESUIVAUTO="X", GPK_NBJANNULTIC=0, GPK_MODEDETAXE="000", GPK_CLIOBLIGDETAXE="-", GPK_TIROIRPORT="", GPK_TIROIRPARAM="", GPK_AFFPORT="", GPK_AFFPARAM="" WHERE GPK_NBJANNULTIC IS NULL');
  ExecuteSQL('UPDATE PARPIECE SET GPP_FILTRECOMM="001" WHERE GPP_NATUREPIECEG IN ("FFO","FFA") AND GPP_FILTRECOMM=""');
  //Compta
  ExecuteSQL('UPDATE SOUCHE SET SH_DATEDEBUT="' + USDateTime(iDate1900) + '" WHERE SH_DATEDEBUT="' + USDateTime(iDate1900-2) + '"'); // FQ 12445


End;

Procedure MajVer619;
var SS : string;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SS := UsDateTime(IDate1900); // nb creer SS
// ga suite chgmt ordre tri tablette

ExecuteSql ('UPDATE TACHE SET ATA_MOISJOURLIB = "J1" WHERE ATA_MOISJOURLIB = "LUN"');
ExecuteSql ('UPDATE TACHE SET ATA_MOISJOURLIB= "J2" WHERE ATA_MOISJOURLIB = "MAR"');
ExecuteSql ('UPDATE TACHE SET ATA_MOISJOURLIB= "J3" WHERE ATA_MOISJOURLIB = "MER"');
ExecuteSql ('UPDATE TACHE SET ATA_MOISJOURLIB= "J4" WHERE ATA_MOISJOURLIB = "JEU"');
ExecuteSql ('UPDATE TACHE SET ATA_MOISJOURLIB= "J5" WHERE ATA_MOISJOURLIB = "VEN"');
ExecuteSql ('UPDATE TACHE SET ATA_MOISJOURLIB= "J6" WHERE ATA_MOISJOURLIB = "SAM"');
ExecuteSql ('UPDATE TACHE SET ATA_MOISJOURLIB= "J7" WHERE ATA_MOISJOURLIB = "DIM"');
 
(* mcd 660 ExecuteSql ('UPDATE TACHERESSOURCE SET ATR_MOISJOURLIB = "J1" WHERE ATR_MOISJOURLIB = "LUN"');
ExecuteSql ('UPDATE TACHERESSOURCE SET ATR_MOISJOURLIB = "J2" WHERE ATR_MOISJOURLIB = "MAR"');
ExecuteSql ('UPDATE TACHERESSOURCE SET ATR_MOISJOURLIB = "J3" WHERE ATR_MOISJOURLIB = "MER"');
ExecuteSql ('UPDATE TACHERESSOURCE SET ATR_MOISJOURLIB = "J4" WHERE ATR_MOISJOURLIB = "JEU"');
ExecuteSql ('UPDATE TACHERESSOURCE SET ATR_MOISJOURLIB = "J5" WHERE ATR_MOISJOURLIB = "VEN"');
ExecuteSql ('UPDATE TACHERESSOURCE SET ATR_MOISJOURLIB = "J6" WHERE ATR_MOISJOURLIB = "SAM"');
ExecuteSql ('UPDATE TACHERESSOURCE SET ATR_MOISJOURLIB = "J7" WHERE ATR_MOISJOURLIB = "DIM"');*)

ExecuteSql ('UPDATE AFMODELETACHE SET AFM_MOISJOURLIB = "J1" WHERE AFM_MOISJOURLIB = "LUN"');
ExecuteSql ('UPDATE AFMODELETACHE SET AFM_MOISJOURLIB = "J2" WHERE AFM_MOISJOURLIB = "MAR"');
ExecuteSql ('UPDATE AFMODELETACHE SET AFM_MOISJOURLIB = "J3" WHERE AFM_MOISJOURLIB = "MER"');
ExecuteSql ('UPDATE AFMODELETACHE SET AFM_MOISJOURLIB = "J4" WHERE AFM_MOISJOURLIB = "JEU"');
ExecuteSql ('UPDATE AFMODELETACHE SET AFM_MOISJOURLIB = "J5" WHERE AFM_MOISJOURLIB = "VEN"');
ExecuteSql ('UPDATE AFMODELETACHE SET AFM_MOISJOURLIB = "J6" WHERE AFM_MOISJOURLIB = "SAM"');
ExecuteSql ('UPDATE AFMODELETACHE SET AFM_MOISJOURLIB = "J7" WHERE AFM_MOISJOURLIB = "DIM"');

ExecuteSQL('UPDATE TIERS  SET T_ETATRISQUE="" WHERE T_ETATRISQUE is NULL');
// fin GA

ExecuteSQL('UPDATE PARPIECE SET GPP_SENSPIECE="ENT" WHERE (GPP_NATUREPIECEG="BFA")');
End;

Procedure MajVer620;
var SS : string;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SS := UsDateTime(IDate1900); // nb creer SS
  //mode
  ExecuteSQL('DELETE FROM USERPREFDIM WHERE GUD_CHAMP LIKE "%CON"');
  //GC
  //ExecuteSQL('UPDATE PARPIECE SET GPP_ACTIONFINI="ENR" WHERE GPP_NATUREPIECEG="TEM"');
  //GP
  //-- Table CONTACT --//
  ExecuteSQL('UPDATE CONTACT SET C_TIERS = "" WHERE (C_TIERS is NULL)');
  ExecuteSQL('UPDATE CONTACT SET C_NATUREAUXI=(SELECT T_NATUREAUXI FROM TIERS WHERE T_AUXILIAIRE=CONTACT.C_AUXILIAIRE) WHERE C_TYPECONTACT="T"');
  ExecuteSQL('UPDATE CONTACT SET C_TIERS=(SELECT T_TIERS FROM TIERS WHERE (T_NATUREAUXI=CONTACT.C_NATUREAUXI) and (T_AUXILIAIRE=CONTACT.C_AUXILIAIRE)) WHERE C_TYPECONTACT="T"');

  //-- Table PARPIECE --//
  ExecuteSQL('UPDATE PARPIECE SET GPP_MAJINFOTIERS="X" WHERE GPP_NATUREPIECEG NOT IN ("EEX","SEX","TEM","TRE","TRV","FFO","FFA")');
  ExecuteSQL('UPDATE PARPIECE SET GPP_MAJINFOTIERS="-" WHERE GPP_NATUREPIECEG IN ("EEX","SEX","TEM","TRE","TRV","FFO","FFA")');
  // COMPTA
  // Nature BGE, Etats BAL et BTL
  SupprimeEtat('E' ,'BGE','BAL');
  SupprimeEtat('E' ,'BGE','BTL');
  // Nature BAU, Etats BAX et BAG et BAV
  SupprimeEtat('E' ,'BAU','BAX');
  SupprimeEtat('E' ,'BAU','BAG');
  SupprimeEtat('E' ,'BAU','BAV');
  // Nature GLA, Etats GLB et ECH et GAG
  SupprimeEtat('E' ,'GLA','GLB');
  SupprimeEtat('E' ,'GLA','ECH');
  SupprimeEtat('E' ,'GLA','GAG');
  // Nature GLG, Etat GLT
  SupprimeEtat('E' ,'GLG','GLT');
  // Nature JDI, Etats JDP et JDD et JDE et JDL
  SupprimeEtat('E' ,'JDI','JDP');
  SupprimeEtat('E' ,'JDI','JDD');
  SupprimeEtat('E' ,'JDI','JDE');
  SupprimeEtat('E' ,'JDI','JDL');

  // Corrections après difusion PCL
  // Pour correction bug sur fiche fournisseurs et fiche tiers
  ExecuteSQL('UPDATE TIERS SET T_CONFIDENTIEL="1" WHERE T_CONFIDENTIEL="X"');
  ExecuteSQL('UPDATE TIERS SET T_CONFIDENTIEL="0" WHERE T_CONFIDENTIEL="-"');
  // ok si socref > 620 CorrigeRTTiers;
  // ExecuteSQL ('update parpiece set gpp_actionfini="COM" where gpp_naturepieceg="TEM"');

  // Nature JCA Journal de caisse
  SupprimeEtat ('E' ,'JCA','COL');
  SupprimeEtat ('E' ,'JCA','FAM');
  SupprimeEtat ('E' ,'JCA','REG');
  SupprimeEtat ('E' ,'JCA','VEN');

  AglNettoieListes('RTMULOUVFERMTIERS;RTMULQUALITE;RTMULQUALITECONTA;RTMULSUPPRTIERS;'
       +'RTMULTIERS;RTMULTIERSACTIONS;RTMULTIERSCL;RTMULTIERSGROUPE;RTMULTIERSLOT;'
       +'RTMULTIERSMAILING;RTMULTIERSPRESCRI;RTMULTIERSSIREN;RTTIERSMAILINGFIC', '',nil);

  // correction du nouveau module Artciles pour S3 apporté avec une mauvaise confidentialité dans la SOCREF     
  ExecuteSQL('UPDATE MENU SET MN_ACCESGRP="'+StringOfChar('0',100)+'" WHERE (MN_1=266) OR (MN_1=0 AND MN_2=266)');

End;

Procedure MajVer625;
var SS : string;
    i:integer;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SS := UsDateTime(IDate1900);

  //GRC
  { ajout 2 champs obligatoires dans la liste }
  AglNettoieListes('RTMULACTIONS', 'RAC_TYPEACTION;RAC_INTERVENANT',nil);
  { bug champs libres suspects : on remplace les codes TX0..2 par TL0..2 }
  if ExisteSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="RSZ" AND CC_CODE LIKE "TX%"') then
  begin
    //for i := 0 to 2 Do
       //executeSql('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE ,CC_LIBELLE,CC_ABREGE,CC_LIBRE) SELECT "RSZ","TL'+IntToStr(i)+'",CC_LIBELLE,CC_ABREGE,CC_LIBRE FROM CHOIXCOD WHERE CC_TYPE="RSZ" and cc_code="TX'+IntToStr(i)+'"');
    for i := 0 to 2 do InsertChoixCode('RSZ', 'TL' + intTostr(i), 'Texte libre ' + intTostr(i + 1), '', '');
    ExecuteSql ('DELETE FROM CHOIXCOD WHERE CC_TYPE="RSZ" AND CC_CODE LIKE "TX%"');
  end;

  if not IsDossierPCL then
  begin
    //GPAO
    ExecuteSQL('UPDATE WPDRTET  SET WPE_AVECQPCB="-"');
    ExecuteSQL('UPDATE WPDRTYPE SET WRT_AVECQPCB="-"');
    //GC
    ExecuteSQL('UPDATE PORT SET GPO_DATESUP="' + USDATETIME(iDate2099) + '" WHERE GPO_DATESUP is null');

    //-- AJOUT DE DONNEES DANS LA TABLE EDIFAMILLEEMG --//
    if not ExisteSQL('SELECT EFM_CODEFAMILLE FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="ARC"') then
      ExecuteSQL('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE, EFM_NATURESPIECE, EFM_SERIALISE) '
                +'VALUES("ARC", "Accusé de réception de commande EDI_ARC", "ACV,CC,CCE,CCR", "-")')
    else
      ExecuteSQL('UPDATE EDIFAMILLEEMG SET EFM_NATURESPIECE="ACV,CC,CCE,CCR", EFM_SERIALISE="-"'
                +'WHERE EFM_CODEFAMILLE="ARC"');

    if not ExisteSQL('SELECT EFM_CODEFAMILLE FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="BL"') then
      ExecuteSQL('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE, EFM_NATURESPIECE, EFM_SERIALISE) '
                +'VALUES("BL", "Bon de livraison standard EDI_BL", "ALF,ALV,APV,BLC,LCE,LCR,LFR,PRE", "-")')
    else
      ExecuteSQL('UPDATE EDIFAMILLEEMG SET EFM_NATURESPIECE="ALF,ALV,APV,BLC,LCE,LCR,LFR,PRE", EFM_SERIALISE="-"'
                +'WHERE EFM_CODEFAMILLE="BL"');

    if not ExisteSQL('SELECT EFM_CODEFAMILLE FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="CDE"') then
      ExecuteSQL('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE, EFM_NATURESPIECE, EFM_SERIALISE) '
                +'VALUES("CDE", "Commande standard EDI_CDE", "ACV,CC,CCE,CCR", "-")')
    else
      ExecuteSQL('UPDATE EDIFAMILLEEMG SET EFM_NATURESPIECE="ACV,CC,CCE,CCR", EFM_SERIALISE="-"'
                +'WHERE EFM_CODEFAMILLE="CDE"');

    if not ExisteSQL('SELECT EFM_CODEFAMILLE FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="FAC"') then
      ExecuteSQL('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE, EFM_NATURESPIECE, EFM_SERIALISE) '
                +'VALUES("FAC", "Facture EDI_FAC", "ABT,AF,AFP,AFS,APR,AVC,AVS,FAC,FBT", "-")')
    else
      ExecuteSQL('UPDATE EDIFAMILLEEMG SET EFM_NATURESPIECE="ABT,AF,AFP,AFS,APR,AVC,AVS,FAC,FBT", EFM_SERIALISE="-"'
                +'WHERE EFM_CODEFAMILLE="FAC"');

    //GIGA
    If Vers_Afrevision < 6 then
     begin
     ExecuteSql ('UPDATE AFREVISION SET AFR_COEFRACCORD1=0,AFR_COEFRACCORD2=0,'
        +'AFR_COEFRACCORD3=0,AFR_COEFRACCORD4=0,AFR_COEFRACCORD5=0,AFR_COEFRACCORD6=0,'
        +'AFR_COEFRACCORD7=0,AFR_COEFRACCORD8=0,AFR_COEFRACCORD9=0,AFR_COEFRACCORD10=0');
     end;
  end;   //not IsDossierPCL

InsertChoixCode('ATU', '094', 'scotraire', '', 'temporaire');

End;

Procedure MajVer626;
var SS : string;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SS := UsDateTime(IDate1900);
  //-- Tables MODEPAIE, PIEDECHE, ACOMPTES -- n° carte --//
  ExecuteSQL('UPDATE MODEPAIE SET MP_AVECINFOCOMPL="-", MP_AVECNUMAUTOR="-", MP_COPIECBDANSCTRL="-", MP_AFFICHNUMCBUS="-"');

  if not IsDossierPCL then
  begin
    UpdateDecoupePiedEche('GPE_CBINTERNET="", GPE_CBLIBELLE="", GPE_DATEEXPIRE="", GPE_TYPECARTE="", GPE_CBNUMCTRL="", GPE_CBNUMAUTOR="", GPE_NUMCHEQUE=""');
    UpdateDecoupeAcomptes('GAC_CBNUMCTRL="", GAC_CBNUMAUTOR=""');
    //-- Tables PARCAISSE --//
    ExecuteSQL('UPDATE PARCAISSE SET GPK_CLISAISIENOM="-"');   //114
    //BTP
      //-- Table AFFAIRE --//
    ExecuteSQL('UPDATE AFFAIRE SET AFF_PREPARE = "-" WHERE AFF_ETATAFFAIRE IN ("ENC","ACP","ACC","REF")');
    ExecuteSQL('UPDATE AFFAIRE SET AFF_PREPARE = "X" WHERE AFF_ETATAFFAIRE IN ("CLO","TER")');
    ExecuteSQL('UPDATE LIGNE SET GL_QTERESTE = GL_QTEFACT WHERE GL_NATUREPIECEG IN ("DBT","FBT","DAC","ABT","ETU")');
    ExecuteSQL('UPDATE BTPARDOC SET BPD_DESCREMPLACE = "-"');
    //GPAO
    ExecuteSQL('UPDATE PARPIECE SET GPP_LIBELLE= "Réception SST d''achat" WHERE GPP_NATUREPIECEG="BSA"');
  End; //not IsDossierPCL
End;

Procedure MajVer627;
var SS : string;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SS := UsDateTime(IDate1900);

  //GIGA
  AglNettoieListes('AFVALINDICE', 'AFV_DEFINITIF;AFV_INDCODE;AFV_PUBCODE;AFV_INDDATEVAL',nil);
  AglNettoieListes('AFMULRECHAFFAIRE', 'AFF_STATUTAFFAIRE',nil);

  if not IsDossierPCL then
  begin
    //MODE
    ExecuteSQL('UPDATE PARCAISSE SET GPK_TPEREPERTOIRE="",GPK_CLAVIERPISTE="-"'); //115
   end; //not IsDossierPCL
End;

Procedure MajVer628;
var SS : string;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SS := UsDateTime(IDate1900);
  if not IsDossierPCL then
  begin
    //GC
    ExecuteSQL('UPDATE CONDITIONNEMENT SET GCO_DEFAUT="-",GCO_TYPESFLUX=""');
  end; // not IsDossierPCL
End;


Procedure MajVer630;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSQL('update imoref set irf_statut="" where irf_statut is null');
    ExecuteSQL('update imovi1 set iV1_statut="" where iV1_statut is null');
    ExecuteSQL('update imovi2 set iV2_statut="" where iV2_statut is null');
  end;  // not IsDossierPCL
end;

Procedure MajVer631;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
//GIGA
//AglNettoieListes('AFMULACTIVITECON1', 'ACT_NUMLIGNEUNIQUE;ACT_TYPEACTIVITE',nil);
  if not IsDossierPCL then
  begin
       //GPI
       ExecuteSQL('UPDATE ARTICLECOMPL SET GA2_ESTUL="-" WHERE GA2_ESTUL IS NULL');
  end; //PasDossierPGI
End;

Procedure MajVer633;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  ExecuteSQL('UPDATE YFILES SET YFI_FILECOMPRESSED = "X", YFI_FILESTORAGE = -1, YFI_FILEDATECREATE = YFI_DATECREATION, YFI_FILEDATEMODIFY = YFI_DATECREATION WHERE YFI_FILESTORAGE IS NULL');
  if not IsDossierPCL then
  begin
    //---------- ARTICLE ------------------
    ExecuteSQL('update ARTICLE set ga_lasstraitance="MAN" where ga_lasstraitance is null');
    //---------- PROFILART ------------------
    ExecuteSQL('update PROFILART set gpf_lasstraitance="MAN" where gpf_lasstraitance is null');
    //---------- WNOMELIG ------------------
    ExecuteSQL('update WNOMELIG set wnl_lasstraitance="MAN" where wnl_lasstraitance is null');
    //---------- WORDREBES ------------------
    ExecuteSQL('update WORDREBES set wob_lasstraitance="MAN" where wob_lasstraitance is null');
    //---------- WNOMEDEC ------------------
    ExecuteSQL('update WNOMEDEC set wnd_lasstraitance="MAN" where wnd_lasstraitance is null');
    //--------- WINITCHAMPLIG -----------------------
    if not ExisteSQL('SELECT 1 FROM  WINITCHAMPLIG WHERE WIL_NOMTABLE="ARTICLE" and WIL_NOMCHAMP="GA_LASSTRAITANCE" and WIL_INITCTXTYPE="DEF" and  WIL_IDENTIFIANT=75 and WIL_NUMORDRE=12') then
      ExecuteSQL('insert into WINITCHAMPLIG (WIL_NOMTABLE,WIL_NOMCHAMP,WIL_INITCTXTYPE,WIL_INITCTXCODE'
               +',WIL_IDENTIFIANT,WIL_NUMORDRE,WIL_PRIORITE, WIL_C1OPERATEUR, WIL_C1VALUE'
               +',WIL_C2OPERATEUR, WIL_C2BRANCHEMENT,WIL_C2VALUE,WIL_C3OPERATEUR'
               +',WIL_C3BRANCHEMENT,WIL_C3VALUE,WIL_INITVALEUR'
               +',WIL_DEFINITVALEUR,WIL_OBLIGATOIRE,WIL_CUSTOMIZABLE,WIL_FAMCHAMP,WIL_BYPARAMSOC,WIL_NUMVERSIONTET) '
               +' values ("ARTICLE","GA_LASSTRAITANCE","DEF","DEFAUT",75,12,"","","","","","","","","","MAN","MAN","X","X","DIV","-",1)');
    //---------- WCHAMP ------------------
    ExecuteSQL('delete FROM WCHAMP where WCA_CONTEXTEPROFIL="PRO" and WCA_NOMTABLE="ARTICLE" and WCA_NOMCHAMP="GA_DELAIPROD"');
    if not ExisteSQL('SELECT 1 FROM   WCHAMP WHERE WCA_CONTEXTEPROFIL="PRO" and WCA_NOMTABLE="ARTICLE" and WCA_NOMCHAMP="GA_LASSTRAITANCE"') then
      ExecuteSQL('insert into WCHAMP (WCA_CONTEXTEPROFIL,WCA_NOMTABLE,WCA_NOMCHAMP) values ("PRO","ARTICLE","GA_LASSTRAITANCE")');
     //GC
    ExecuteSQL('UPDATE CONDITIONNEMENT SET GCO_SAISICOND="-",GCO_TYPESFLUX="VEN;",GCO_DECOND="-" WHERE GCO_TYPESFLUX is null');
    // DBR pour le champ TypesFlux, il n'y avait que des ventes avant, donc normal de forcer a vente et pas vide.
  end; // not IsDossierPCL

  //TRESO
  MajFluxTreso;
End;
Procedure MajVer634;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GPAO
    //--------- WCBNTYPEMOUV -----------------------
    if not ExisteSQL('select 1 from wcbntypemouv where wtm_typemouvement = "PRP"') then
    begin
      ExecuteSQL('insert into WCBNTYPEMOUV (WTM_TYPEMOUVEMENT,WTM_LIBELLE) values ("PRP","Proposition transfert ID- prévision")');
    end;
    ExecuteSQL('update WCBNTYPEMOUV set WTM_LIBELLE="Proposition transfert ID- ferme" where WTM_TYPEMOUVEMENT="PRT"');
    ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA="CC;CCR;" WHERE SOC_NOM="SO_WORDREVTE"');
    ExecuteSQL('update parpiece set gpp_typearticle = "CNS;" || gpp_typearticle where gpp_typearticle not like "%CNS%"');
  end;  // not IsDossierPCL
End;

Procedure MajVer635;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  InsertChoixCode('ATU', '007', 'Chiffre des missions', '', 'Chiffre d''affaires');
  //GIGA
  // ajout 15092004 MCD
  AglNettoieListes('AFMULACTIVITECON1', 'ACT_NUMLIGNEUNIQUE;ACT_TYPEACTIVITE;ACT_TIERS;ACT_DATEACTIVITE;ACT_RESSOURCE;ACT_AFFAIRE;ACT_FOLIO;ACT_ACTIVITEREPRIS',nil);

  if not IsDossierPCL then
  begin
     // MCD 651 champs renommés if Vers_Piece < 162 then  UpdateDecoupePiece('GP_DATEDEBFAC =GP_DATEPIECE,GP_DATEFINFAC =GP_DATEPIECE');
     // MCD pour 651 if Vers_afcumul < 110 then    ExecuteSQL('UPDATE AFCUMUL SET ACU_TRAITECPTA = "X"');
     if Vers_activite < 130 then    ExecuteSQL('UPDATE ACTIVITE SET ACT_MONTANTTTC = 0');
     ExecuteSQL('UPDATE EACTIVITE SET EAC_MONTANTTTC = 0,EAC_MONTANTTVA=0');
     if  Vers_ArticleCompl  < 105 then    ExecuteSQL('UPDATE ARTICLECOMPL SET GA2_REMBOURSABLE = "X",GA2_RECUPERABLE="X",GA2_POURCRECUP=100');
  end;  // not IsDossierPCL
End;

Procedure MajVer636;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //GC
  ExecuteSQL('update CPPROFILUSERC set CPU_PCPACHAT = "-" where CPU_PCPACHAT is NULL ');
  ExecuteSQL('update CPPROFILUSERC set CPU_PCPVENTE = "-" where CPU_PCPVENTE is NULL');
  ExecuteSQL('update CPPROFILUSERC set CPU_FORCEDEPOT = "-" where CPU_FORCEDEPOT is NULL');
  ExecuteSQL('update CPPROFILUSERC set CPU_DEPOT = "" where CPU_DEPOT is NULL');
  ExecuteSQL('Update CPPROFILUSERC set CPU_GRPUTI = "GRP" where CPU_USER = "..."');
  ExecuteSQL('Update CPPROFILUSERC set CPU_GRPUTI = "UTI" where CPU_USER <> "..."');
  ExecuteSQL('Update CPPROFILUSERC set CPU_GRPUTI = "" where CPU_GRPUTI is NULL');
  //GPAO
  ExecuteSQL('Update TIERS set T_DELAIMOYEN = 0');

End;

Procedure MajVer637;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  { ajout 2 champs obligatoires dans la liste }
  AglNettoieListes('RTMULACTIONSTIERS', 'RAC_TYPEACTION;RAC_INTERVENANT',nil);

  if not IsDossierPCL then
  begin
    //GRC
    RT_InsertLibelleInfoComplContact;
    //GIGA
     if Vers_activite < 131 then ExecuteSQL('UPDATE ACTIVITE SET ACT_TRAITECPTA = 0');
  end;  // not IsDossierPCL
End;

Procedure MajVer639;
var i : integer;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  if not IsDossierPCL then
  begin
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT=""');
    for i := 26 to 34 do InsertChoixCode('RLZ', 'CL'+chr(i+55), 'Table libre ' + intTostr(i), '', '');

    //GPAO
    ExecuteSQL('UPDATE WNATURETRAVAIL SET WNA_WITHVALWGT1 = WNA_WITHVAL,WNA_WITHMODWGT1 = WNA_WITHMOD,WNA_WITHPERWGT1 = WNA_WITHPER,'
               +'WNA_WITHVALWNT2 = WNA_WITHVAL,WNA_WITHMODWNT2 = WNA_WITHMOD,WNA_WITHPERWNT2 = WNA_WITHPER,WNA_WITHVALWGT2 = WNA_WITHVAL,WNA_WITHMODWGT2 = WNA_WITHMOD,'
               +'WNA_WITHPERWGT2 = WNA_WITHPER,WNA_WITHVALWNT3 = WNA_WITHVAL,WNA_WITHMODWNT3 = WNA_WITHMOD,WNA_WITHPERWNT3 = WNA_WITHPER,WNA_WITHVALWGT3 = WNA_WITHVAL,'
               +'WNA_WITHMODWGT3 = WNA_WITHMOD,WNA_WITHPERWGT3 = WNA_WITHPER,wNA_MVTPRODUITFINI="", WNA_MVTCOMPOSANT="", WNA_MVTCOPRODUIT=""');
    ExecuteSQL('UPDATE WRECORDCMS SET WRD_RESSOURCE="" WHERE WRD_RESSOURCE IS NULL');
    ExecuteSQL('UPDATE DISPO SET GQ_AFFECTE=0, GQ_BLOCAGE=0 ,GQ_EMPLACEMENTACH="",GQ_EMPLACEMENTCON="",  GQ_EMPLACEMENTPRO=""'
               +', GQ_EMPLACEMENTVTE="",GQ_GEREEMPLACE="-",GQ_STOCKALERTE=0, GQ_STOCKRECOMPL=0, GQ_TRANSFERTRECU=0');
  //GRC
    ExecuteSQL('update prospects set rpr_rprlibtable26="",rpr_rprlibtable27="",rpr_rprlibtable28=""'
              +',rpr_rprlibtable29="",rpr_rprlibtable30="",rpr_rprlibtable31="",rpr_rprlibtable32=""'
              +',rpr_rprlibtable33="",rpr_rprlibtable34=""');
  end;  // not IsDossierPCL

  //DP
  ExecuteSQL('update DOSSIER set DOS_NETEXPERT="-", DOS_NECPSEQ=0');      // supprimé en 651
  If Not ExisteSQL('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="JFE" AND CC_CODE="NET"') then
     ExecuteSQL('INSERT INTO CHOIXCOD(CC_TYPE, CC_CODE, CC_LIBELLE, CC_ABREGE, CC_LIBRE) '
     + 'VALUES ("JFE", "NET", "Echanges NetExpert", "NetExpert", "63")' );
End;

Procedure MajVer640;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //DP
  ExecuteSQL('update jubibrubrique set jbr_histo = "-" ' );
  if not IsDossierPCL then
  begin
    // GPAO
    // ExecuteSql('update wnaturetravail set WNA_MVTPRODUITFINI="APR", WNA_MVTCOMPOSANT="RPR" where WNA_NATURETRAVAIL="FAB"');
    ExecuteSql('UPDATE WNATURETRAVAIL SET WNA_MVTPRODUITFINI="APR",WNA_MVTCOMPOSANT="RPR",WNA_MVTCOPRODUIT="ACP" WHERE WNA_NATURETRAVAIL="FAB"');
    ExecuteSql('update article set ga_gereparc = "-" where ga_gereparc is null') ;
    ExecuteSql('Update WINITCHAMPLIG set WIL_INITVALEUR = "UNI", WIL_DEFINITVALEUR = "UNI" where (WIL_INITVALEUR = "U") OR (WIL_DEFINITVALEUR = "U")');
    //--------- WCBNTYPEMOUV -----------------------
    if not ExisteSQL('select 1 from wcbntypemouv where wtm_typemouvement = "PRP"') then
    begin
      ExecuteSQL('insert into WCBNTYPEMOUV (WTM_TYPEMOUVEMENT,WTM_LIBELLE) values ("PRP","Proposition transfert ID- prévision")');
    end;
    ExecuteSQL('update WCBNTYPEMOUV set WTM_LIBELLE="Proposition transfert ID- ferme" where WTM_TYPEMOUVEMENT="PRT"');
    ExecuteSQL('UPDATE WNATURETRAVAIL SET WNA_TYPENATURE="PRD"');
    ExecuteSQL('UPDATE WORDRETET '
             + 'SET WOT_NOM="", WOT_RESSOURCE="", WOT_DATEINTER="'+UsDateTime(IDate1900)+'", WOT_REFEXTERNE="", WOT_DATEREFEXTERNE="' + USDATETIME(StrToDate('01/01/1900')) + '"'
             + 'WHERE WOT_NOM is null AND WOT_RESSOURCE is null AND WOT_DATEINTER is null AND WOT_REFEXTERNE is null AND WOT_DATEREFEXTERNE is null');
    ExecuteSQL('UPDATE WORDRELIG '
             + 'SET WOL_SERIEINTERNE="",  WOL_LIBELLE="" ' { WOL_ORDREINTER=0, suopprimé en 642 ?}
             + 'WHERE WOL_SERIEINTERNE is null AND WOL_LIBELLE is null');
  end;  // not IsDossierPCL

//
End;

Procedure MajVer641;
var SS : string;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SS := UsDateTime(IDate1900);
  //DP
  ExecuteSQL('update DOSSIER set DOS_PASSWORD="",DOS_NECPDATEARRET="'+SS+'", DOS_NERECNBFIC=0, DOS_NERECDATE="'+SS+'"');
  ExecuteSQL('UPDATE YDATATYPELINKS SET YDL_LCOMMUNE="-"');
  ExecuteSQL('UPDATE ETABLISS SET ET_NODOSSIER=""');
End;

Procedure MajVer642;
var SS : string;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SS := UsDateTime(IDate1900);
  if not IsDossierPCL then
  begin
    //GPAO
    UpdateDecoupePiece('GP_POIDSPESEE = 0');
    { MàJ table COLISAGE }
    ExecuteSQL('UPDATE COLISAGE '
              + 'SET GCS_PREPARATEUR="", GCS_EMBALLEUR="", '
              + 'GCS_CHARLIBRE1="", GCS_CHARLIBRE2="", '
              + 'GCS_VALLIBRE1=0, GCS_VALLIBRE2=0, '
              + 'GCS_BOOLLIBRE1="-", GCS_BOOLLIBRE2="-", '
              + 'GCS_DATELIBRE1="' + SS + '", GCS_DATELIBRE2="' + SS + '", '
              + 'GCS_DATEEMBALLAGE="' + SS + '", '
              + 'GCS_MODELE="", '
              + 'GCS_TYPECODEBARRE=(SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_TYPESSCC") '
              + 'WHERE GCS_PREPARATEUR is Null');
    { MàJ table ARTICLECOMPL }
    ExecuteSQL('UPDATE ARTICLECOMPL SET GA2_MODELE="" WHERE GA2_MODELE is Null');
    executesql('update stkmouvement set gsm_affaire = "" where gsm_affaire is null');
    executesql('update wordrelig set wol_rang=0, wol_Affaire="" where wol_rang is null and wol_Affaire is null');
    executesql('update dispo set GQ_UNITEMIN = "QTE", GQ_UNITEMAX = "QTE", GQ_UNITEALERTE = "QTE", GQ_UNITERECOMPL = "QTE" where GQ_UNITEMIN is null');
    ExecuteSQL('UPDATE QPHASE SET QPH_TRAITEMENT = "" WHERE QPH_TRAITEMENT IS NULL');
    ExecuteSQL('UPDATE QDETCIRC SET QDE_TRAITEMENT = "" WHERE QDE_TRAITEMENT IS NULL');
    ExecuteSQL('UPDATE WORDREPHASE SET WOP_TRAITEMENT = "" WHERE WOP_TRAITEMENT IS NULL');
    //ExecuteSQL('update WORDRETET set wot_affaire="",wot_affaire1="",wot_affaire2="",wot_affaire3="",wot_avenant=""  where wot_affaire is null');
    //Executesql('update wordrelig set wol_rang=0, wol_Affaire="",wol_Affaire1="",wol_Affaire2=""'
    //            +',wol_Affaire3="",wol_avenant="" '
    //            +' where wol_rang is null and wol_Affaire is null and wol_Affaire1 is null and wol_Affaire2 is null and wol_Affaire3 is null and wol_avenant is null');
    //ExecuteSQL('update WCBNEVOLUTION set wev_datedebut="' + SS  + '", wev_affaire="", wev_affaire1="", wev_affaire2="", wev_affaire3="", wev_avenant="" where wev_affaire1 is null');
    //Executesql('update stkmouvement set gsm_affaire = "",gsm_affaire1 = "",gsm_affaire2 = "",gsm_affaire3 = "",gsm_avenant = "" where gsm_affaire is null');
    //ExecuteSQL('UPDATE WORDRELIG SET WOL_TYPEDEFAIL="",WOL_NATUREDEFAIL="",WOL_GRAVITEDEFAIL="" WHERE WOL_TYPEDEFAIL IS NULL');
    // modif suite à suppression de champs en 644
    Executesql('UPDATE STKMOUVEMENT set gsm_affaire = "" where gsm_affaire is null');
    ExecuteSQL('UPDATE WCBNEVOLUTION set wev_datedebut="' + UsDateTime(iDate1900)  + '", wev_affaire="" where wev_affaire is null');
    ExecuteSQL('UPDATE WORDRETET set wot_affaire="" where wot_affaire is null');
    ExecuteSQL('UPDATE WORDRELIG SET WOL_TYPEDEFAIL="",WOL_NATUREDEFAIL="",WOL_GRAVITEDEFAIL="",wol_rang=0, wol_Affaire="" WHERE WOL_TYPEDEFAIL IS NULL');

    ExecuteSQL('UPDATE ARTICLE SET GA_GEREPARC="-" WHERE GA_GEREPARC IS NULL');


    W_SAV_COMPETENCE;
    W_SAV_COMPL;
  end;  // not IsDossierPCL
End;


Procedure MajVer644;
var SS : string;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SS := UsDateTime(IDate1900);
  if not IsDossierPCL then
  begin

    //GPAO
    Executesql('UPDATE LIGNECOMPL SET GLC_NATURETRAVAIL = "", GLC_LIGNEORDRE=0, GLC_OPECIRC="", GLC_TRAITEMENT="" WHERE GLC_NATURETRAVAIL IS NULL');
    ExecuteSQL('update STKMOUVEMENT set gsm_codelisteori="" where gsm_codelisteori is null');
    ExecuteSQL('update TRANSINVLIG set GIN_STATUTFLUX="",GIN_LOTINTERNE="",GIN_SERIEINTERNE="",GIN_TIERSPROP="" '
                       +',GIN_INDICEARTICLE="",GIN_MARQUE="", GIN_CHOIXQUALITE="",GIN_STATUTDISPO="",GIN_REFAFFECTATION="" '
                       +',GIN_LOT=(select GA_LOT from ARTICLE where GIN_ARTICLE=GA_ARTICLE)'
                       +',GIN_NUMEROSERIE=(select GA_NUMEROSERIE from ARTICLE where GIN_ARTICLE=GA_ARTICLE) '
                       +' where GIN_REFAFFECTATION is null');
     //GIGA
     GIGAApres644;
  end;  // not IsDossierPCL

  //PAIE
  ExecuteSQL('UPDATE ETABCOMPL SET ETB_EDITBULCP="NN1" WHERE ETB_CONGESPAYES="X"');
  ExecuteSQL('UPDATE ETABCOMPL SET ETB_NBRECPSUPP = 0, ETB_REGIMEALSACE ="-"');
  ExecuteSQL('UPDATE PAIEENCOURS SET PPU_CPACQUISMOD="-"');
  ExecuteSQL('UPDATE SALARIES SET PSA_TYPEDITBULCP="ETB"');
  ExecuteSQL('UPDATE SALARIES SET PSA_CONGESPAYES = ( SELECT ETB_CONGESPAYES FROM ETABCOMPL WHERE  PSA_ETABLISSEMENT = ETB_ETABLISSEMENT)');
  ExecuteSQL('update salaries set psa_cpacquissupp="PER" where psa_nbrecpsupp <> 0');
  ExecuteSQL('update salaries set psa_cpacquissupp="ETB" where psa_nbrecpsupp = 0');
  ExecuteSQL('UPDATE ABSENCESALARIE SET PCN_TYPEIMPUTE=""');
  ExecuteSQL('UPDATE ABSENCESALARIE SET PCN_TYPEIMPUTE="ACQ" WHERE PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="AJU"');
  ExecuteSQL('UPDATE ELTNATIONAUX SET PEL_REGIMEALSACE ="-" WHERE PEL_PREDEFINI <> "CEG"'); // modif P.Dumet 9/8/2004
  // P Dumet 31082004 ExecuteSQL('update variablepaie set PVA_MTARRONDI="7"');
  ExecuteSQL('update variablepaie set PVA_MTARRONDI="7" WHERE PVA_PREDEFINI <> "CEG"');
  ExecuteSQL('UPDATE ATTESTATIONS SET PAS_DATEPREAVISD2="'+SS+'",PAS_DATEPREAVISF2="'+SS+'"'
             +',PAS_MOTIFPREAVIS2="-",PAS_DATEPREAVISD3="'+SS+'",PAS_DATEPREAVISF3="'+SS+'"'
             +',PAS_MOTIFPREAVIS3="-",PAS_MOTIFPREAVIS1="-"');
  // ExecuteSQL('UPDATE DADSLEXIQUE SET PDL_EXERCICEDEB="", PDL_EXERCICEFIN=""');
  // remplacée à partir de 657 par les deux ligne suivantes (Vincent GALLIOT 31/08/2004)
  ExecuteSQL('UPDATE DADSLEXIQUE SET PDL_EXERCICEDEB="" WHERE PDL_EXERCICEDEB IS NULL ');
  ExecuteSQL('UPDATE DADSLEXIQUE SET PDL_EXERCICEFIN="" WHERE PDL_EXERCICEFIN IS NULL ');

  ExecuteSQL('UPDATE ORGANISMEPAIE SET POG_PREVOYANCE="-"');
  ExecuteSQL('UPDATE ETABCOMPL SET ETB_LIBELLE = (SELECT ET_LIBELLE FROM ETABLISS WHERE ETB_ETABLISSEMENT=ET_ETABLISSEMENT)');
  ExecuteSQL('UPDATE ETABCOMPl SET ETB_MEDTRAV=-1 WHERE ETB_MEDTRAV=0 AND NOT EXISTS (select * from annuaire where ANN_TYPEPER="MED" AND ANN_CODEPER=0)');
  // ExecuteSQL('UPDATE JEUECRPAIE SET PJP_TYPEPROVCP=""');
  ExecuteSQL('UPDATE JEUECRPAIE SET PJP_TYPEPROVCP="COD" WHERE PJP_PREDEFINI <> "CEG"'); //P.Dumet 08092004
  // ajout 1509204
  AglNettoieListes('PGMULMVTCP', 'PCN_TYPEMVT;PCN_ORDRE;PCN_ETABLISSEMENT;PCN_SALARIE;PCN_TYPECONGE;PCN_SENSABS;PCN_DATEVALIDITE;PCN_BASE;PCN_NBREMOIS;PCN_JOURS;PCN_APAYES',nil);
  AglNettoieListes('PGENVOIDUCS', 'PES_FICHIEREMIS',nil);

//PAIE
  ExecuteSQL('UPDATE MOTIFABSENCE SET PMA_OKSAISIESAL = "-", PMA_TYPEABS="" , PMA_GESTIONIJSS="-"');
  ExecuteSQL('UPDATE MOTIFABSENCE SET PMA_TYPEABS="RTT" WHERE PMA_TYPERTT="X"');
  ExecuteSQL('UPDATE ABSENCESALARIE SET PCN_NBJCARENCE=0, PCN_NBJCALEND=0, PCN_NBJIJSS=0, PCN_IJSSSOLDEE="-", PCN_MVTORIGINE=""');
  ExecuteSQL('UPDATE ATTESTATIONS SET PAS_TYPEABS=""');
  ExecuteSQL('UPDATE ETABCOMPL SET ETB_SUBROGATION="-"');
  ExecuteSQL('UPDATE SALARIES SET PSA_CATEGMAINTIEN=""');

  if not IsDossierPCL then
  begin
    //GPAO (TARIF)
    // init d'un champ ajouté en 699
    ExecuteSQL('UPDATE YTARIFSPARAMETRES  SET YFO_CODEPORT="" WHERE YFO_CODEPORT IS NULL ');
    // supprimé en 702 GPMajParamTarif;
    GPMajTypeTarif;
    // GPMajSTKNature; mis en fin de MajAvant systématique
    ExecuteSQL('UPDATE TARIF SET GF_TRANSFERE="-", GF_ATRANSFERER="", GF_TRFMESSAGE="" WHERE GF_TRANSFERE IS NULL');
    UpdateDecoupeLigne('GL_REMISELIBRE=0, GL_REMLIBCASCADE=""');
  end;  // not IsDossierPCL

  InsertChoixCode ('YTU','P','à la pièce','','');
  InsertChoixCode ('YTU','L','à la ligne','','');
  InsertChoixCode ('YTN','L','Des linéaires','Linéaires','');
  InsertChoixCode ('YTN','P','Des poids','Poids','');
  InsertChoixCode ('YTN','Q','Des quantités','Quantités','');
  InsertChoixCode ('YTN','S','Des surfaces','Surfaces','');
  InsertChoixCode ('YTN','V','Des volumes','Volumes','');
  InsertChoixCode ('YTR','P','à la pièce','','');
  InsertChoixCode ('YTR','L','à la ligne','','');

End;


Procedure MajVer645;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
  //GPAO
    ExecuteSQL('UPDATE ARTICLE SET GA_FAMILLEVALO="" WHERE GA_FAMILLEVALO IS NULL') ;
  end;  // not IsDossierPCL
  //PAIE
  ExecuteSQL('UPDATE CHOIXCOD SET CC_ABREGE=CC_CODE WHERE CC_TYPE="PMS"');
  ExecuteSQL('UPDATE EXERFORMATION SET PFE_TAUXBUDGET=0');
  ExecuteSQL('UPDATE DEPORTSAL SET PSE_MSATECHNIQUE="X"');
  ExecuteSQL('UPDATE DEPORTSAL SET PSE_MSATECHNIQUE="X" WHERE PSE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES LEFT JOIN TAUXAT ON PAT_ORDREAT=PSA_ORDREAT AND PAT_ETABLISSEMENT=PSA_ETABLISSEMENT WHERE PAT_CODEBUREAU<>"BUR")');
  ExecuteSQL('UPDATE DEPORTSAL SET PSE_MSATECHNIQUE="-" WHERE PSE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES LEFT JOIN TAUXAT ON PAT_ORDREAT=PSA_ORDREAT AND PAT_ETABLISSEMENT=PSA_ETABLISSEMENT WHERE PAT_CODEBUREAU="BUR")');
  ExecuteSQL('UPDATE ABSENCESALARIE SET PCN_GESTIONIJSS="-" WHERE PCN_TYPECONGE NOT IN (SELECT PMA_MOTIFABSENCE FROM MOTIFABSENCE WHERE PMA_GESTIONIJSS="X")');
  ExecuteSQL('UPDATE ABSENCESALARIE SET PCN_GESTIONIJSS="X" WHERE PCN_TYPECONGE IN (SELECT PMA_MOTIFABSENCE FROM MOTIFABSENCE WHERE PMA_GESTIONIJSS="X")');
  ExecuteSQL('UPDATE MOTIFABSENCE SET PMA_OKSAISIESAL="-"');
  // ajout 15092004
  AglNettoieListes('PGMULCONGES', 'PSA_ETABLISSEMENT;PSA_SALARIE;PSA_DATEENTREE;PSA_DATESORTIE;PSA_CONFIDENTIEL;PSA_CONGESPAYES',nil);
  AglNettoieListes('PGMULMVTCPPRISGRP', 'PCN_ETABLISSEMENT;PCN_SALARIE;PCN_TYPECONGE;PCN_DATEVALIDITE;PCN_ORDRE;PCN_TYPEMVT',nil);

End;

Procedure MajVer646;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    executesql('update stkmouvement set gsm_qprepa = 0 where gsm_qprepa is null');
  end;  // not IsDossierPCL
  //PAIE
  ExecuteSQL('UPDATE CURSUSSTAGE SET PCC_MILLESIME="0000"');
  ExecuteSQL('UPDATE CURSUS SET PCU_INCLUSCAT="X"');
  ExecuteSQL('UPDATE COTISATION  SET PCT_ACTIVITE="" WHERE PCT_PREDEFINI<>"CEG"');
  ExecuteSQL('UPDATE REMUNERATION  SET PRM_ACTIVITE="" WHERE PRM_PREDEFINI<>"CEG"');
  //DP
  ExecuteSQL('UPDATE YGEDDICO SET YGD_CODEGED = SUBSTRING(YGD_CODEGED, 1, 1)||"0"||SUBSTRING(YGD_CODEGED, 2, 1) WHERE (LEN(YGD_CODEGED) = 2)');
  ExecuteSQL('UPDATE YDATATYPETREES SET YDT_SCODE = SUBSTRING(YDT_SCODE , 1, 1)||"0"||SUBSTRING(YDT_SCODE , 2, 1) WHERE (LEN(YDT_SCODE ) = 2) AND (YDT_CODEHDTLINK = "YYGEDNIV1GEDNIV2")');
  ExecuteSQL('UPDATE DPDOCUMENT SET DPD_CODEGED = SUBSTRING(DPD_CODEGED, 1, 1)||"0"||SUBSTRING(DPD_CODEGED, 2, 1) WHERE (LEN(DPD_CODEGED) = 2)');
End;

Procedure MajVer647;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    // IMMO
    MajImo646;
    //GIGA
    ExecuteSql ('Update AFFAIRE set AFF_PERSPECTIVE=0');
  end;  // not IsDossierPCL
  InsertChoixCode('ATU', '905', 'marchandises', '', 'fournitures');
  // Compta
  ExecuteSql ('update journal set j_accelerateur="-"');
  // Paye
  ExecuteSQL('UPDATE DADS2SALARIES SET PD2_NTIC="", PD2_EPARGNERETRAI=0');
  ExecuteSQL('UPDATE DADS2HONORAIRES SET PDH_NTIC=""');
  ExecuteSQL('UPDATE DEPORTSAL SET PSE_BTPPOSITION="", PSE_BTPECHELON="", PSE_BTPCATEGORIE="", PSE_BTPAFFILIRCIP=""');
  ExecuteSQL('UPDATE DADS2HONORAIRES SET PDH_NTIC=""');
  ExecuteSQL('UPDATE REMUNERATION SET PRM_SAISIEARRETAC="-" WHERE PRM_PREDEFINI<>"CEG"');
End;

Procedure MajVer648;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  ExecuteSQL('UPDATE DOSSIER SET DOS_PWDGLOBAL="-", DOS_APPLISPROTEC="", DOS_NETEXPERTGED="-"');
  if not IsDossierPCL then
  begin
    //GRC
    ExecuteSQL('UPDATE ACTIONS SET RAC_AFFAIRE="",RAC_AFFAIRE0="",RAC_AFFAIRE1="",RAC_AFFAIRE2="",RAC_AFFAIRE3="",RAC_AVENANT=""');
    ExecuteSQL('UPDATE ACTIONSCHAINEES SET RCH_AFFAIRE="",RCH_AFFAIRE0="",RCH_AFFAIRE1="",RCH_AFFAIRE2="",RCH_AFFAIRE3="",RCH_AVENANT=""');
    ExecuteSQL('UPDATE ACTIONSGENERIQUES SET RAG_AFFAIRE="",RAG_AFFAIRE0="",RAG_AFFAIRE1="",RAG_AFFAIRE2="",RAG_AFFAIRE3="",RAG_AVENANT=""');
    //GPAO
    ExecuteSQL('UPDATE LIGNECOMPL SET GLC_PHASE = "" WHERE GLC_PHASE IS NULL');
    ExecuteSQL('UPDATE LISTEINVENT SET GIE_TYPEINVENTAIRE="INV" WHERE GIE_TYPEINVENTAIRE IS NULL');
    ExecuteSQL('UPDATE LISTEINVLIG SET GIL_STATUTDISPO="",GIL_STATUTFLUX="",GIL_LOTINTERNE=""'
       +',GIL_SERIEINTERNE="",GIL_TIERSPROP="",GIL_INDICEARTICLE="",GIL_MARQUE="", GIL_CHOIXQUALITE=""'
       +', GIL_REFAFFECTATION="",GIL_NUMEROSERIE=(SELECT GA_NUMEROSERIE FROM ARTICLE'
       +'  WHERE GIL_ARTICLE=GA_ARTICLE) WHERE GIL_REFAFFECTATION IS NULL');
    ExecuteSQL('DELETE FROM CHOIXCOD WHERE CC_TYPE="YTN" AND CC_CODE="P"');
    InsertChoixCode ('YTN','B','Des poids bruts','Poids bruts','');
    InsertChoixCode ('YTN','N','Des poids nets','Poids nets','');
  end;  // not IsDossierPCL
    // compta
  ExecuteSQL('UPDATE TIERSCOMPL SET YTC_PROFESSION="",YTC_REMUNERATION="",YTC_INDEMNITE="",YTC_AVANTAGE="",YTC_DAS2="",ytc_accelerateur="",ytc_schemagen=""');
End;


Procedure MajVer649;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  ExecuteSql ('Update ressource set ars_saisieact="-"');
  //GC
  ExecuteSql ('UPDATE PAYS SET PY_LIMITROPHE = "-"');
  //PAIE
  ExecuteSQL('UPDATE SALARIES SET PSA_NBREACQUISCP=0 WHERE PSA_NBREACQUISCP is null');
  ExecuteSQL('UPDATE SALARIES SET PSA_DADSNTIC="-"');
  ExecuteSQL('UPDATE COTISATION SET PCT_DADSEPARGNE="-"');
  ExecuteSQL('UPDATE INVESTFORMATION SET PIF_MTREALISEOPCA=0');
  // ajout 15092004
  AglNettoieListes('PGMULSALARIE', 'PSA_ETABLISSEMENT',nil);
  //compta
  ExecuteSQL('update journal set j_tresoimport="-",j_tresochainage="-",j_tresomontant="-",j_tresodate="-",j_tresolibelle="-"');
End;


Procedure MajVer650;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GPAO
    ExecuteSQL('UPDATE WORDREPHASE SET WOP_TRAITEMENT = "" WHERE WOP_TRAITEMENT IS NULL');
    ExecuteSQL('UPDATE WPDRTET SET WPE_VALEURPDR=0 WHERE WPE_VALEURPDR IS NULL');
    ExecuteSQL('UPDATE STKVALOPARAM SET GVP_NATURETRAVAIL="" WHERE GVP_NATURETRAVAIL IS NULL');
    ExecuteSQL('UPDATE STKVALOPARAM SET GVP_STKFLUX="" WHERE GVP_STKFLUX IS NULL');
    ExecuteSQL('UPDATE WCBNPREVTET SET wpt_affaire="" WHERE wpt_affaire IS NULL');
    ExecuteSQL('UPDATE STKMOUVEMENT SET GSM_NUMORDRE=0 WHERE GSM_NUMORDRE IS NULL');
    { Mise à jour du champ GL_ETATSOLDE }
    UpdateDecoupeLigne('GL_ETATSOLDE = iif(GL_ARTICLE = "", "", iif(GL_VIVANTE = "X", iif(GL_QTERESTE > 0, "ENC", "SOL"), "SOL"))');
    ExecuteSQL('UPDATE WORDRELIG SET WOL_MISEENPROD="" WHERE WOL_MISEENPROD is Null');
    ExecuteSQL('UPDATE WORDRELIG SET WOL_MISEENPROD=WOL_LIBREWOLA WHERE WOL_LIBREWOLA IN ("DEC", "ALL")');
    ExecuteSQL('UPDATE WORDRELIG SET WOL_LIBREWOLA="" WHERE WOL_MISEENPROD IN ("DEC", "ALL")');
    ExecuteSQL('UPDATE LIGNECOMPL '
             + 'SET GLC_MISEENPROD='
             + '(select isnull(WAN_MISEENPROD, "NON") FROM WARTNAT, LIGNE '
             + '  WHERE GL_ARTICLE=WAN_ARTICLE AND WAN_NATURETRAVAIL="FAB" AND GLC_NATUREPIECEG=GL_NATUREPIECEG '
             + '        AND GLC_SOUCHE=GL_SOUCHE AND GLC_NUMERO=GL_NUMERO AND GLC_INDICEG=GL_INDICEG '
             + '        AND GLC_NUMORDRE=GL_NUMORDRE)');       // modifié 17/08/2004 PCS

    ExecuteSQL('UPDATE WNOMELIG SET WNL_TYPELIEN="COM"');
    ExecuteSQL('UPDATE WORDREBES SET WOB_TYPELIEN="COM"');
    // déjà en 644 02/09/2004 GPMajSTKNature;
    //GRF
    ExecuteSQL('UPDATE PARACTIONS SET RPA_PRODUITPGI="GRC",RPA_TABLELIBREF1="",RPA_TABLELIBREF2="",RPA_TABLELIBREF3=""');
    ExecuteSQL('UPDATE ACTIONS SET RAC_PRODUITPGI="GRC",RAC_TABLELIBREF1="",RAC_TABLELIBREF2="",RAC_TABLELIBREF3=""');
    ExecuteSQL('UPDATE ACTIONSCHAINEES SET RCH_PRODUITPGI="GRC",RCH_TABLELIBRECHF1="",RCH_TABLELIBRECHF2="",RCH_TABLELIBRECHF3=""');
    ExecuteSQL('UPDATE PARCHAINAGES SET RPG_PRODUITPGI="GRC",RPG_TABLELIBRECHF1="",RPG_TABLELIBRECHF2="",RPG_TABLELIBRECHF3=""');
    ExecuteSQL('UPDATE ACTIONSGENERIQUES SET RAG_PRODUITPGI="GRC",RAG_TABLELIBREF1="",RAG_TABLELIBREF2="",RAG_TABLELIBREF3=""');
    ExecuteSQL('UPDATE CHAINAGEPIECES SET RLC_PRODUITPGI="GRC"');
    ExecuteSQL('UPDATE OPERATIONS SET ROP_PRODUITPGI="GRC",ROP_OBJETOPEF=""');
    ExecuteSQL('UPDATE PROSPECTCONF SET RTC_PRODUITPGI="GRC"');

    //GIGA
    ExecuteSql ('Update ACTIVITE set ACT_NUMORDREACH=0');
    // pour forcer les natures de piéces affaire paramsoc (pbm vielles bases)
    if not (ExisteSQL('SELECT AFF_AFFAIRE0 FROM AFFAIRE')) then
    begin
      ExecuteSQL('UPDATE paramsoc  SET SOC_DATA="AFF" where soc_nom = "so_afnataffaire"');
      ExecuteSQL('UPDATE paramsoc  SET SOC_DATA="PAF" where soc_nom = "so_afnatproposition"');
    end;
  end;  // not IsDossierPCL

  //GRF
  RT_InsertLibelleTablettesFour;


  // PAie
  ExecuteSQL ('update paieencours set PPU_CBASESSPRAT = (SELECT SUM(PHC_MONTANT) FROM HISTOCUMSAL WHERE PHC_CUMULPAIE="40" AND PHC_SALARIE=PPU_SALARIE AND PPU_DATEDEBUT=PHC_DATEDEBUT AND PPU_DATEFIN=PHC_DATEFIN AND PPU_ETABLISSEMENT=PHC_ETABLISSEMENT)');
  ExecuteSQL ('UPDATE CONTRATTRAVAIL SET PCI_MOTIFSORTIE=""');
  ExecuteSQL ('UPDATE CONTRATTRAVAIL SET PCI_LIBELLEEMPLOI=(SELECT PSA_LIBELLEEMPLOI FROM SALARIES WHERE PSA_SALARIE=PCI_SALARIE)');
  ExecuteSQL ('UPDATE STAGE SET PST_NIVEAUFORMINIT="",PST_TYPEFORMINIT="",PST_DOMFORMINIT=""');
  ExecuteSQL ('UPDATE VISITEMEDTRAV SET PVM_APTESR="-"');
End;


Procedure MajVer651;
var v : variant;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  v := GetParamSoc('SO_CPRDDATERECEPTION');
  if VarIsNull(v) then v := iDate1900;
  //DP
  ExecuteSQL('UPDATE DPFISCAL SET DFI_TAXEPROF = "X", DFI_CONTREVENUSLOC = "-"'
                     +', DFI_TAXEFONCIERE = "-", DFI_CONTSOCSOLDOC = "-"'
                     +',DFI_TAXEGRDSURF = "-", DFI_TAXEANNIMM = "-", DFI_TAXEVEHICSOC = "-"'
                     +', DFI_VIGNETTEAUTO = "-",DFI_IMPSOLFORTUNE = "-"');
  ExecuteSQL('UPDATE DPSOCIAL SET DSO_TAXESALARIE = "-", DSO_TXSALPERIODIC = "-", DSO_TAXEAPPRENT = "-"'
      +', DSO_PARTFORMCONT = "-",DSO_PARTCONSTRUC = "-", DSO_PAIECAB = "-", DSO_PAIEENT = "-"'
      +', DSO_PAIEENTSYS = "-",DSO_REGPERS = "-", DSO_DECUNEMB = "-", DSO_ELTVARIAENT = "-"'
      +', DSO_DATEDERPAIE = "'+UsDateTime(iDate1900)+'",DSO_GESTCONGES = "-", DSO_MUTSOCAGR = "-", DSO_INTERMSPEC = "-"'
      +', DSO_BTP = "-",DSO_TICKETREST = "-", DSO_GESTIONETS = "-", DSO_TELEDADS = "-", DSO_PLANPAIEACT = "-"'
      +', DSO_CDD = "-", DSO_CDI = "-", DSO_ABATFRAISPRO = "-", DSO_TPSPARTIEL = "-"'
      +', DSO_TPSPARTIEL30 = "-", DSO_CIE = "-", DSO_CEC = "-", DSO_CES = "-",DSO_CRE = "-", DSO_EMBSAL1 = "-"'
      +', DSO_EMBSAL23 = "-", DSO_CONTAPPRENT = "-",DSO_CONTQUAL = "-", DSO_CONTORIENT = "-"'
      +', DSO_EXOCHARGES = "-", DSO_DATEEXSOC = "'+UsDateTime(iDate1900)+'",DSO_COMITEENT = "-", DSO_DELEGUEPERS = "-"'
      +', DSO_DELEGUESYND = "-", DSO_EXISTTICKREST = "-",DSO_TAUXACCTRAV = 0, DSO_VERSETRANS = "-"');


  // DESGOUTTE 23092004  ExecuteSQL('update DPDOCUMENT set DPD_NEPUBLIABLE="X", DPD_NEAPUBLIER="-", DPD_CLAID=0, DPD_NEDATEPUBL="", DPD_NEPUBLIEUR=""');
  // DESGOUTTE 23092004   DOS_CPLIENGAMME="", DOS_CPRDDATERECEP="'+UsDateTime(iDate1900)+'", DOS_CPRDREPERTOIRE=""
  ExecuteSQL('update DOSSIER set DOS_NECDKEY=""');  // , DOS_CPPOINTAGESX="AUC" supprimé en 662
  // DESGOUTTE 23092004 ExecuteSQL('update YGEDDICO set YGD_CLAID=0');
  // remontée de paramsoc compta dans la base commune
 // DESGOUTTE 23092004
 { ExecuteSql('update ##DP##.DOSSIER set DOS_CPLIENGAMME="' + GetParamSoc('SO_CPLIENGAMME') + '",'
   +' DOS_CPRDREPERTOIRE="' + GetParamSoc('SO_CPRDREPERTOIRE') + '",'
   +' DOS_CPRDDATERECEP="' +  UsDateTime(v)  + '"'
   +' WHERE DOS_NODOSSIER="'+V_PGI.NoDossier+'"');
 }
  ExecuteSQL('UPDATE DPTABCOMPTA SET DTC_NBENTREEIMMO=0, DTC_NBSORTIEIMMO=0, DTC_NBLIGNEIMMO=0');
  //GIGA
  ExecuteSQL('update paramsoc set soc_data=(select soc_data from paramsoc where soc_nom="so_afvaloactpr") where soc_nom = "so_afvalofraispr"');
  ExecuteSQL('update paramsoc set soc_data=(select soc_data from paramsoc where soc_nom="so_afvaloactpr") where soc_nom = "so_afvalofourpr"');
  ExecuteSQL('update paramsoc set soc_data=(select soc_data from paramsoc where soc_nom="so_afvalopv") where soc_nom = "so_afvalofourpv"');
  ExecuteSQL('update paramsoc set soc_data=(select soc_data from paramsoc where soc_nom="so_afvalopv") where soc_nom = "so_afvalofraispv"');

  if not IsDossierPCL then
  begin
    //GIGA
    ExecuteSQL('update afcumul set acu_errcpta=0, acu_familletaxe="",acu_regimetaxe="",acu_etablissement=""');

    if Vers_Piece < 164 then  UpdateDecoupePiece('GP_DATEDEBUTFAC =GP_DATEPIECE,GP_DATEFINFAC =GP_DATEPIECE');
    //GESCOM
    ExecuteSQL('UPDATE ACOMPTES SET GAC_MONTANTTRA=0, GAC_MONTANTTRADEV=0, GAC_PIECEPRECEDENTE="" ');
    ExecuteSQL('update article set ga_calculpa="-" ');
  end;  // not IsDossierPCL
  //compta
  ExecuteSQL('update IMMOLOG SET IL_DATEOPREELLE=IL_DATEOP ');
  ExecuteSQL('UPDATE GENERAUX SET G_IAS14 = "-" ');
  AglNettoieListes('CPMODIFENTPIE', 'E_NUMGROUPEECR', nil);
End;

Procedure MajVer652;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //Compta
  ExecuteSQL('UPDATE IMMO SET I_INFOUO=""');
  if not IsDossierPCL then
  begin
    //GC
    ExecuteSQL('UPDATE OPERCAISSE SET GOC_NUMZCAISSE=0, GOC_CAISSE="", GOC_ANTERIEUR="-"');
    //GP
    ExecuteSQL('update QSATTRIB set qsb_affaire="", qsb_ordfabpere="" where qsb_affaire is null');
    ExecuteSQL('update QSTKDISPO set qsd_affaire="", qsd_ordfabpere="" where qsd_affaire is null');
    ExecuteSQL('update QHISTOAFF set qha_affaire="", qha_ordfabpere="" where qha_affaire is null');
    ExecuteSQL('update WCBNEVOLUTION set wev_ordfabpere="" where wev_ordfabpere is null');


  //GC   (champs créés en 639)
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="SAC", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="AFS"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="EVE", GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG="AVS"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="SAC", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="BFA"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="SVE", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="BLC"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="EAC", GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG="BLF"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="EVE", GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG="BRC"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="EAC", GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG="BSA"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="RVE", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="CC"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="RVE", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="CCE"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="AAC", GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG="CF"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="AAC", GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG="CSA"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="SVE", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="FAC"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="AAC", GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG="FCF"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="EAC", GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG="FF"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="SVE", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="FFO"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="SVE", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="LCE"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="RPL", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="PRE"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="STR", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="TEM"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="ETR", GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG="TRE"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="TAV", GPP_SENSPIECE="MIX"  WHERE GPP_NATUREPIECEG="TRV"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="SAC", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="AFP"');
  end;  // not IsDossierPCL

  //DP
  // DESGOUTTE 23092004  ExecuteSQL('UPDATE DOSSIER SET DOS_NEDOSACREER="-", DOS_NEGEDACREER="-"');
End;

Procedure MajVer653;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //DP
  ExecuteSQL('UPDATE DOSSIER SET DOS_PWDGLOBAL="X" WHERE DOS_APPLISPROTEC=""');
  //GIGA
   // GI/GA
  InsertChoixCode('AAO', 'AGE', 'Activité réalisé de l''agenda', 'Agenda', '');
  InsertChoixCode('AAO', 'APL', 'Planning dans l''activité', 'Planning', '');
  InsertChoixCode('AAO', 'SAI', 'Saisie', 'Saisie', '');
  InsertChoixCode('AAO', 'SDE', 'Saisie décentralisée', 'SDA', '');

End;

Procedure MajVer654;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    // JTR - Mise à jour des nouveaux évènements PCP pour SocRef 654
    ExecuteSQL('DELETE FROM STOXEVENTS WHERE SEV_CODEEVENT = "PCP_DEMANDE"');
    ExecuteSQL('DELETE FROM STOXEVENTS WHERE SEV_CODEEVENT = "PCP_GENERIQUE"');
    ExecuteSQL('DELETE FROM STOXEVENTS WHERE SEV_CODEEVENT = "PCP_INIT_ACH"');
    ExecuteSQL('DELETE FROM STOXEVENTS WHERE SEV_CODEEVENT = "PCP_INIT_VTE"');
    ExecuteSQL('DELETE FROM STOXEVENTS WHERE SEV_CODEEVENT = "PCP_MOD_ACH"');
    ExecuteSQL('DELETE FROM STOXEVENTS WHERE SEV_CODEEVENT = "PCP_MOD_VTE"');
    ExecuteSQL('DELETE FROM STOXEVENTS WHERE SEV_CODEEVENT = "PCP_REMONTEE_ACH"');
    ExecuteSQL('DELETE FROM STOXEVENTS WHERE SEV_CODEEVENT = "PCP_REMONTEE_VTE"');
    MajEvtPCP_V654;
    // Fin JTR
    // GC initialisation des nouveaux stocks
    if (V_PGI.driver <> dbORACLE7) and not (V_PGI.ModePCL='1') then      // pas si PCL
    begin
      if V_PGI.SAV then LogAGL('Début Init Stock ' + DateTimeToStr(Now));
      if StkMoulinetteCanStart then CallSTKMoulinette(True);
      if V_PGI.SAV then LogAGL('Fin Init Stock ' + DateTimeToStr(Now));
    end;

    //MM 23092004 GPMajSTKValoParam; // ajout param de valo stock si table vide

    // GC initialisation des nouveaux Tarifs
    if (V_PGI.driver <> dbORACLE7) then
    begin
      if V_PGI.SAV then LogAGL('Début moulinette tarifs ' + DateTimeToStr(Now));
      CallTarifsMoulinette(True);
      if V_PGI.SAV then LogAGL('Fin moulinette tarifs ' + DateTimeToStr(Now));
    end;
  end;  // not IsDossierPCL

End;

Procedure MajVer655 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSQL('UPDATE ARTICLE SET GA_UNITEQTEVTE = IIF(GA_UNITEQTEVTE = "", GA_QUALIFUNITESTO, GA_UNITEQTEVTE) '
      + ',GA_QUALIFUNITEVTE = IIF(GA_QUALIFUNITEVTE = "", GA_QUALIFUNITESTO, GA_QUALIFUNITEVTE) '
      + ',GA_UNITEQTEACH = IIF(GA_UNITEQTEACH = "", GA_QUALIFUNITESTO, GA_UNITEQTEACH) '
      + ',GA_UNITEPROD = IIF(GA_UNITEPROD  = "", GA_QUALIFUNITESTO, GA_UNITEPROD) '
      + ',GA_UNITECONSO = IIF(GA_UNITECONSO  = "", GA_QUALIFUNITESTO, GA_UNITECONSO) '
      + 'WHERE (GA_QUALIFUNITESTO <> "") AND ((GA_UNITEQTEVTE = "") OR (GA_QUALIFUNITEVTE = "") '
      + 'OR (GA_UNITEQTEACH="") OR (GA_UNITEPROD = "") OR (GA_UNITECONSO = ""))');
  end;  // not IsDossierPCL
End;


Procedure MajVer656 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //GC
  AglNettoieListes('GCMULLIGNE;GCMULLIGNEACH;GCMULARTCOM;WARTICLEPIECES', 'GL_NUMORDRE',nil);
  //GRC
  ExecuteSQL('delete from CHOIXEXT WHERE YX_TYPE like "LB%" and len(yx_code) > 3');
  //GI
  AglNettoieListes('AFMULLIGNE1', 'GL_NUMORDRE', nil);
End;

Procedure MajVer657 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //GRC
  ExecuteSQL('UPDATE CONTACT SET C_CLETELEPHONE="" where C_CLETELEPHONE is null');
  if not IsDossierPCL then
  begin
    //GP
    ExecuteSQL('UPDATE PARPIECE SET GPP_REFINTEXT="INT", GPP_MAJINFOTIERS="X" WHERE GPP_NATUREPIECEG="BSP"');
  end;  // not IsDossierPCL
End;

Procedure MajVer658 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //GC
  AglNettoieListes('GCMULPEC', 'GL_NUMORDRE',nil);

  if not IsDossierPCL then
  begin
    //GPAO
    ExecuteSQL('UPDATE PARPIECE SET GPP_ACTIONFINI="TRA",GPP_TAXE="X",GPP_NATURETIERS="FOU;",GPP_APPELPRIX="PAS",GPP_TYPEARTICLE="MAR;"'
       +  ',GPP_MAJPRIXVALO="DPA;DPR;PPA;PPR;",GPP_PRIORECHART1="ART",GPP_DIMSAISIE="TOU",GPP_IFL1="004",GPP_IFL2="002",'
       +  'GPP_IFL3="011", GPP_GEREECHEANCE="DEM", GPP_COMPANALLIGNE="SAN",GPP_COMPANALPIED="SAN" WHERE GPP_NATUREPIECEG = "BSP"');

    InsertChoixCode('WMA','AFF', 'A l''affaire',' A l''affaire','');
    //ExecuteSQL(‘UPDATE RESSOURCE SET ARS_DEPOT= “‘+VH_GC.GCDepotDefaut+’“ WHERE ARS_DEPOT=  “” ‘);
    ExecuteSQL('UPDATE RESSOURCE SET ARS_DEPOT=(Select SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_GCDEPOTDEFAUT") WHERE ARS_DEPOT=""');
    W_SAV_COMPETENCE_UPDATE;
    W_SAV_COMPETENCE;
    W_SAV_INTER_COMPL;
    //--------- WCBNTYPEMOUV -----------------------
    if not ExisteSQL('select 1 from wcbntypemouv where wtm_typemouvement = "PLC"') then
    begin
      ExecuteSQL('insert into WCBNTYPEMOUV (WTM_TYPEMOUVEMENT,WTM_LIBELLE) values ("PLC","Plan commercial")');
    end;

    if not ExisteSQL('select 1 from wcbntypemouv where wtm_typemouvement = "DEV"') then
    begin
      ExecuteSQL('insert into WCBNTYPEMOUV (WTM_TYPEMOUVEMENT,WTM_LIBELLE) values ("DEV","Devis")');
    end;
  end;  // not IsDossierPCL

  //GRC
  InsertChoixCode('R1Z', 'BLO', 'Bloc-notes ', '', '');
  //GIGA
  AglNettoieListes('AFMULFACTAFFAPP', 'AFF_REGSURCAF',nil);
End;

Procedure MajVer659 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    // GPAO
    ExecuteSQL('UPDATE DISPO SET GQ_UNITEMIN="1",GQ_UNITEMAX="1",GQ_UNITEALERTE="1",GQ_UNITERECOMPL="1" WHERE (GQ_UNITEMIN IS NULL) OR (GQ_UNITEMIN ="")');
    ExecuteSQL('UPDATE PARPIECE SET GPP_SOUCHE="BSP",GPP_LISTESAISIE="GCSAISIEBSP",GPP_RELIQUAT="X" WHERE GPP_NATUREPIECEG = "BSP"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_SOUCHE="CSP",GPP_LISTESAISIE="GCSAISIECSP",GPP_RELIQUAT="X" WHERE GPP_NATUREPIECEG = "CSP"');
    UpdateDecoupeLigne('gl_typenomenc="ASC"',
         ' AND gl_typearticle="NOM" and gl_typenomenc="ASS" and gl_tenuestock="-" and gl_naturepieceg not in ("FF", "BLF", "CF", "AF", "AFP", "AFS", "BFA", "BSA", "BSP", "CFR", "CSA", "CSP", "DEF", "EEX", "SEX", "FCF", "FRF", "LFR", "REA", "TEM", "TRE")');
    //GC
    AglNettoieListes('LSTOXQUERYS', 'SQE_TYPETRT',nil);
    AglNettoieListes('LSTOXCHRONO', 'SXC_CODESITE',nil);
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="" WHERE GPP_NATUREPIECEG = "BSP"');
  end;  // not IsDossierPCL
End;

Procedure MajVer660 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    // GPAO
    ExecuteSql('Update stknature set GSN_SIGNEMVT = "SOR" where gsn_StkTypeMvt = "RES"');
    ExecuteSQL('update WCBNTYPEMOUV set WTM_LIBELLE="Besoin / vente"  where WTM_TYPEMOUVEMENT="VEN"');
    ExecuteSQL('update DISPO set GQ_UNITEMIN="1",GQ_UNITEMAX="1",GQ_UNITEALERTE="1",GQ_UNITERECOMPL="1" where (GQ_UNITEMIN is null) or (GQ_UNITEMIN ="") or (GQ_UNITEMIN ="QTE")');
    ExecuteSQL('update PARPIECE set GPP_PILOTEORDRE="X", GPP_IMPBESOIN="X", GPP_IMPAUTOETATCBN="-"  where GPP_NATUREPIECEG="CSA"');
    ExecuteSQL('update PARPIECE set GPP_PILOTEORDRE="X", GPP_IMPBESOIN="-", GPP_IMPAUTOETATCBN="-"  where GPP_NATUREPIECEG="BSA"');

    { Initialisation de la table des types de prix de revient utilisée notamment dans STKVALOTYPE }
    GPMajPdrType;
    { Initialisation des deux tables STKVALOTYPE et STKVALOPARAM pour un fonctionnement standard }
    GPMajSTKValoParam;
    { Réinitialisation des N° de jetons notamment pour les tables STKVALOTYPE et STKVALOPARAM }
    {wResetWJT;}
  end;  // not IsDossierPCL
  // PAie
  AglNettoieListes('PGEMULMVTABS', 'PCN_SALARIE;PCN_ORDRE;PCN_TYPEMVT',nil);
  AglNettoieListes('PGEMULMVTABSR', 'PCN_SALARIE;PCN_ORDRE;PCN_TYPEMVT',nil);
  //DP
  ExecuteSQL('update YGEDDICO set YGD_CDPUBLI="", YGD_EWSID=""');
  ExecuteSQL('update DPDOCUMENT set DPD_EWSID="", DPD_EWSPUBLIABLE="X", DPD_EWSAPUBLIER="-", DPD_EWSDATEPUBL="'+UsDateTime(iDate1900)+'", DPD_EWSPUBLIEUR=""');
  ExecuteSQL('update DOSSIER set DOS_USRS1="", DOS_PWDS1="", DOS_EWSCREE="-"');
	// supprimé le 30092004 ExecuteSQL('update JQ6_fonction set jq6_formeq = ""');
	ExecuteSQL('update histoannulien set hnl_typemodif = ""');
  if IsMonoOuCommune then
     ExecuteSQL('insert into userconf(uco_groupeconf, uco_user) select cc_code, us_utilisateur '
        +' from choixcod, utilisat where cc_type="UCO" and us_superviseur="X" '
        +' and not exists(select 1 from userconf where uco_user=us_utilisateur '
        +' and uco_groupeconf=cc_code) group by cc_code, us_utilisateur');
End;

Procedure MajVer661 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GPAO
     ExecuteSQL('UPDATE STKNATURE SET GSN_QTEPLUS = "GQ_PHYSIQUE;" WHERE GSN_QUALIFMVT = "SEX"') ;
  end;  // not IsDossierPCL
  //TRESO
   // AglNettoieListes('TRECRITURERAPPRO', 'TE_DATECOMPTABLE;TE_NUMTRANSAC;TE_NUMEROPIECE;TE_CODEFLUX;TE_CODEBANQUE;TE_EXERCICE;BQ_LIBELLE;TE_JOURNAL;TE_MONTANT;TE_DEVISE;TE_MONTANTDEV;TE_DATERAPPRO;TE_DATEVALEUR;TE_NUMLIGNE;TE_CLEVALEUR;TE_NATURE;TE_GENERAL;TE_CPNUMLIGNE;TE_NUMECHE;');
   ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA = "-" WHERE SOC_NOM = "SO_PREMIERESYNCHRO" ');
   //COMPTA
   AglNettoieListes('MULMIMMOS', 'I_GROUPEIMMO',nil);
   ExecuteSQL('update journal set j_accelerateur="X" where j_naturejal="ACH" or j_naturejal="VTE" ');
   //DP
   ExecuteSQL('UPDATE DOSSIER SET DOS_SERIAS1="X"');
End;

Procedure MajVer662 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GC
    ExecuteSQL('UPDATE PARPIECE SET GPP_ARTSTOCK="-" WHERE GPP_NATUREPIECEG="BFA"');
    //GPAO
    ExecuteSQL('UPDATE PARPIECE SET GPP_LIBELLE="Commande SST d''achat" WHERE GPP_NATUREPIECEG="CSA"') ;
    if not ExisteSQL('SELECT 1 FROM WCBNTYPEMOUV WHERE WTM_TYPEMOUVEMENT = "BOS"') then
      ExecuteSQL('INSERT INTO WCBNTYPEMOUV (WTM_TYPEMOUVEMENT,WTM_LIBELLE) VALUES ("BOS","Besoin / ordre de sst d''achat")');
    if not ExisteSQL('SELECT 1 FROM WCBNTYPEMOUV WHERE WTM_TYPEMOUVEMENT = "BPS"') then
      ExecuteSQL('INSERT INTO WCBNTYPEMOUV (WTM_TYPEMOUVEMENT,WTM_LIBELLE) VALUES ("BPS","Besoin / proposition de sst d''achat")');
    ExecuteSQL('UPDATE WCBNTYPEMOUV SET WTM_LIBELLE="Proposition de sst d''achat" WHERE WTM_TYPEMOUVEMENT="PRS"');
  end;
End;

Procedure MajVer663 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    GPMajxWizards;
		// GIGA
    ExecuteSql('UPDATE PARPIECE SET GPP_LISTESAISIE="AFSAISIEFAC" WHERE GPP_NATUREPIECEG="AFF"  AND  GPP_LISTESAISIE="GCSAISIECC"');
    ExecuteSql('UPDATE PARPIECE SET GPP_LISTESAISIE="AFSAISIEFAC" WHERE GPP_NATUREPIECEG="PAF"  AND  GPP_LISTESAISIE="GCSAISIEDEC"');
    AglNettoieListes('AFMULRECHAFFAIRE', 'AFF_TIERS',nil);
    //  Ajout pour 6.01
    ExecuteSql('Update stknature set gsn_signemvt="MIX" where gsn_qualifmvt in ("EVE", "SAC")');
  end;
    // nouveaux concepts pour présentations et filtres initialisés aux mêmes droits que les concepts existants
  ExecuteSQL('update menu set mn_ACCESGRP=(select MAX(mn_accesgrp) from menu MN where MN.mn_tag=26000 and MN.mn_1 =26) '
    +' where (mn_1 = 26) and mn_tag = 26004 ');
  ExecuteSQL('update menu set mn_ACCESGRP=(select MAX(mn_accesgrp) from menu MN where MN.mn_tag=26001 and MN.mn_1 =26) '
    +' where (mn_1 = 26) and mn_tag = 26003 ');
   //COMPTA
   ExecuteSql('UPDATE BANQUES SET PQ_ETABBQ="" where PQ_ETABBQ is null');
   ExecuteSql('UPDATE BANQUES SET PQ_ABREGE="" where PQ_ABREGE is null');
   AglNettoieListes('CPEEXBQ','EE_ORIGINERELEVE;EE_NUMERO');
   AglNettoieListes('CPEEXBQ2','EE_ORIGINERELEVE;EE_NUMERO');
   AglNettoieListes('CPRELANCECLIENT','E_BLOCNOTE');
End;

Procedure MajVer667 ;
var iInd : integer;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GC
    for iInd := 1 to 8 do
      ExecuteSQL('UPDATE PARPIECE SET GPP_IFL' + intToStr(iInd) + '=IIF(GPP_IFL'
        + intToStr(iInd) + '="411","015",IIF(GPP_IFL' + intToStr(iInd) + '="400","001",IIF(GPP_IFL'
        + intToStr(iInd) + '="409","002",IIF(GPP_IFL'
        + intToStr(iInd) + '="404","010",GPP_IFL' + intToStr(iInd) + '))))');
    //GPAO
    ExecuteSQL('UPDATE PARPIECE SET GPP_IMPMODELE = "", GPP_IMPETAT= "BSP" WHERE GPP_NATUREPIECEG="BSP"');
    ExecuteSQL('UPDATE EDILIGNE SET ELI_PUTTCDEV=0,ELI_REMISELIBRE=0,ELI_REMISELIGNE=0,'
             + ' ELI_REMISEPIED=0,ELI_VALEURFIXEDEV=0,ELI_VALEURREMDEV=0'
             + ' WHERE ELI_PUTTCDEV is Null AND ELI_REMISELIBRE is Null'
             + ' AND ELI_REMISELIGNE is Null AND ELI_REMISEPIED is Null'
             + ' AND ELI_VALEURFIXEDEV is Null AND ELI_VALEURREMDEV is Null');
    if Not ExisteSQL('SELECT WDB_QUALIFIANT FROM WCUQUALIFIANT WHERE WDB_QUALIFIANT="QAF"') then
        ExecuteSQL('INSERT INTO WCUQUALIFIANT VALUES ("QAF", "Quantité affectée", "*")');

    if Not ExisteSQL('SELECT WDX_CONTEXTE FROM WFORMULECHAMP WHERE WDX_CONTEXTE="WOB" AND WDX_QUALIFIANT="QAF"') then
        ExecuteSQL('INSERT INTO WFORMULECHAMP'
              +' VALUES ("WOB","QAF","WOB_QAFFSAIS","WOB_UNITELIEN","WOB_COEFLIEN","WOB_QAFFSTOC","WOB_QUALIFUNITESTO","'+ UsDateTime(V_PGI.DateEntree)+'", "'+ UsDateTime(V_PGI.DateEntree)  + '","' + V_PGI.User + '","' + V_PGI.User + '")');
    AglNettoieListes('EDIPIECE','', nil);
    AglNettoieListes('EDILIGNE','ELI_ARTICLE', nil);
    AglNettoieListes('WORDRECMP','GL_DATEDEPARTUSINE', nil);
    AglNettoieListes('UORDREBES','WOB_TYPELIEN',nil);
    AglNettoieListes('YTARIFSMULART201', 'YTS_SITE;YTS_ARTICLE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_ACTIF;YTS_LIBELLETARIF;YTS_DATEREFERENCE;YTS_TTCOUHT;YTS_CASCCONTEXTE', nil);
    AglNettoieListes('YTARIFSMULART211', 'YTS_SITE;YTS_ARTICLE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_ACTIF;YTS_LIBELLETARIF;YTS_DATEREFERENCE;YTS_TTCOUHT;YTS_CASCCONTEXTE', nil);
    AglNettoieListes('YTARIFSMULART301', 'YTS_SITE;YTS_ARTICLE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_ACTIF;YTS_LIBELLETARIF;YTS_DATEREFERENCE;YTS_TTCOUHT;YTS_CASCCONTEXTE', nil);
    AglNettoieListes('YTARIFSMULART401', 'YTS_SITE;YTS_ARTICLE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_ACTIF;YTS_LIBELLETARIF;YTS_DATEREFERENCE;YTS_TTCOUHT;YTS_CASCCONTEXTE', nil);
    AglNettoieListes('YTARIFSMULTIE101', 'YTS_SITE;YTS_ARTICLE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_ACTIF;YTS_LIBELLETARIF;YTS_DATEREFERENCE;YTS_TTCOUHT;YTS_CASCCONTEXTE', nil);
    AglNettoieListes('YTARIFSMULTIE201', 'YTS_SITE;YTS_ARTICLE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_ACTIF;YTS_LIBELLETARIF;YTS_DATEREFERENCE;YTS_TTCOUHT;YTS_CASCCONTEXTE', nil);
    AglNettoieListes('YTARIFSMULTIE211', 'YTS_SITE;YTS_ARTICLE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_ACTIF;YTS_LIBELLETARIF;YTS_DATEREFERENCE;YTS_TTCOUHT;YTS_CASCCONTEXTE', nil);
    AglNettoieListes('YTARIFSMULTIE301', 'YTS_SITE;YTS_ARTICLE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_ACTIF;YTS_LIBELLETARIF;YTS_DATEREFERENCE;YTS_TTCOUHT;YTS_CASCCONTEXTE', nil);
    AglNettoieListes('YTARIFSMULTIE401', 'YTS_SITE;YTS_ARTICLE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_ACTIF;YTS_LIBELLETARIF;YTS_DATEREFERENCE;YTS_TTCOUHT;YTS_CASCCONTEXTE', nil);
    AglNettoieListes('WPDRTET', 'WPE_TYPEPDR;WPE_NATURETRAVAIL;WPE_PERIODESAUV;'
                                  +'WPE_CODITI;WPE_CIRCUIT'
                                  +'WPE_ARTICLE;WPE_CODEARTICLE;'
                                  +'WPE_ARTICLEWNT;WPE_CODEARTICLEWNT,WPE_MAJEURWNT,WPE_MINEURWNT,WPE_ETATREVWNT;'
                                  +'WPE_ARTICLEWGT;WPE_CODEARTICLEWGT,WPE_MAJEURWGT,WPE_MINEURWGT,WPE_ETATREVWGT;'
                                 , nil);
  End;

  //Treso
  // J.PASTERIS 17012005 ExecuteSQL('UPDATE TRECRITURE SET TE_COMMISSION = "S"');
  // J.PASTERIS 17012005 ExecuteSQL('UPDATE FLUXTRESO SET TFT_CLASSEFLUX = "FIN" WHERE TFT_PREVISIONNEL = "-" OR TFT_PREVISIONNEL = "" OR TFT_PREVISIONNEL IS NULL');
  // J.PASTERIS 17012005 ExecuteSQL('UPDATE FLUXTRESO SET TFT_CLASSEFLUX = "PRE" WHERE TFT_PREVISIONNEL = "X"');
  // J.PASTERIS 17012005 ExecuteSQL('UPDATE FLUXTRESO SET TFT_CLASSEFLUX = "REF" WHERE TFT_FLUX IN ("EQR", "EQD", "REI")');
  //compta
  AglNettoieListes('MULVMVTS','E_EXERCICE;E_NUMECHE;E_QUALIFPIECE');
  AglNettoieListes('MULMMVTS','E_EXERCICE;E_NUMECHE;E_QUALIFPIECE');

End;

Procedure MajVer668 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
     //giGA
    ExecuteSql ('UPDATE TACHE SET ATA_PAVANCEMENT=0.0');
    executeSql ('Update afplanningParam set app_libelleparam="Accès par affaire / planning des tâches" where app_typeafparam="PLA" and app_codeparam="P1J"');
    executeSql ('Update afplanningParam set app_libelleparam="Accès par ressource / planning des tâches" where app_typeafparam="PLA" and app_codeparam="P2J"');
    executeSql ('Update afplanningParam set app_libelleparam="Accès par affaire / planning des ressources" where app_typeafparam="PLA" and app_codeparam="P3J"');
    executeSql ('Update afplanningParam set app_libelleparam="Accès par ressource / planning des affaires" where app_typeafparam="PLA" and app_codeparam="P4J"');
  end;
End;


Procedure MajVer669 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GC
    AglNettoieListes('GCMULARCHIVAGEMVT','',nil);
    ExecuteSql ('UPDATE CODECPTA SET GCP_CPTEGENESCACH="", GCP_CPTEGENESCVTE="", GCP_CPTEGENREMACH="", GCP_CPTEGENREMVTE=""');
    ExecuteSql ('UPDATE STKNATURE SET GSN_GERECOMPTA="-"');
    ExecuteSql ('UPDATE STKMOUVEMENT SET GSM_COMPTABILISE="-"');
  end;
End;

Procedure MajVer670 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSQL('UPDATE STKMOUVEMENT SET GSM_QRUPTURE = 0.0');
  end;
End;

Procedure MajVer671 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GPAO
    ExecuteSql('update stkmouvement set GSM_DPA=0.0, GSM_DPR=0.0, GSM_PMAP=0.0, GSM_PMRP=0.0, GSM_PRIXSAISIS="X" where GSM_PRIXSAISIS is null');
    AglNettoieListes('GCSTKDISPO', 'GA_QUALIFUNITEVTE', nil);
    ExecuteSQL('UPDATE STKMOUVEMENT SET GSM_ETATTRANSFERT="" WHERE GSM_ETATTRANSFERT IS NULL');
  end;
End;

Procedure MajVer673 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    // GI/Ga
    Executesql ('Update tache set ata_termine="-"');
    Executesql ('Update parpiece set gpp_naturereprise=""');
    // GPAO
    AglNettoieListes('WCOMPAREWGL','WGL_OPEITI',nil);
    AGLNettoieLIstes('WORDRELIG','WOL_MISEENPROD',nil);
    //GC
    Executesql ('Update PARPIECE set GPP_modeecheances="RS" where GPP_modeecheances is null');
    Executesql ('Update PARPIECE set GPP_naturereprise="" where GPP_naturereprise is null');
  end;
  //TOX
  // XP 22-12-2004 : Les données sont maintenant dans la table COMMUN, un  peu de ménage ..
  ExecuteSql ('DELETE FROM CHOIXCOD WHERE CC_TYPE="SME" OR CC_TYPE="STH" OR CC_TYPE="STT"') ;
End;

Procedure MajVer674 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GRC
    ExecuteSql ('update parpiece set gpp_infoscompl="-" ') ;
    ExecuteSql ('update rtinfosdesc set rde_chporigine="" ') ;
    ExecuteSql ('update champspro set rcl_chporigine="" ') ;
    RT_InsertLibelleInfoComplArticle;
    RT_InsertLibelleInfoComplLigne;
    ExecuteSql ('update paractions set rpa_weekend = "-"') ;
  end;
End;

Procedure MajVer675 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GRC
    AglNettoieListes('RTMULACTREPORT', 'RAC_DATEECHEANCE;RAC_CHAINAGE;RAC_NUMLIGNE',nil);
    //GPAO
    ExecuteSQL('update TRANSINVENT set GIT_CODELISTE="" where GIT_CODELISTE is null');
    ExecuteSQL('UPDATE DEPOTS SET GDE_STKTRFEMPLACE="", GDE_STKTRFNOMINV="", GDE_STKTRFNOMENT="", GDE_STKTRFNOMSOR="" WHERE GDE_STKTRFEMPLACE IS NULL');
    if not ExisteSQL('select 1 from STKNATURE where GSN_QUALIFMVT="EIT"') then
    begin
      ExecuteSQL('INSERT INTO STKNATURE (GSN_QUALIFMVT,GSN_LIBELLE,GSN_STKTYPEMVT,GSN_QTEPLUS,GSN_QUALIFMVTSUIV,GSN_SIGNEMVT,GSN_STKFLUX,GSN_GERECOMPTA) VALUES ("EIT","Entrée d''inventaire : transtockeur","PHY","GQ_PHYSIQUE;","","ENT","STO","-")');
    end;
    if not ExisteSQL('select 1 from STKNATURE where GSN_QUALIFMVT="SIT"') then
    begin
      ExecuteSQL('INSERT INTO STKNATURE (GSN_QUALIFMVT,GSN_LIBELLE,GSN_STKTYPEMVT,GSN_QTEPLUS,GSN_QUALIFMVTSUIV,GSN_SIGNEMVT,GSN_STKFLUX,GSN_GERECOMPTA) VALUES ("SIT","Sortie d''inventaire : transtockeur","PHY","GQ_PHYSIQUE;","","SOR","STO","-")');
    end;
   // GI/Ga
   executesql ('Update parpiece set gpp_activitepupr=""');
  end;
End;


Procedure MajVer676 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GPAO
    ExecuteSQL('UPDATE WPDRTET SET WPE_NATUREPDR= "BTH", WPE_LIGNEORDRE=0 WHERE WPE_TYPEPDR="ACT" AND WPE_NATUREPDR IS NULL');
    ExecuteSQL('UPDATE WPDRTYPE SET WRT_NATUREPDR= "BTH" WHERE WRT_NATUREPDR IS NULL');
    ExecuteSQL('UPDATE WPDRTYPE SET WRT_PARTYPEPARTN= "-", WRT_PARTYPECOUTN="-", WRT_PARTYPEENTITEN="-", WRT_PARENTITEN="-", WRT_PARSECTIONN="-"'
                  +', WRT_PARRUBRIQUEN="-", WRT_PARORIGINEN="-", WRT_PAROPECIRCN="-", WRT_PARTYPEPARTG= "-"'
                  +', WRT_PARTYPECOUTG="-", WRT_PARTYPEENTITEG="-", WRT_PARENTITEG="-", WRT_PARSECTIONG="-"'
                  +', WRT_PARRUBRIQUEG="-", WRT_PARORIGINEG="-",  WRT_PAROPECIRCG="-", WRT_DEFAUTN= "X"'
                  +', WRT_DEFAUTG = "X"  WHERE WRT_PARTYPEPARTN IS NULL');
    // supprimé en 691 ExecuteSQL('UPDATE WORDREAUTO SET WOA_BOOLLIBRE9= "-" WHERE WOA_BOOLLIBRE9 IS NULL');
    ExecuteSQL('UPDATE WORDREBES SET WOB_TYPECOMPOSANT= "" WHERE WOB_TYPECOMPOSANT IS NULL');
    ExecuteSQL('UPDATE WORDRERES SET WOR_SECTIONPDR="", WOR_RUBRIQUEPDR="" WHERE WOR_SECTIONPDR IS NULL');
    ExecuteSQL('UPDATE WGAMMERES SET WGR_SECTIONPDR="", WGR_RUBRIQUEPDR="" WHERE WGR_SECTIONPDR IS NULL');
    //
    if not ExisteSQL('SELECT QDE_CIRCUIT FROM QDETCIRC WHERE QDE_CIRCUIT="NUL"') then
    begin
     ExecuteSQL('INSERT INTO QDETCIRC (QDE_CTX,QDE_CIRCUIT,QDE_OPECIRC,QDE_JALON,QDE_POLE,QDE_SITE,QDE_GRP,QDE_TYPTRANS,QDE_JALONMAX,QDE_JALONCOURT,QDE_TYPTRANSCOURT,QDE_DATEMODIF)'
           + ' VALUES ("0","NUL","110","","","","","","","","","'+ UsDateTime(iDate1900)  +'")');
    end
    else
    begin
     ExecuteSQL('update QDETCIRC set QDE_JALON="",QDE_JALONMAX="",QDE_JALONCOURT=""  where QDE_CIRCUIT="NUL"');
    end;
    ExecuteSQL('UPDATE TRECRITURE SET TE_COMMISSION = "S"');
    ExecuteSQL('UPDATE FLUXTRESO SET TFT_CLASSEFLUX = "REF" WHERE TFT_FLUX IN ("EQD", "EQR", "REI")');
    InsertChoixCode('TRF', 'COO', 'Commissions pour OPCVM', 'Commissions OPCVM', '');
    InsertChoixCode('TRF', 'FRO', 'Frais pour OPCVM', 'Frais OPCVM', '');
  end;
End;


Procedure MajVer680 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
  end;

  //JURI
  CFE5_RegulLiacodedo;
End;

Procedure MajVer681 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSQL('UPDATE WPARAM SET WPA_VARCHAR08="BTH" WHERE WPA_CODEPARAM="PRIXDEREVIENT" AND WPA_VARCHAR08 IS NULL');
  end;
End;

Procedure MajVer683 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
  end;
  //PCL
  CPMajTLVersCHOIXCOD;
  //GIGA
  //pour tablette hierarchique
  ExecuteSql (' update YDATATYPETREES set YDT_PREDEFINI="DOS" where YDT_PREDEFINI<>"CEG" and YDT_CODEHDTLINK <> "YYGEDNIV1GEDNIV2" and YDT_CODEHDTLINK <> "YYUSERMASTERSLAVE"');
	//pour GIGA
  ExecuteSql ('DELETE FROM CHOIXCOD WHERE CC_TYPE="ATU" AND CC_CODE in ("030","031","032","033","034","035","036")');
  InsertChoixCode('ATU', '030', 'opérations commerciales', '', 'campagnes marketing');
  InsertChoixCode('ATU', '031', 'opération commerciale', '', 'campagne marketing');
  InsertChoixCode('ATU', '032', 'opérations', '', 'campagnes');
  InsertChoixCode('ATU', '033', 'opération', '', 'campagne');
  InsertChoixCode('ATU', '034', 'commerciaux', '', 'représentants');
  InsertChoixCode('ATU', '035', 'commerciale', '', 'du représentant');
  InsertChoixCode('ATU', '036', 'commercial', '', 'représentant');
  InsertChoixCode('ATU', '100', 'concurrents', '', 'confrères');
  InsertChoixCode('ATU', '101', 'concurrent', '', 'confrère');


End;

Procedure MajVer684 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSql ('Update FACTAFF set afa_transfencours="-"');
    UpdateDecoupePiece ('GP_FERMETUREAFF="-"');
    UpdateNatureSuivante;
  end;
End;

Procedure MajVer685 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GI/GA
    InsertChoixCode('ATU', '800', 'date fin garantie', '', 'date d''épuration');
    InsertChoixCode('ATU', '801', 'date garantie', '', 'date d''épuration');
    InsertChoixCode('ATU', '802', 'date de fin de garantie', '', 'date d''épuration');
    ExecuteSql ('UPDATE FACTAFF SET AFA_PROCENCPR=0, AFA_PROCENCFR=0, AFA_PROCENCFO=0, AFA_AFFAIRETRAENC=""');
    //GPAO
    // supprimée en 691 ExecuteSQL('UPDATE WORDREAUTO SET WOA_BOOLLIBREA= "-" WHERE WOA_BOOLLIBREA IS NULL');
    ExecuteSQL('UPDATE WPDRLIG SET WPL_TIERS="", WPL_CIRCUIT="", WPL_DEPOT="", WPL_SITE="" WHERE WPL_TIERS IS NULL');
    ExecuteSQL('UPDATE DEVISE SET D_PARITEEUROFIXING=D_PARITEEURO,D_ARRONDIPRIXACHAT=D_DECIMALE,D_ARRONDIPRIXVENTE=D_DECIMALE WHERE D_ARRONDIPRIXVENTE<1');
    ExecuteSQL('UPDATE WORDREPHASE SET WOP_VALEURFIXE=0, WOP_PRIXUNIT=0 WHERE WOP_VALEURFIXE IS NULL');
  end;
End;

Procedure MajVer686 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSQL('UPDATE WTRAITEMENT SET WTR_TYPETRAITSTR="", WTR_SECTIONPDR="", WTR_RUBRIQUEPDR="" WHERE WTR_TYPETRAITSTR IS NULL');

    if not ExisteSQL('SELECT WRT_TYPEPDR FROM WPDRTYPE WHERE WRT_TYPEPDR="ORR"') then
    begin
      ExecuteSQL('INSERT INTO WPDRTYPE '
      +'(WRT_TYPEPDR, WRT_NATUREPDR, WRT_LIBELLE, WRT_AVECPERTE, WRT_AVECINDIRECT, WRT_AVECFLUX, WRT_AVECMARGE, '
      +' WRT_AVECTARIF, WRT_FRACTFIXE, WRT_AVECQPCB, WRT_TVACHAT, WRT_TVRESSOURCE, WRT_DATECREATION, WRT_DATEMODIF, '
      +' WRT_CREATEUR, WRT_UTILISATEUR, WRT_SOCIETE, WRT_DEFAUTN, WRT_PARTYPEPARTN, WRT_PARTYPECOUTN, WRT_PARTYPEENTITEN, '
      +' WRT_PARENTITEN, WRT_PARSECTIONN, WRT_PARRUBRIQUEN, WRT_PARORIGINEN, WRT_PAROPECIRCN, WRT_DEFAUTG, WRT_PARTYPEPARTG, '
      +' WRT_PARTYPECOUTG, WRT_PARTYPEENTITEG, WRT_PARENTITEG, WRT_PARSECTIONG, WRT_PARRUBRIQUEG, WRT_PARORIGINEG, WRT_PAROPECIRCG, '
      +' WRT_DEFAUTP, WRT_PARTYPEPARTP, WRT_PARTYPECOUTP, WRT_PARTYPEENTITEP, WRT_PARENTITEP, WRT_PARSECTIONP, WRT_PARRUBRIQUEP, WRT_PARORIGINEP, WRT_PAROPECIRCP)'
      +'VALUES("ORR","ORD","Ordre : réel","-","-","-","-","-","D","-","01","T","'+usDateTime(V_PGI.DateEntree)+'","'+usDateTime(V_PGI.DateEntree)+'","CEG","CEG","001", '
      +'"X","-","-","-","-","-","-","-","-","X","-","-","-","-","-","-","-","-","X","-","-","-","-","-","-","-","-")'
      );
    end;

    if not ExisteSQL('SELECT WRT_TYPEPDR FROM WPDRTYPE WHERE WRT_TYPEPDR="ORT"') then
    begin
      ExecuteSQL('INSERT INTO WPDRTYPE '
      +'(WRT_TYPEPDR, WRT_NATUREPDR, WRT_LIBELLE, WRT_AVECPERTE, WRT_AVECINDIRECT, WRT_AVECFLUX, WRT_AVECMARGE, '
      +' WRT_AVECTARIF, WRT_FRACTFIXE, WRT_AVECQPCB, WRT_TVACHAT, WRT_TVRESSOURCE, WRT_DATECREATION, WRT_DATEMODIF, '
      +' WRT_CREATEUR, WRT_UTILISATEUR, WRT_SOCIETE, WRT_DEFAUTN, WRT_PARTYPEPARTN, WRT_PARTYPECOUTN, WRT_PARTYPEENTITEN, '
      +' WRT_PARENTITEN, WRT_PARSECTIONN, WRT_PARRUBRIQUEN, WRT_PARORIGINEN, WRT_PAROPECIRCN, WRT_DEFAUTG, WRT_PARTYPEPARTG, '
      +' WRT_PARTYPECOUTG, WRT_PARTYPEENTITEG, WRT_PARENTITEG, WRT_PARSECTIONG, WRT_PARRUBRIQUEG, WRT_PARORIGINEG, WRT_PAROPECIRCG, '
      +' WRT_DEFAUTP, WRT_PARTYPEPARTP, WRT_PARTYPECOUTP, WRT_PARTYPEENTITEP, WRT_PARENTITEP, WRT_PARSECTIONP, WRT_PARRUBRIQUEP, WRT_PARORIGINEP, WRT_PAROPECIRCP)'
      +'VALUES("ORT","ORD","Ordre : théorique","-","-","-","-","-","D","-","01","T","'+usDateTime(V_PGI.DateEntree)+'","'+usDateTime(V_PGI.DateEntree)+'","CEG","CEG","001", '
      +'"X","-","-","-","-","-","-","-","-","X","-","-","-","-","-","-","-","-","X","-","-","-","-","-","-","-","-")'
      );
    end;
    //GC
    ExecuteSQL('UPDATE ARTICLE SET GA_CLASSE1="",GA_CLASSE2="",GA_CLASSE3="",GA_CLASSE4="",GA_CLASSE5=""');


    If IsMonoOuCommune and (V_PGI.ModePCL='1') then
    begin // mise à jour table tiers si base 00 et PCL
          ExecuteSql ('update tiers set t_enseigne=(select ##TOP 1## ann_enseigne from annuaire where ann_tiers=t_tiers) where t_tiers in (select ann_tiers from annuaire where ann_tiers=t_tiers and ann_enseigne <>"")');
    end;
  end;
End;

Procedure MajVer687 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GIGA
    AglNettoieListes('AFMULACTIVITE', 'ACT_ACTORIGINE',nil);
    AglNettoieListes('AFMULACTIVITEBM', 'ACT_ACTORIGINE',nil);
    AglNettoieListes('AFMULAPPRECON', 'AFA_AFFAIRETRAENC',nil);
    // GPAO
    ExecuteSQL('UPDATE EMPLACEMENT SET GEM_DATECREATION="' + UsdateTime(iDate1900) + '", GEM_DATEMODIF="' + UsdateTime(iDate1900) + '", GEM_CREATEUR="", GEM_UTILISATEUR="", GEM_STATUTDISPO="", GEM_STATUTFLUX=""');
    GPCopyWPRINTTV;
  end;

End;


Procedure MajVer688 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GPAO
    ExecuteSQL('UPDATE LIGNECOMPL SET GLC_INDICEARTICLE=""');
    ExecuteSQL('UPDATE WNOMEDEC SET WND_INDICECOMPOSAN=""');
    ExecuteSQL('UPDATE WNOMELIG SET WNL_INDICECOMPOSAN=""');
    ExecuteSQL('UPDATE WORDREBES SET WOB_INDICECOMPOSAN=""');
    ExecuteSQL('UPDATE WORDRELIG SET WOL_INDICEARTICLE=""');
    ExecuteSQL('UPDATE WGAMMELIG SET WGL_TYPEEMPLOI=""');
    ExecuteSQL('UPDATE WORDREBES SET WOB_TYPEEMPLOI="", WOB_INDICECOMPOSAN=""');
    ExecuteSQL('UPDATE WORDREGAMME SET WOG_TYPEEMPLOI=""');
    //GC
    ExecuteSQL('UPDATE PARAMCLASSEABC SET GPM_NATUREPIECEG=ISNULL(GPM_NATUREPIECEG,"") '
               +',GPM_CLASSE1=ISNULL(GPM_CLASSE1,""),GPM_CLASSE2=ISNULL(GPM_CLASSE2,"")'
               +',GPM_CLASSE3=ISNULL(GPM_CLASSE3,""),GPM_CLASSE4=ISNULL(GPM_CLASSE4,"")'
               +', GPM_CLASSE5=ISNULL(GPM_CLASSE5,""),GPM_SEUIL1=ISNULL(GPM_SEUIL1,0)'
               +',GPM_SEUIL2=ISNULL(GPM_SEUIL2,0),GPM_SEUIL3=ISNULL(GPM_SEUIL3,0)'
               +',GPM_SEUIL4=ISNULL(GPM_SEUIL4,0),GPM_SEUIL5=ISNULL(GPM_SEUIL5,0)');
  //giGA
    AglNettoieListes('AFMULFACPROAFF', 'GP_FERMETUREAFF',nil);
    AglNettoieListes('AFMULFACTAFFAPP', 'AFA_LIQUIDATIVE',nil);
    AglNettoieListes('AFMULPIECE1', 'GP_FERMETUREAFF',nil);
    AglNettoieListes('AFMULVISFACPROAFF', 'GP_FERMETUREAFF',nil);
    AglNettoieListes('GCEDITPIECE', 'GP_FERMETUREAFF',nil);

 end;
 // Compta
 ExecuteSQL('update generaux set g_cutoff="-",g_cutoffperiode="",g_cutoffechue="-",g_visarevision="-", g_totdebn2=0,g_totcren2=0,g_cyclerevision=""');
 ExecuteSQL('update generauxref set ger_cutoff="-",ger_cutoffperiode="",ger_cutoffechue="-"');
End;

Procedure MajVer689 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GPAO
    AGLNettoieListes('WORDRELIG2', 'WOL_ORDREPERE', nil);
    AGLNettoieListes('GCDUPLICPIECE', 'GP_CDETYPE', nil);
    AGLNettoieListes('GCGROUPEPIECE', 'GP_CDETYPE', nil);
    AGLNettoieListes('GCMULMODIFDOCVEN', 'GP_CDETYPE', nil);
    AGLNettoieListes('GCMULPIECE', 'GP_CDETYPE', nil);
    AGLNettoieListes('GCMULPIECEVISA', 'GP_CDETYPE', nil);
    AGLNettoieListes('GCMULTRANSFDOCVEN', 'GP_CDETYPE', nil);
    AGLNettoieListes('GCMULLIGNE', 'GL_TYPECADENCE', nil);

    UpdateDecoupePiece('GP_CDETYPE=IIF(GP_NATUREPIECEG="CC", "FER", ""), GP_ARTICLE="", GP_CODEARTICLE= ""');
    UpdateDecoupeLigne('GL_TYPECADENCE=IIF(GL_NATUREPIECEG="CC", "FER", ""), GL_REFTRANSPORT="", GL_DATEENLEVEMENT = "' + UsDateTime(IDate1900) + '"');

    ExecuteSql('UPDATE WNOMELIG SET WNL_INDICECOMPOSAN=""');
    ExecuteSQL('UPDATE WORDRELIG SET WOL_ORDREPERE=0, WOL_OPECIRCWOB="", WOL_LIENNOMEWOB=0');
  end;
End;

Procedure MajVer690 ;
var
  iWPF: integer;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSql('UPDATE COMPTADIFFEREE SET GCD_TIERSPAYEUR=""');
    //GRC
    ExecuteSql('UPDATE PARCHAINAGES SET RPG_NONTERMINE="-"');
    ExecuteSql('UPDATE PARPIECE SET GPP_CHAINAGE=""');
    ExecuteSql('UPDATE DOMAINEPIECE SET GDP_CHAINAGE=""');
    // GPAO
    // supprimé 691 ExecuteSQL('UPDATE TIERS SET T_ESTTRANSPORTEUR="-", T_SECTEURGEO ="", T_TIMBRE=0,  T_SURTAXE=0  WHERE T_ESTTRANSPORTEUR IS NULL');
    ExecuteSQL('UPDATE ADRESSES SET ADR_SECTEURGEO="" WHERE ADR_SECTEURGEO IS NULL');
    ExecuteSQL('UPDATE PIECEADRESSE SET GPA_SECTEURGEO="" WHERE GPA_SECTEURGEO IS NULL');
    ExecuteSQL('UPDATE YTARIFSPARAMETRES SET YFO_OKSECTEURGEO="",YFO_OKADRESSE="",YFO_OKMODEEXP="" WHERE YFO_OKSECTEURGEO IS NULL');
    ExecuteSQL('UPDATE YTARIFS SET YTS_SECTEURGEO="",YTS_CASCSECTEURGEO="-", YTS_CASCTOUSSECT="-",YTS_CODEPOSTAL="",YTS_PAYS="",YTS_REGION="",YTS_MODEEXP=""  WHERE YTS_SECTEURGEO IS NULL');
    If (V_PGI.driver <> dbORACLE7) then
      ExecuteSQL('UPDATE YTARIFS SET YTS_POIDSRECHERCHE= '
      +' iif((ISNULL(YTS_TRAITEMENT," ")<>" "),1,0) '
      +'+iif((ISNULL(YTS_PHASE," ")<>" "),2,0) '
      +'+iif((ISNULL(YTS_CIRCUIT," ")<>" "),4,0) '
      +'+iif((ISNULL(YTS_FAMILLENIV3," ")<>" "),8,0) '
      +'+iif((ISNULL(YTS_FAMILLENIV2," ")<>" "),16,0) '
      +'+iif((ISNULL(YTS_FAMILLENIV1," ")<>" "),32,0) '
      +'+iif((ISNULL(YTS_TARIFARTICLE," ")<>" "),64,0) '
      +'+iif(((ISNULL(YTS_CODEARTICLE," ")<>" ")and(Len(YTS_CODEARTICLE)=0)),128,0) '
      +'+iif(((ISNULL(YTS_CODEARTICLE," ")<>" ")and(Len(YTS_CODEARTICLE)<>0)),256,0) '
      +'+iif((ISNULL(YTS_SITE," ")<>" "),512,0) '
      +'+iif((ISNULL(YTS_TARIFAFFAIRE," ")<>" "),1024,0) '
      +'+iif((ISNULL(YTS_AFFAIRE," ")<>" "),2048,0) '
      +'+iif((ISNULL(YTS_PAYS," ")<>" "),4096,0) '
      +'+iif((ISNULL(YTS_REGION," ")<>" "),8192,0) '
      +'+iif((Len(YTS_CODEPOSTAL)=2),16384,0) '
      +'+iif((ISNULL(YTS_CODEPOSTAL," ")<>" "),32768,0) '
      +'+iif((ISNULL(YTS_SECTEURGEO," ")<>" "),65536,0) '
      +'+iif((ISNULL(YTS_MODEEXP," ")<>" "),131072,0) '
      +'+iif((ISNULL(YTS_TARIFTIERS," ")<>" "),262144,0) '
      +'+iif((ISNULL(YTS_TIERS," ")<>" "),524288,0) '
      +'+iif((ISNULL(YTS_TARIFSPECIAL," ")<>" "),1048576,0) '
      +'+iif(((ISNULL(YTS_DEVISE," ")<>" ")and(YTS_DEVISE<>"EUR")),2097152,0) '
      +'+iif((ISNULL(YTS_TARIFDEPOT," ")<>" "),4194304,0) '
      +'+iif((ISNULL(YTS_DEPOT," ")<>" "),8388608,0)');

    AGLNettoieListes('YTARIFSMULART101','YTS_CODEPOSTAL;YTS_PAYS;YTS_REGION;YTS_SECTEURGEO;YTS_MODEEXP',nil);
    AGLNettoieListes('YTARIFSMULART201','YTS_CODEPOSTAL;YTS_PAYS;YTS_REGION;YTS_SECTEURGEO;YTS_MODEEXP',nil);
    AGLNettoieListes('YTARIFSMULART211','YTS_CODEPOSTAL;YTS_PAYS;YTS_REGION;YTS_SECTEURGEO;YTS_MODEEXP',nil);
    AGLNettoieListes('YTARIFSMULART301','YTS_CODEPOSTAL;YTS_PAYS;YTS_REGION;YTS_SECTEURGEO;YTS_MODEEXP',nil);
    AGLNettoieListes('YTARIFSMULART401','YTS_CODEPOSTAL;YTS_PAYS;YTS_REGION;YTS_SECTEURGEO;YTS_MODEEXP',nil);
    AGLNettoieListes('YTARIFSMULART101','YTS_CODEPOSTAL;YTS_PAYS;YTS_REGION;YTS_SECTEURGEO;YTS_MODEEXP',nil);
    AGLNettoieListes('YTARIFSMULTIE101','YTS_CODEPOSTAL;YTS_PAYS;YTS_REGION;YTS_SECTEURGEO;YTS_MODEEXP',nil);
    AGLNettoieListes('YTARIFSMULTIE201','YTS_CODEPOSTAL;YTS_PAYS;YTS_REGION;YTS_SECTEURGEO;YTS_MODEEXP',nil);
    AGLNettoieListes('YTARIFSMULTIE211','YTS_CODEPOSTAL;YTS_PAYS;YTS_REGION;YTS_SECTEURGEO;YTS_MODEEXP',nil);
    AGLNettoieListes('YTARIFSMULTIE301','YTS_CODEPOSTAL;YTS_PAYS;YTS_REGION;YTS_SECTEURGEO;YTS_MODEEXP',nil);
    AGLNettoieListes('YTARIFSMULTIE401','YTS_CODEPOSTAL;YTS_PAYS;YTS_REGION;YTS_SECTEURGEO;YTS_MODEEXP',nil);

    { Mise à jour de WPARAMFONCTION }
    if Assigned(TobWPF) then
    begin
      for iWPF := 0 to TobWPF.Detail.Count-1 do
      begin
        TobWPF.Detail[iWPF].VirtuelleToReelle('WPARAMFONCTION');
        TobWPF.Detail[iWPF].InsertDB(nil);
      end;
      FreeAndNil(TobWPF);
    end;
  end;
End;

Procedure MajVer691 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    // GPAO
    // MCD 20102005 AGLNettoieListes('RTMULTIERSRECH','YTC_ESTTRANSPORTR',nil);
    // {YTC_ESTTRANSPORTR="-",} supprimé en 699
    ExecuteSQL('UPDATE TIERSCOMPL SET  YTC_SECTEURGEO ="",YTC_TIMBRE=0,YTC_SURTAXE=0  WHERE YTC_SECTEURGEO IS NULL');
    ExecuteSQL('UPDATE WPDRLIG SET WPL_CONSOLIDE="-" WHERE WPL_CONSOLIDE IS NULL');
    AGLNettoieListes('YTARIFSMULART101','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULART201','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULART211','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULART301','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULART401','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULART501','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULTIE101','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULTIE201','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULTIE211','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULTIE301','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULTIE401','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULTIE501','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AglNettoieListes('WORDRELIG','WOL_ORDREPERE',nil);
    AglNettoieListes('WORDRELIG2','WOL_ORDREPERE',nil);
    ExecuteSQL('UPDATE WPARAMFONCTION SET WPF_PREFIXE="", WPF_NATUREPIECEG=""');
    // GC
    ExecuteSQL('UPDATE MODEPAIE SET MP_AVECINFOCOMPL = "-" WHERE MP_AVECINFOCOMPL IS NULL');
    ExecuteSQL('UPDATE MODEPAIE SET MP_AVECNUMAUTOR = "-" WHERE MP_AVECNUMAUTOR IS NULL ');
    ExecuteSQL('UPDATE MODEPAIE SET MP_COPIECBDANSCTRL = "-" WHERE MP_COPIECBDANSCTRL IS NULL ');
    ExecuteSQL('UPDATE MODEPAIE SET MP_AFFICHNUMCBUS = "-" WHERE MP_AFFICHNUMCBUS IS NULL ');
    ExecuteSQL('UPDATE MODEPAIE SET MP_CPTECAISSE = "", MP_JALREMBQ = "", MP_CPTEREMBQ = "" ');
    ExecuteSQL('UPDATE MODEPAIECOMPL SET MPC_JALREMBQ = "", MPC_CPTEREMBQ = ""  ');
  end;
End;

Procedure MajVer692 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    // GPAO
    ExecuteSQL('UPDATE LIGNECOMPL SET GLC_ARTICLECFX="" WHERE GLC_ARTICLECFX IS NULL');
    ExecuteSQL('UPDATE YTARIFSPARAMETRES SET YFO_OKSECTEURGEO="---" WHERE YFO_OKSECTEURGEO IS NULL OR YFO_OKSECTEURGEO=""');
    ExecuteSQL('UPDATE YTARIFSPARAMETRES SET YFO_OKADRESSE="---" WHERE YFO_OKADRESSE IS NULL OR YFO_OKADRESSE=""');
    ExecuteSQL('UPDATE YTARIFSPARAMETRES SET YFO_OKMODEEXP="---" WHERE YFO_OKMODEEXP IS NULL OR YFO_OKMODEEXP=""');
    //GC
    ExecuteSQL('UPDATE COMPTADIFFEREE SET GCD_TIERSPAYEUR = "" WHERE GCD_TIERSPAYEUR IS NULL');
    ExecuteSQL('UPDATE PARPIECE SET GPP_SOLDETRANSFO="-" WHERE GPP_SOLDETRANSFO IS NULL');
    ExecuteSQL('UPDATE PARPIECE SET GPP_REGROUPE="X" WHERE GPP_REGROUPE IS NULL');
    //GI/GA
    AglNettoieListes('AFMULAPPRECON', 'AFA_TRANSFENCOURS',nil);
    AglNettoieListes('AFMULAPPRECONANU', 'AFA_AFFAIRETRAENC;AFA_TRANSFENCOURS',nil);
  end;
  ExecuteSQL('update generaux set g_cutoffcompte="" ');
  ExecuteSQL('update generauxref set ger_cutoffcompte=""');
End;

Procedure MajVer694 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GPAO  reprise initialisation fait sur 677 pour 6.20
    ExecuteSQL('UPDATE ARTICLE SET GA_LOWLEVELCODE=-1 WHERE GA_LOWLEVELCODE IS NULL');
    // GPAO
    ExecuteSQL('DELETE FROM YTARIFSPARAMETRES WHERE YFO_FONCTIONNALITE="501" AND (YFO_ORIENTATION="TIE" OR YFO_ORIENTATION="ART")');
  end;
End;

Procedure MajVer695 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
     //Gi/GA
    ExecuteSql ('Update ANNUAIRE SET ANN_NATUREAUXI=""');
    ExecuteSql ('update EPIECE set ep_majlibretiers="XXXXXXXXXXXXX"');
    ExecuteSql ('update  JUTYPEPER set JTP_NATUREAUXI=""');
    ExecuteSQL ('UPDATE PIECE SET GP_MAJLIBRETIERS=substring(GP_MAJLIBRETIERS,1,11)||"XX"'
            + ' WHERE (GP_NATUREPIECEG="' + GetParamSoc ('SO_AFNATPROPOSITION') + '"'
            + ' OR    GP_NATUREPIECEG="' + GetParamSoc ('SO_AFNATAFFAIRE') + '")');
    AglNettoieListes('AFCTRL_ANNTIERS', '',nil);
    //fin GIGA
    //GPAO
    AglNettoieListes('WORDRETET', 'WOT_NATUREPIECEG;WOT_SOUCHE;WOT_INDICEG', nil);
    AglNettoieListes('WORDRELIG', 'WOL_ORDREPERE', nil);
    AglNettoieListes('WORDRELIG3', 'WOL_ORDREPERE', nil);

  end;
// PAYE
ExecuteSql('DELETE FROM CHOIXCOD WHERE CC_TYPE="PSC" AND (CC_LIBRE="" OR CC_LIBRE IS NULL)');
AglNettoieListes('PGACOMPTE', 'PSA_ETABLISSEMENT;PSD_SALARIE;PSA_LIBELLE;PSA_PRENOM;PSD_MONTANT;PSD_DATEDEBUT;PSD_DATEPAIEMENT;PSD_DATEFIN;PSD_ORIGINEMVT;PSD_ORDRE;PSD_RUBRIQUE',nil);
AglNettoieListes('PGVIRSALAIRES', 'PVI_SALARIE;PVI_LIBELLE;PVI_PRENOM;PVI_ETABLISSEMENT;PVI_DATEDEBUT;PVI_DATEFIN;PVI_MONTANT;PVI_PAYELE;PVI_BANQUE;PVI_BANQUEEMIS;PVI_TOPREGLE;PVI_AUXILIAIRE;PVI_RIBSALAIRE;PVI_RIBSALAIREEMIS',nil);
AglNettoieListes('PGMULMVTCP', 'PCN_ORDRE;PCN_ETABLISSEMENT;PCN_SALARIE;PCN_DATEVALIDITE;PCN_TYPECONGE;PCN_SENSABS;PCN_JOURS;PCN_HEURES;PCN_LIBELLE;PCN_CODETAPE;PCN_TYPEMVT;PCN_APAYES;PCN_BASE;PCN_NBREMOIS',nil);
AglNettoieListes('PGMULCONTRAT', 'PSA_ETABLISSEMENT;PCI_ETABLISSEMENT;PCI_SALARIE;PSA_LIBELLE;PSA_PRENOM;PCI_TYPECONTRAT;PCI_DEBUTCONTRAT;PCI_FINCONTRAT',nil);
ExecuteSql('UPDATE PAIEENCOURS SET PPU_ECHEANCE = PPU_PAYELE');
//PDumet 12102005
v_pgi.enableDEShare := True;
V_PGI.StandardSurDP := True;
ExecuteSql('update histoanalpaie set pha_ordreetat = (SELECT prm_ordreetat from remuneration WHERE ##prm_predefini## pha_rubrique=prm_rubrique) WHERE pha_naturerub="AAA" ');
//PDumet 12102005
v_pgi.enableDEShare := false;
V_PGI.StandardSurDP := False;
ExecuteSql('UPDATE histoanalpaie set pha_ordreetat = 3 where pha_naturerub<>"AAA"');
If ExisteSQL('SELECT * FROM ABSENCESALARIE WHERE PCN_TYPEMVT="ABS"') then
   ExecuteSql('UPDATE PARAMSOC SET SOC_DATA = "X" WHERE SOC_NOM="SO_PGABSENCE"');
ExecuteSql('UPDATE RHCOMPETENCES SET PCO_LIBELLECOMPL= ""');
// ExecuteSql('UPDATE histoanalpaie SET PHA_COTREGUL= "..."');
// ExecuteSql('UPDATE histobulletin SET PHB_COTREGUL= "..."');
// COMPTA
// AYEL 25102005 ExecuteSql('UPDATE STDCPTA SET STC_DATEMODIF="' + UsdateTime(iDate1900) + '"');
End;


Procedure MajVer696 ;
var ss,SQL : string;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
       //GC
       ExecuteSql('UPDATE STKMOUVEMENT SET GSM_CONTREMARQUE="-", GSM_FOURNISSEUR="", GSM_REFERENCE=""');
       ExecuteSql('UPDATE DISPODETAIL SET GQD_CONTREMARQUE="-", GQD_FOURNISSEUR="", GQD_REFERENCE=""');
       //GPAO
       ExecuteSQL('update ARTICLECOMPL set GA2_FAMILLEMAT="" where GA2_FAMILLEMAT is null');
       ExecuteSQL('update PROFILART  set GPF_FAMILLEMAT="" where GPF_FAMILLEMAT is null');
       ExecuteSQL('update WORDRELIG  set WOL_SCMORDO="" where WOL_SCMORDO is null');
       ExecuteSQL('update WPARAMFONCTION  set WPF_FAMILLEMAT="", WPF_TYPCOMP="" where WPF_FAMILLEMAT is null');
       ExecuteSQL('update UCOMPFOURNI  set UCF_FAMILLEMAT="", UCF_TYPCOMP="" where UCF_FAMILLEMAT is null');
       //DP
       ExecuteSQL('UPDATE DPFISCAL SET DFI_DOMHORSFRANCE="", DFI_BAFORFAIT="", DFI_DISTRIBDIVID="-"');
       ExecuteSQL('UPDATE DPORGA SET DOR_UTILRESPCOMPTA="", DOR_UTILRESPSOCIAL="", DOR_UTILRESPJURID="", DOR_UTILCHEFGROUPE="", DOR_CABINETASSOC=""');
        //GIGA
       ExecuteSql ('UPDATE ANNUAIRE SET ANN_NATUREAUXI="CLI" where ANN_TIERS <>""');
       ExecuteSql ('UPDATE AFFAIRE set AFF_ISECOLE="-",AFF_DOMAINE=""');
       SS := UsDateTime(IDate1900);
       Sql := 'update afftiers set AFT_TYPEORIG="AFF",AFT_NUMORIG=0,AFT_TYPECONTACT="",AFT_AUXILIAIRE="",'
         +'AFT_NUMEROCONTACT=0,AFT_INTERVENTION="",AFT_DATECREATION="'+SS+'",AFT_DATEMODIF="'+SS
         +'",AFT_CREATEUR ="",AFT_UTILISATEUR=""';
       ExecuteSQL(SQL);
       ExecuteSQL('update afftiers  set AFT_TYPEINTERV="CLI" where  AFT_TIERS <>""');
       ExecuteSQL('update afftiers  set AFT_TYPEINTERV="RES" where  AFT_TIERS =""' );
       ExecuteSql ('UPDATE articlecompl set GA2_typeplanif="NON",GA2_QTEAPLANIF=0');
       ExecuteSql ('UPDATE DOMAINEPIECE set GDP_MODPLANIFIABLE=""');
       ExecuteSql ('UPDATE lignecompl set glc_typeplanif="NON",GLC_QTEAPLANIF=0,GLC_AFFAIRELIEE=""');
       ExecuteSql ('UPDATE PARPIECE set GPP_MODPLANIFIABLE=""');
       UpdateDecoupePiece('GP_PLANIFIABLE="-"');
       ExecuteSql ('update ressource set ars_equiperess="" ');
       ExecuteSql ('update tache set ata_qtefactpla=0,ata_qtefactref=0,ata_qteplanifpla=0,ata_qteplanifref=0,ata_typeplanif=""');
       ExecuteSql ('update afplanning set apl_bloque="-",apl_estfacture="-",apl_numpiece=""');
        //fin GIGA
  end;
End;


Procedure MajVer697 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
      //GIGA
    InsertChoixCode('ZLI', 'OT1', 'Table libre 1', 'Table libre 1', '');
    InsertChoixCode('ZLI', 'OT2', 'Table libre 2', 'Table libre 2', '');
    InsertChoixCode('ZLI', 'OT3', 'Table libre 3', 'Table libre 3', '');
    AglNettoieListes('AFMULFACTAFFAPP','AFF_NATUREPIECEG',nil);
    AglNettoieListes('AFMULFACTIERSAFF','AFF_NATUREPIECEG',nil);
    AglNettoieListes('AFLIGPLANNING','APL_ARTICLE;APL_CODEARTICLE;APL_TYPEARTICLE',nil);


	//Fin GIGA
  //GRC
    ExecuteSQL('UPDATE PARPIECE SET GPP_MODELEWORD="" ');
    ExecuteSQL('UPDATE PARPIECECOMPL SET GPC_MODELEWORD=""  ');
  // GPAO
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="", GSN_SFLUXPICKING="" where GSN_QUALIFMVT="000"'
            + ' or GSN_QUALIFMVT="AAC" or gsn_qualifmvt="ACP" or gsn_qualifmvt="ADM" or gsn_qualifmvt="APF" or gsn_qualifmvt="APR" or gsn_qualifmvt="ATD"'
            + ' or gsn_qualifmvt="EAC" or gsn_qualifmvt="ECP" or gsn_qualifmvt="EDM" or gsn_qualifmvt="EEX" or gsn_qualifmvt="EIN" or gsn_qualifmvt="EIT"'
            + ' or gsn_qualifmvt="EPR" or gsn_qualifmvt="ETR" or gsn_qualifmvt="EVE" or gsn_qualifmvt="MPR"');

    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;AFF;", GSN_SFLUXPICKING="STD;PCO;PVE" where GSN_QUALIFMVT="CAF"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="BLQ", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;", GSN_SFLUXPICKING="STD;PCO;PVE" where GSN_QUALIFMVT="CBS"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="LBR", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="BLQ;", GSN_SFLUXPICKING="STD;PCO;PVE" where GSN_QUALIFMVT="CDS"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="<<Tous>>", GSN_SFLUXPICKING="<<Tous>>" where GSN_QUALIFMVT="CEM"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;", GSN_SFLUXPICKING="STD;" where GSN_QUALIFMVT="CIA"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="<<Tous>>", GSN_SFLUXPICKING="<<Tous>>" where GSN_QUALIFMVT="CLO"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;AFF;", GSN_SFLUXPICKING="STD;PCO;" where GSN_QUALIFMVT="RDM"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;AFF;", GSN_SFLUXPICKING="STD;PVE;" where GSN_QUALIFMVT="RPL"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;AFF;", GSN_SFLUXPICKING="STD;PCO;" where GSN_QUALIFMVT="RPR"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;AFF;", GSN_SFLUXPICKING="STD;PVE;" where GSN_QUALIFMVT="RTD"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;AFF;", GSN_SFLUXPICKING="STD;PVE;" where GSN_QUALIFMVT="RVE"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;", GSN_SFLUXPICKING="STD;CTR;" where GSN_QUALIFMVT="SAC"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;", GSN_SFLUXPICKING="STD;CTR;" where GSN_QUALIFMVT="SDM"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;BLQ;", GSN_SFLUXPICKING="<<Tous>>" where GSN_QUALIFMVT="SEX"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="<<Tous>>", GSN_SFLUXPICKING="<<Tous>>" where GSN_QUALIFMVT="SIN"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="<<Tous>>", GSN_SFLUXPICKING="<<Tous>>" where GSN_QUALIFMVT="SIT"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;AFF;", GSN_SFLUXPICKING="STD;PCO;" where GSN_QUALIFMVT="SPR"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;", GSN_SFLUXPICKING="STD;PVE" where GSN_QUALIFMVT="STR"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;AFF;", GSN_SFLUXPICKING="STD;PVE;" where GSN_QUALIFMVT="SVE"');

    ExecuteSQL('UPDATE QTYPTRANS SET QTY_PORTDEPART="", QTY_PORTARRIVEE="" WHERE QTY_PORTDEPART IS NULL');

    InsertChoixCode('YTR', 'X', 'pas de recherche', '', '');
    InsertChoixCode('YTU', 'X', 'ne somme pas', '', '');

    //executeSql('update Dispodetail set gqd_dateperemption="' + UsDateTime(iDate2099) + '" where gqd_dateperemption="' + UsdateTime(iDate1900) + '"');
    //executeSql('update stkmouvement set gsm_dateperemption="' + UsDateTime(iDate2099) + '" where gsm_dateperemption="' + UsdateTime(iDate1900) + '"');
    // mofif KB 31082005
    ExecuteSQL('UPDATE DISPODETAIL SET GQD_DATEPEREMPTION="' + UsDateTime(iDate2099) + '" WHERE GQD_DATEPEREMPTION IS NULL OR GQD_DATEPEREMPTION="' + UsdateTime(iDate1900) + '"');
    ExecuteSQL('UPDATE STKMOUVEMENT SET GSM_DATEPEREMPTION="' + UsDateTime(iDate2099) + '" WHERE GSM_DATEPEREMPTION IS NULL OR GSM_DATEPEREMPTION="' + UsdateTime(iDate1900) + '"');
  end;
End;


Procedure MajVer698 ;
var SS : string;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SS := UsDateTime(IDate1900);
  if not IsDossierPCL then
  begin
     ExecuteSQL('UPDATE DPORGA SET DOR_NONTRAITE="-", DOR_MOTIFNONTRAITE="", DOR_WETABENTITE="", DOR_WCATREVENUS="", DOR_WNATUREATTEST="", DOR_WREGIMEIMPO="", DOR_WREGIMETVA="", DOR_WTYPEIMPO=""');
	//bureau, suite destruciton ann_auxiliaire
    AglNettoieListes('AFDOUBLONS_ANNTIE', '',nil);
    AglNettoieListes('AFDOUBLONS_ANNU', '',nil);
    AglNettoieListes('DPANNULIEN2', '',nil);
    AglNettoieListes('DPANNULIEN3', '',nil);
    AglNettoieListes('DPANNUPERS', '',nil);
    AglNettoieListes('DPANNUPERS2', '',nil);
    AglNettoieListes('DPCREAT_TIERS', '',nil);
    AglNettoieListes('DPMULANU', '',nil);
    AglNettoieListes('DPMULANU2', '',nil);
    AglNettoieListes('DPMULANU3', '',nil);
    AglNettoieListes('FPDECLARATIONS', '',nil);
    AglNettoieListes('FPMULDECLA', '',nil);
    AglNettoieListes('FPSELDECLA', '',nil);
    AglNettoieListes('FPSELDECLAR', '',nil);
    AglNettoieListes('JUACTIONNAIRES', '',nil);
    AglNettoieListes('JUPERSONNE_MUL', '',nil);
    AglNettoieListes('JUPERSONNE_SEL', '',nil);
    AglNettoieListes('L2072ASSOCIE04', '',nil);
    AglNettoieListes('L2072GERANT', '',nil);
    AglNettoieListes('L2072LOCBES', '',nil);
    AglNettoieListes('PGANNUAIRE', '',nil);
    AglNettoieListes('YYDOSSIER_SEL', '',nil);
    AglNettoieListes('YYMULANNDOSS', '',nil);
    AglNettoieListes('YYMELSELDOSS', '',nil);
         //fin bureau
	//GIGA
    ExecuteSql ('update afmodeletache set afm_commentaire=""');
    AglNettoieListes('AFMULCOMPARTICLE', 'AAC_COMPETENCE;AAC_CODEARTICLE;AAC_ARTICLE',nil);
    if V_PGI.Driver in [dbORACLE7, dbORACLE8 ,dbORACLE9] then
        begin
        ExecuteSql('UPDATE AFPLANNING set APL_HEUREDEB_PLA = APL_DATEDEBPLA + 8/24');
        ExecuteSql('UPDATE AFPLANNING set APL_HEUREFIN_PLA = APL_DATEFINPLA + 18/24');
        ExecuteSql('UPDATE AFPLANNING set APL_HEUREDEB_REAL = APL_DATEDEBREAL + 8/24'	);
        ExecuteSql('UPDATE AFPLANNING set APL_HEUREFIN_REAL = APL_DATEFINREAL + 18/24');
        end
    else if (V_PGI.Driver = dbDB2) then      //rien à faire
    else begin  // cas SQl ou access ou ...
      ExecuteSql('UPDATE AFPLANNING set APL_HEUREDEB_PLA = APL_DATEDEBPLA + "8:00:00"' );
      ExecuteSql('UPDATE AFPLANNING set APL_HEUREFIN_PLA = APL_DATEFINPLA + "18:00:00"' );
      ExecuteSql('UPDATE AFPLANNING set APL_HEUREDEB_REAL = APL_DATEDEBREAL + "8:00:00"');
      ExecuteSql('UPDATE AFPLANNING set APL_HEUREFIN_REAL = APL_DATEFINREAL + "18:00:00"');
    end;
	//fin GIGA
  //GC
    ExecuteSql ('UPDATE LIGNECOMPL SET GLC_REPRESENTANT2="", GLC_REPRESENTANT3="", GLC_COMMISSIONR2=0, GLC_COMMISSIONR3=0, GLC_TYPECOM2="-", GLC_TYPECOM3="-", GLC_VALIDECOM2="-", GLC_VALIDECOM3="-", GLC_MONTANTTPF=0, GLC_COMMISSIONNABL="-"');
    ExecuteSql ('UPDATE ARTICLECOMPL SET GA2_MONTANTTPF=0 ');
  // GPAO
    ExecuteSQL('update STKNATURE SET GSN_CALLGSS="-" WHERE GSN_CALLGSS IS NULL');
    UpdateDecoupeLigne('GL_COEFCONVQTE=1', ' AND (ISNULL(GL_ARTICLE, " ") <> " ")');
  end;

  ExecuteSql('update remuneration set prm_ordreetat=9 where prm_ordreetat=6');
  ExecuteSql('update remuneration set prm_ordreetat=7 where prm_ordreetat=5');
  ExecuteSql('update remuneration set prm_ordreetat=6 where prm_ordreetat=4');

  //PDumet 12102005
  v_pgi.enableDEShare := True;
  V_PGI.StandardSurDP := True;
  ExecuteSql('UPDATE histobulletin set phb_ordreetat=(Select pct_ordreetat from cotisation '
     +' where ##PCT_predefini## pct_rubrique=phb_rubrique) where phb_naturerub<>"AAA" and phb_rubrique in (select pct_rubrique from cotisation cot where ##cot.pct_predefini## cot.pct_themecot="RDC")');
  ExecuteSql('UPDATE histoanalpaie set pha_ordreetat=(Select pct_ordreetat from cotisation '
     +' where ##PCT_predefini## pct_rubrique=pha_rubrique) where pha_naturerub<>"AAA" and pha_rubrique in (select pct_rubrique from cotisation cot where ##cot.pct_predefini## cot.pct_themecot="RDC")');
  ExecuteSql('UPDATE histobulletin set phb_ordreetat=(Select prm_ordreetat from remuneration '
     +' where ##Prm_predefini## prm_rubrique=phb_rubrique) where phb_naturerub="AAA" and phb_ordreetat>3');
  ExecuteSql('UPDATE histoanalpaie set pha_ordreetat=(Select prm_ordreetat from remuneration '
    +' where ##Prm_predefini## prm_rubrique=pha_rubrique) where pha_naturerub="AAA" and pha_ordreetat>3');

  v_pgi.enableDEShare := False;

  V_PGI.StandardSurDP := FALSE;


  If ExisteSQL('SELECT PFO_SALARIE FROM FORMATIONS WHERE PFO_TYPEPLANPREV IS NULL OR PFO_TYPEPLANPREV=""') then

  begin

    ExecuteSQL('UPDATE FORMATIONS SET PFO_TYPEPLANPREV="PLF", PFO_HTPSTRAV=PFO_NBREHEURE,PFO_HTPSNONTRAV=0,PFO_PAYER="-",PFO_DATEPAIE="'+SS+'",PFO_TYPOFORMATION="001",PFO_ALLOCFORM=0');

    ExecuteSql('UPDATE INSCFORMATION SET PFI_TYPEPLANPREV="PLF",PFI_HTPSTRAV=PFI_DUREESTAGE,PFI_HTPSNONTRAV=0,'

      + 'PFI_TYPOFORMATION="001",PFI_DATEDIF=PFI_DATECREATION,'
      + 'PFI_NATUREFORM=(SELECT PST_NATUREFORM FROM STAGE WHERE PST_CODESTAGE=PFI_CODESTAGE AND PST_MILLESIME=PFI_MILLESIME),PFI_ALLOCFORM=0');

  end;    

  ExecuteSql('update ducsentete set pdu_typbordereau="913" where '+
    'PDU_ABREGEPERIODE not like "%%00" AND PDU_DUCSDOSSIER = "-" AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE '+
        'ON PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE  POG_NATUREORG ="100")');
  ExecuteSql('update ducsentete set pdu_typbordereau="914" where '+
    'PDU_ABREGEPERIODE not like "%%00" AND PDU_DUCSDOSSIER = "X" AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE '+
        'ON PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE  POG_NATUREORG ="100")');
  ExecuteSql('update ducsentete set pdu_typbordereau="915" where '+
    'PDU_ABREGEPERIODE  like "%%00" AND PDU_DUCSDOSSIER = "-" AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE '+
        'ON PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE POG_NATUREORG ="100")');
  ExecuteSql('update ducsentete set pdu_typbordereau="916" where '+
    'PDU_ABREGEPERIODE  like "%%00" AND PDU_DUCSDOSSIER = "X" AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE '+
        'ON PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE POG_NATUREORG ="100")');
  ExecuteSql('update ducsentete set pdu_typbordereau="920" where '+
    'PDU_ABREGEPERIODE  not like "%%00" AND PDU_DUCSDOSSIER = "-" AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE ON '+
        'PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE POG_NATUREORG ="200")');
  ExecuteSql('update ducsentete set pdu_typbordereau="921" where '+
    'PDU_ABREGEPERIODE  not like "%%00" AND PDU_DUCSDOSSIER = "X" AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE '+
       'ON PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE POG_NATUREORG ="200")');
  ExecuteSql('update ducsentete set pdu_typbordereau="922" where '+
    'PDU_ABREGEPERIODE   like "%%00" AND PDU_DUCSDOSSIER = "-" AND PDU_ORGANISME '+
      'IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE ON '+
        'PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE POG_NATUREORG ="200")');
  ExecuteSql('update ducsentete set pdu_typbordereau="923" where '+
    'PDU_ABREGEPERIODE   like "%%00" AND PDU_DUCSDOSSIER = "X" AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE '+
        'ON PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE POG_NATUREORG ="200")');
  ExecuteSql('update ducsentete set pdu_typbordereau="930" where '+
    'PDU_ABREGEPERIODE  like "%%0%" AND PDU_ABREGEPERIODE  not  like "%%00"  AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE '+
        'ON PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE POG_NATUREORG ="300")');
  ExecuteSql('update ducsentete set pdu_typbordereau="931" where '+
    'PDU_ABREGEPERIODE   not like "%%0%" AND PDU_ABREGEPERIODE  not  like "%%00"  AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE '+
        'ON PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE POG_NATUREORG ="300")');
  ExecuteSql('update ducsentete set pdu_typbordereau="932" where '+
    'PDU_ABREGEPERIODE  like "%%00"  AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE '+
        'ON PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE POG_NATUREORG ="300")');
  ExecuteSql('update ducsentete set PDU_NBSALQ944=0, '+
    'PDU_NBSALQ945=0, PDU_NBSALQ960=0, PDU_NBSALQ961=0,PDU_NBSALQ962=0, '+
      'PDU_NBSALQ963=0,PDU_NBSALQ964=0, PDU_NBSALQ965=0,PDU_NBSALQ966=0, '+
        'PDU_NBSALQ967=0,PDU_ECARTZE1="-",PDU_ECARTZE2="-",PDU_ECARTZE3="-",'+
          'PDU_ECARTZE4="-",PDU_ECARTZE5="-",PDU_ECARTZE6="-",PDU_ECARTZE7="-", PDU_ECARTZE8 ="-"');
  ExecuteSql('update emetteursocial set PET_EMAILDUCS=""');
  ExecuteSql('update ducsdetail set PDD_CODECOMMUNE=""');
  ExecuteSql('update organismepaie set POG_TITULAIRECPT=""');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEPERIOD="913" where '+
    'POG_NATUREORG="100" AND POG_DUCSDOSSIER="-" AND (POG_PERIODICITDUCS="M" OR POG_PERIODICITDUCS="T")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEPERIOD="914" where '+
    'POG_NATUREORG="100" AND POG_DUCSDOSSIER="X" AND (POG_PERIODICITDUCS="M" OR POG_PERIODICITDUCS="T")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEPERIOD="915" where '+
    'POG_NATUREORG="100" AND POG_DUCSDOSSIER="-" AND (POG_PERIODICITDUCS="A")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEPERIOD="916" where '+
    'POG_NATUREORG="100" AND POG_DUCSDOSSIER="X" AND (POG_PERIODICITDUCS="A")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEPERIOD="920" where '+
    'POG_NATUREORG="200" AND POG_DUCSDOSSIER="-" AND (POG_PERIODICITDUCS="M" OR POG_PERIODICITDUCS="T")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEPERIOD="921" where '+
    'POG_NATUREORG="200" AND POG_DUCSDOSSIER="X" AND ( POG_PERIODICITDUCS="M" OR POG_PERIODICITDUCS="T")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEPERIOD="922" where '+
    'POG_NATUREORG="200" AND POG_DUCSDOSSIER="-" AND ( POG_PERIODICITDUCS="A")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEPERIOD="923" where '+
    'POG_NATUREORG="200" AND POG_DUCSDOSSIER="X" AND (POG_PERIODICITDUCS="A")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEPERIOD="930" where '+
    'POG_NATUREORG="300" AND (POG_PERIODICITDUCS="T")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEPERIOD="931" where '+
    'POG_NATUREORG="300" AND (POG_PERIODICITDUCS="M")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEPERIOD="932" where '+
    'POG_NATUREORG="300" AND (POG_PERIODICITDUCS="A")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="913" where '+
    'POG_NATUREORG="100" AND POG_DUCSDOSSIER="-" AND (POG_AUTREPERIODUCS="M" OR POG_AUTREPERIODUCS="T")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="914" where '+
    'POG_NATUREORG="100" AND POG_DUCSDOSSIER="X" AND (POG_AUTREPERIODUCS="M" OR POG_AUTREPERIODUCS="T")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="915" where '+
    'POG_NATUREORG="100" AND POG_DUCSDOSSIER="-" AND (POG_AUTREPERIODUCS="A")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="916" where '+
    'POG_NATUREORG="100" AND POG_DUCSDOSSIER="X" AND (POG_AUTREPERIODUCS="A")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="920" where '+
    'POG_NATUREORG="200" AND POG_DUCSDOSSIER="-" AND (POG_AUTREPERIODUCS="M" OR POG_AUTREPERIODUCS="T")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="921" where '+
    'POG_NATUREORG="200" AND POG_DUCSDOSSIER="X" AND (POG_AUTREPERIODUCS="M" OR POG_AUTREPERIODUCS="T")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="922" where '+
    'POG_NATUREORG="200" AND POG_DUCSDOSSIER="-" AND (POG_AUTREPERIODUCS="A")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="923" where '+
    'POG_NATUREORG="200" AND POG_DUCSDOSSIER="X" AND (POG_AUTREPERIODUCS="A")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="930" where '+
    'POG_NATUREORG="300" AND (POG_AUTREPERIODUCS="T")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="931" where '+
    'POG_NATUREORG="300" AND (POG_AUTREPERIODUCS="M")');
  ExecuteSql('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="932" where '+
    'POG_NATUREORG="300" AND (POG_AUTREPERIODUCS="A")');

  AglNettoieListes('PGPRESENCEFORM', 'PFO_TYPEPLANPREV;PFO_SALARIE;PFO_NOMSALARIE;PFO_PRENOM;PFO_CODESTAGE;PFO_EFFECTUE;PFO_NBREHEURE;PFO_HTPSTRAV;PFO_MILLESIME;PFO_ORDRE;PFO_HTPSNONTRAV;PFO_TYPOFORMATION;PFO_DATEPAIE',nil);

  AglNettoieListes('PGSAISIFORMATION', 'PFO_TYPEPLANPREV;PFO_SALARIE;PFO_NOMSALARIE;PFO_PRENOM;PFO_MILLESIME;PFO_ORDRE;PFO_NBREHEURE;PFO_HTPSTRAV;PFO_HTPSNONTRAV;PFO_TYPOFORMATION;PFO_DATEPAIE',nil);
  // FIN paie
End;

Procedure MajVer699 ;
var
  Usr: string;
  SS : string;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SS := UsDateTime(IDate1900);
  if not IsDossierPCL then
  begin
    //CHR
    // reprise init de 622 CHR
    ExecuteSQL('UPDATE COMMERCIAL SET GCL_HRRANG="" WHERE (GCL_HRRANG IS NULL)');
    ExecuteSQL('UPDATE NOMENLIG SET GNL_QTEMINMENU=0, GNL_QTEMAXMENU=0 , GNL_POIDSMENU=0 WHERE (GNL_QTEMINMENU IS NULL)');
    ExecuteSQL('UPDATE ARTICLECOMPL SET GA2_DEPOT="",GA2_HRQTESTAT="-",GA2_INTERDITACHAT="-" WHERE (GA2_DEPOT IS NULL)');
    ExecuteSQL('UPDATE HRGROUPESSTAT SET HGS_LFAMRES1="",HGS_LFAMRES2="",HGS_LFAMRES3="",HGS_LFAMRES4="",HGS_LFAMRES5="",'+
               'HGS_LFAMRES6="", HGS_LIBELLECHAMP1="",HGS_LIBELLECHAMP2="",HGS_LIBELLECHAMP3="",HGS_LIBELLECHAMP4="",HGS_LIBELLECHAMP5="",HGS_LIBELLECHAMP6=""  WHERE HGS_LFAMRES1 IS NULL');
    ExecuteSQL('UPDATE HRDOSRES SET HDR_CODEFACT="", HDR_NOMBREPERSOFF=0 WHERE HDR_CODEFACT IS NULL');
    ExecuteSQL('UPDATE HRDOSRES SET HDR_LIBELLERESA1="",HDR_LIBELLERESA2="" WHERE HDR_LIBELLERESA1 IS NULL');
    ExecuteSQL('UPDATE HRDOSRES SET HDR_TYPRESASSOC="" WHERE HDR_TYPRESASSOC IS NULL');
    ExecuteSQL('UPDATE HRALLOTEMENT SET HAL_TYPRESASSOC="" WHERE HAL_TYPRESASSOC IS NULL');
    ExecuteSQL('UPDATE HRLIGNE SET HRL_ORDREFACTURE=0 WHERE HRL_ORDREFACTURE IS NULL');
    ExecuteSQL('UPDATE HRDOSSIER SET HDC_FAMILLETAXE1="" WHERE HDC_FAMILLETAXE1 IS NULL');
    { nouveaux champs  }
    Usr := 'CEG';
    ExecuteSQL ('UPDATE HRPREFACT SET HPF_PUHT=0,HPF_PUHTORIGINE=0,HPF_PUTTCORIGINE=0,HPF_FAMILLETAXE1="",' +
      'HPF_FAMILLETAXE2="",HPF_FAMILLETAXE3="",HPF_FAMILLETAXE4="",HPF_FAMILLETAXE5="",HPF_MONTANTHT=0,' +
      'HPF_MONTANTTTC=0,HPF_TOTALTAXE1=0,HPF_TOTALTAXE2=0,HPF_TOTALTAXE3=0,HPF_TOTALTAXE4=0,HPF_TOTALTAXE5=0,' +
      'HPF_CREATEUR="' + Usr + '",HPF_DATECREATION="' + SS + '",HPF_UTILISATEUR="' + Usr + '",HPF_DATEMODIF="' + SS + '" ' +
      'WHERE HPF_FAMILLETAXE1 IS NULL') ;
    ExecuteSQL ('UPDATE HRPARAMPLANNING SET HPP_TAILLECOLENT4=0,HPP_TAILLECOLENT5=0,HPP_TAILLECOLENT6=0,' +
      'HPP_TAILLECOLENT7=0,HPP_TAILLECOLENT8=0,HPP_TAILLECOLENT9=0,HPP_TAILLECOLENTA=0,' +
      'HPP_LIBCOLENT4="",HPP_LIBCOLENT5="",HPP_LIBCOLENT6="",HPP_LIBCOLENT7="",HPP_LIBCOLENT8="",' +
      'HPP_LIBCOLENT9="",HPP_LIBCOLENTA="",HPP_AXESSTOTAUX1="",HPP_AXESSTOTAUX2="",HPP_VISULIGNETRF="" ' +
      'WHERE HPP_TAILLECOLENT4 IS NULL') ;
    ExecuteSQL ('UPDATE HRCONTINGENT SET HCG_ALERTE1=0,HCG_ALERTE2=0,HCG_ALERTE3=0 WHERE HCG_ALERTE1 IS NULL') ;
    ExecuteSQL ('UPDATE HRCONTRAT SET HCO_ETABLISSEMENT="' + GetParamSoc ('SO_ETABLISDEFAUT') + '", HCO_DOMAINE="",' +
      'HCO_CREATEUR="' + Usr + '",HCO_DATECREATION="' + SS + '",HCO_UTILISATEUR="' + Usr + '",HCO_DATEMODIF="' + SS + '",' +
      'HCO_CONFIDENTIEL="0" WHERE HCO_ETABLISSEMENT IS NULL') ;
    ExecuteSQL ('UPDATE HRDOSSIER SET HDC_BOOLLIBRE4="-",HDC_BOOLLIBRE5="-",HDC_BOOLLIBRE6="-",HDC_BOOLLIBRE7="-",' +
      'HDC_BOOLLIBRE8="-",HDC_BOOLLIBRE9="-",HDC_BOOLLIBREA="-",HDC_ETABLISSEMENT="' + GetParamSoc ('SO_ETABLISDEFAUT') + '",' +
      'HDC_DOMAINE="",HDC_CONFIDENTIEL="0",HDC_NATUREDOS="" WHERE HDC_BOOLLIBRE4 IS NULL') ;
    ExecuteSQL ('UPDATE HRDOSRES SET HDR_ETABLISSEMENT="' + GetParamSoc ('SO_ETABLISDEFAUT') + '",HDR_DOMAINE="",' +
      'HDR_CONFIDENTIEL="0",HDR_NOSUITE=0 WHERE HDR_ETABLISSEMENT IS NULL') ;
    ExecuteSQL ('UPDATE HRFAMRES SET HFR_ETABLISSEMENT="' + GetParamSoc ('SO_ETABLISDEFAUT') + '",HFR_DOMAINE="",HFR_NOMBRERES=0 ' +
      'WHERE HFR_ETABLISSEMENT IS NULL') ;
    ExecuteSQL ('UPDATE HRTYPRES SET HTR_ETABLISSEMENT="' + GetParamSoc ('SO_ETABLISDEFAUT') + '",HTR_DOMAINE="" ' +
      'WHERE HTR_ETABLISSEMENT IS NULL') ;
    ExecuteSQL ('UPDATE HRCAISSE SET HRC_COULEURLIBRE=0,HRC_COULEUROUV=0,HRC_COULEURIMP=0,' +
      'HRC_DELAIATTENTE=0 WHERE HRC_COULEURLIBRE IS NULL') ;
    ExecuteSQL ('UPDATE HRNBPERSONNE SET HNP_ETABLISSEMENT="' + GetParamSoc ('SO_ETABLISDEFAUT') + '" ' +
      'WHERE HNP_ETABLISSEMENT IS NULL') ;
    ExecuteSQL ('UPDATE HRLIGNE SET HRL_COMCUISINE="-",HRL_QTEPREM=0,HRL_ARTICLE="" ' +
      'WHERE HRL_COMCUISINE IS NULL') ;
    ExecuteSQL('UPDATE PARCAISSE SET GPK_LFAMRES="" WHERE (GPK_LFAMRES IS NULL)');
    ExecuteSQL('UPDATE ARTICLECOMPL SET GA2_CUISINE1="",GA2_CUISINE2="-",GA2_CUISINE3="-",GA2_CUISINE4="-",GA2_CUISINE5="-" WHERE (GA2_CUISINE1 IS NULL)');
    //DP
    // Requête de réparation pour les liasses (cf YJ/SL)
    ExecuteSQL('update DPFISCAL set DFI_OPTIONAUTID=isnull(DFI_OPTIONAUTID,"-"),  DFI_OPTIONRDSUP=isnull(DFI_OPTIONRDSUP,"-"), DFI_OPTIONRSS=isnull(DFI_OPTIONRSS,"-"), DFI_OPTIONREPORT=isnull(DFI_OPTIONREPORT,"-"), '
      +' DFI_OPTIONRISUP=isnull(DFI_OPTIONRISUP,"-"), DFI_OPTIONEXIG=isnull(DFI_OPTIONEXIG,"-"), DFI_REGLEMENTEURO=isnull(DFI_REGLEMENTEURO,"-"), DFI_TAXESALAIRES=isnull(DFI_TAXESALAIRES,"-"), '
      +' DFI_AUTRESTAXES=isnull(DFI_AUTRESTAXES,"-"), DFI_INTEGRAFISC=isnull(DFI_INTEGRAFISC,"-"), DFI_TETEGROUPE=isnull(DFI_TETEGROUPE,"-"), DFI_EXONERE=isnull(DFI_EXONERE,"-"), DFI_EXONERETP=isnull(DFI_EXONERETP,"-"), '
      +' DFI_ACPTEJUIN=isnull(DFI_ACPTEJUIN,"-"), DFI_DECLA1003R=isnull(DFI_DECLA1003R,"-"), DFI_CAMIONSCARS=isnull(DFI_CAMIONSCARS,"-"), DFI_COTISMIN=isnull(DFI_COTISMIN,"-"), DFI_ALLEGETRANS=isnull(DFI_ALLEGETRANS,"-"), '
      +' DFI_ABATTEFIXE=isnull(DFI_ABATTEFIXE,"-"), DFI_DEGREVTREDUC=isnull(DFI_DEGREVTREDUC,"-"), DFI_ADHEREOGA=isnull(DFI_ADHEREOGA,"-"), DFI_REGIMFUSION=isnull(DFI_REGIMFUSION,"-"), DFI_EXISTEPVLT=isnull(DFI_EXISTEPVLT,"-"), '
      +' DFI_DROITSAPPORT=isnull(DFI_DROITSAPPORT,"-"), DFI_CTRLFISC=isnull(DFI_CTRLFISC,"-"), DFI_DEMATERIATDFC=isnull(DFI_DEMATERIATDFC,"-"), DFI_MONDECLAEURO=isnull(DFI_MONDECLAEURO,"-"), DFI_DECLARATION=isnull(DFI_DECLARATION,"-"), '
      +' DFI_DISTRIBDIVID=isnull(DFI_DISTRIBDIVID,"-"), DFI_TAXEPROF=isnull(DFI_TAXEPROF,"-"), DFI_CONTREVENUSLOC=isnull(DFI_CONTREVENUSLOC,"-"), DFI_TAXEFONCIERE=isnull(DFI_TAXEFONCIERE,"-"), DFI_CONTSOCSOLDOC=isnull(DFI_CONTSOCSOLDOC,"-"), '
      +' DFI_TAXEGRDSURF=isnull(DFI_TAXEGRDSURF,"-"), DFI_TAXEANNIMM=isnull(DFI_TAXEANNIMM,"-"), DFI_TAXEVEHICSOC=isnull(DFI_TAXEVEHICSOC,"-"), DFI_VIGNETTEAUTO=isnull(DFI_VIGNETTEAUTO,"-"), DFI_IMPSOLFORTUNE=isnull(DFI_IMPSOLFORTUNE,"-"), '
      +' DFI_ACTIVFISC=isnull(DFI_ACTIVFISC,""), DFI_NOFRP=isnull(DFI_NOFRP,""), DFI_IMPODIR=isnull(DFI_IMPODIR,""), DFI_REGIMFISCDIR=isnull(DFI_REGIMFISCDIR,""), DFI_IMPOINDIR=isnull(DFI_IMPOINDIR,""), DFI_EXIGIBILITE=isnull(DFI_EXIGIBILITE,""), '
      +' DFI_TYPECA12=isnull(DFI_TYPECA12,""), DFI_PERIODIIMPIND=isnull(DFI_PERIODIIMPIND,""), DFI_MODEPAIEFISC=isnull(DFI_MODEPAIEFISC,""), DFI_NOINTRACOMM=isnull(DFI_NOINTRACOMM,""), DFI_TYPETAXETVA=isnull(DFI_TYPETAXETVA,""), '
      +' DFI_EXONERATION=isnull(DFI_EXONERATION,""), DFI_EXONERATIONTP=isnull(DFI_EXONERATIONTP,""), DFI_TYPETAXEPRO=isnull(DFI_TYPETAXEPRO,""), DFI_ACTIVTAXEPRO=isnull(DFI_ACTIVTAXEPRO,""), DFI_NOADHOGA=isnull(DFI_NOADHOGA,""), '
      +' DFI_NODOSSINSPEC=isnull(DFI_NODOSSINSPEC,""), DFI_NOINSPECTION=isnull(DFI_NOINSPECTION,""), DFI_ORIGMAJBENEF=isnull(DFI_ORIGMAJBENEF,""), DFI_ORIGMAJCA=isnull(DFI_ORIGMAJCA,""), DFI_UTILISATEUR=isnull(DFI_UTILISATEUR,""), '
      +' DFI_CLEINSPECT=isnull(DFI_CLEINSPECT,""), DFI_REGIMEINSPECT=isnull(DFI_REGIMEINSPECT,""), DFI_REDEVABILITE=isnull(DFI_REDEVABILITE,""), DFI_DOMHORSFRANCE=isnull(DFI_DOMHORSFRANCE,""), DFI_BAFORFAIT=isnull(DFI_BAFORFAIT,""), '
      +' DFI_NODP=isnull(DFI_NODP,0), DFI_JOURDECLA=isnull(DFI_JOURDECLA,0), DFI_PRORATATVA=isnull(DFI_PRORATATVA,0), DFI_PRORTVAREVIS=isnull(DFI_PRORTVAREVIS,0), DFI_NOREFTETEGRDP=isnull(DFI_NOREFTETEGRDP,0), '
      +' DFI_MTTDEGREVTRED=isnull(DFI_MTTDEGREVTRED,0), DFI_NOOGADP=isnull(DFI_NOOGADP,0), DFI_RESULTFISC=isnull(DFI_RESULTFISC,0), DFI_REINTEGR=isnull(DFI_REINTEGR,0), DFI_REDUC=isnull(DFI_REDUC,0), DFI_OLDBENEFFISC=isnull(DFI_OLDBENEFFISC,0), '
      +' DFI_OLDREINTEGR=isnull(DFI_OLDREINTEGR,0), DFI_OLDREDUC=isnull(DFI_OLDREDUC,0), DFI_CA=isnull(DFI_CA,0), DFI_ANNEECIVILE=isnull(DFI_ANNEECIVILE,0), DFI_OLDCA=isnull(DFI_OLDCA,0), DFI_BENEFFISC=isnull(DFI_BENEFFISC,0), '
      +' DFI_DATEOPTRDSUP=isnull(DFI_DATEOPTRDSUP,"'+UsDateTime(iDate1900)+'"), DFI_DATEOPTRSS=isnull(DFI_DATEOPTRSS,"'+UsDateTime(iDate1900)+'"), DFI_DATEOPTREPORT=isnull(DFI_DATEOPTREPORT,"'+UsDateTime(iDate1900)+'"), '
      +' DFI_DATEOPTEXIG=isnull(DFI_DATEOPTEXIG,"'+UsDateTime(iDate1900)+'"), DFI_DATEPRORTVA=isnull(DFI_DATEPRORTVA,"'+UsDateTime(iDate1900)+'"), DFI_DATEDEBEXO=isnull(DFI_DATEDEBEXO,"'+UsDateTime(iDate1900)+'"), '
      +' DFI_DATEFINEXO=isnull(DFI_DATEFINEXO,"'+UsDateTime(iDate1900)+'"), DFI_DATEDEBEXOTP=isnull(DFI_DATEDEBEXOTP,"'+UsDateTime(iDate1900)+'"), DFI_DATEFINEXOTP=isnull(DFI_DATEFINEXOTP,"'+UsDateTime(iDate1900)+'"), '
      +' DFI_DATEADHOGA=isnull(DFI_DATEADHOGA,"'+UsDateTime(iDate1900)+'"), DFI_DATEREGFUS=isnull(DFI_DATEREGFUS,"'+UsDateTime(iDate1900)+'"), DFI_DATEAPPORT=isnull(DFI_DATEAPPORT,"'+UsDateTime(iDate1900)+'"), '
      +' DFI_DATECONVTDFC=isnull(DFI_DATECONVTDFC,"'+UsDateTime(iDate1900)+'"), DFI_DATEMAJBENEF=isnull(DFI_DATEMAJBENEF,"'+UsDateTime(iDate1900)+'"), DFI_DATECREATION=isnull(DFI_DATECREATION,"'+UsDateTime(iDate1900)+'"), '
      +' DFI_DATEMODIF=isnull(DFI_DATEMODIF,"'+UsDateTime(iDate1900)+'"), DFI_DATEOPTRG=isnull(DFI_DATEOPTRG,"'+UsDateTime(iDate1900)+'")');
    //GIGA
     ExecuteSql('UPDATE AFFAIRE set AFF_SEUILINFO = 0');
     // GPAO
    InsertChoixCode('YTN', 'H', 'des heures', 'Heures', '') ;
    ExecuteSQL('UPDATE PORT SET GPO_TIERS="",GPO_VENTPIECE="",GPO_VENTCOMPTA="",GPO_REPARTITION="",GPO_FACTURABLE="X",GPO_TYPEPDR="", GPO_TYPEFOURNI=""  WHERE GPO_TIERS IS NULL');
    ExecuteSQL('UPDATE YTARIFS SET YTS_CODEPORT="",YTS_NATURETRAVAIL="",YTS_RESSOURCE="",YTS_TARIFRESSOURCE="",YTS_ATTRIBUTION="", YTS_CASCRESSOURCE="-",yts_casctarifresso="-",YTS_CASCTOUSRESSO="-" WHERE YTS_CODEPORT IS NULL');
    ExecuteSQL('UPDATE PIEDPORT SET GPT_MONTANTREALISE=0 WHERE GPT_MONTANTREALISE IS NULL');
    // ExecuteSQL('UPDATE LIGNECOMPL SET GLC_GUIDGSM="" WHERE GLC_GUIDGSM IS NULL');
    ExecuteSQL('UPDATE YTARIFSPARAMETRES SET YFO_CODEPORT="",YFO_OKCODEPORT="----",YFO_OKNATURETRA="----", YFO_OKRESSOURCE="----" WHERE YFO_CODEPORT IS NULL');
    ExecuteSQL('UPDATE PORT SET GPO_TYPEFRAIS="FAN" WHERE GPO_TYPEPORT="TRA"');
    ExecuteSQL('UPDATE PORT SET GPO_TYPEFRAIS="CIN" WHERE GPO_TYPEFRAIS=""');
    ExecuteSQL('UPDATE PORT SET GPO_TYPEPORT="TAR" WHERE GPO_TYPEPORT="TRA"');
    ExecuteSQL('UPDATE PORT SET GPO_VENTPIECE="P", GPO_REPARTITION="Q" WHERE GPO_TYPEPORT="TAR" AND GPO_VENTPIECE IS NULL');
    ExecuteSQL('UPDATE TIERSCOMPL SET YTC_TYPEFOURNI="" WHERE YTC_TYPEFOURNI IS NULL');
    AGLNettoieListes('YTARIFSMULART201','YTC_CODEPORT;YTC_NATURETRAVAIL;YTC_RESSOURCE;YTC_TARIFRESSOURCE',nil);
    AGLNettoieListes('YTARIFSMULART211','YTC_CODEPORT;YTC_NATURETRAVAIL;YTC_RESSOURCE;YTC_TARIFRESSOURCE',nil);
    AGLNettoieListes('YTARIFSMULART301','YTC_CODEPORT;YTC_NATURETRAVAIL;YTC_RESSOURCE;YTC_TARIFRESSOURCE',nil);
    AGLNettoieListes('YTARIFSMULART401','YTC_CODEPORT;YTC_NATURETRAVAIL;YTC_RESSOURCE;YTC_TARIFRESSOURCE',nil);
    AGLNettoieListes('YTARIFSMULART501','YTC_CODEPORT;YTC_NATURETRAVAIL;YTC_RESSOURCE;YTC_TARIFRESSOURCE',nil);
    AGLNettoieListes('YTARIFSMULTIE201','YTC_CODEPORT;YTC_NATURETRAVAIL;YTC_RESSOURCE;YTC_TARIFRESSOURCE',nil);
    AGLNettoieListes('YTARIFSMULTIE211','YTC_CODEPORT;YTC_NATURETRAVAIL;YTC_RESSOURCE;YTC_TARIFRESSOURCE',nil);
    AGLNettoieListes('YTARIFSMULTIE301','YTC_CODEPORT;YTC_NATURETRAVAIL;YTC_RESSOURCE;YTC_TARIFRESSOURCE',nil);
    AGLNettoieListes('YTARIFSMULTIE401','YTC_CODEPORT;YTC_NATURETRAVAIL;YTC_RESSOURCE;YTC_TARIFRESSOURCE',nil);
    AGLNettoieListes('YTARIFSMULTIE501','YTC_CODEPORT;YTC_NATURETRAVAIL;YTC_RESSOURCE;YTC_TARIFRESSOURCE',nil);
    AGLNettoieListes('RTMULTIERSRECH','YTC_TYPEFOURNI',nil);
    // GC
    ExecuteSQL('UPDATE LISTEINVLIG SET GIL_REFERENCE = "", GIL_FOURNISSEUR = "", GIL_CONTREMARQUE = "-"');
    //GPAO
    //ExecuteSQL('UPDATE paramsoc SET soc_data="001" WHERE soc_nom = "SO_GCGESTUNITEMODE"');
    //ExecuteSQL('UPDATE paramSoc SET soc_data="PIE" WHERE soc_nom = "SO_GCDEFQUALIFUNITEGA"');
    //ExecuteSQL('UPDATE paramSoc SET soc_data="UNI" WHERE soc_nom = "SO_GCDEFUNITEGA"');
    ExecuteSQL('UPDATE winitchamplig SET wil_famchamp="DIV" WHERE wil_nomchamp = "GA_UNITEQTEVTE"');
    ExecuteSQL('UPDATE PORT SET GPO_TIERS="",GPO_VENTPIECE="A",GPO_VENTCOMPTA="P",GPO_REPARTITION="Q",GPO_FACTURABLE="X",GPO_TYPEPDR="", GPO_TYPEFOURNI=""  WHERE GPO_TIERS IS NULL');
    // La requète ci dessous remplace la requète d'init. suivante :
    ExecuteSQL('UPDATE PORT SET GPO_TIERS="",GPO_VENTPIECE="",GPO_VENTCOMPTA="",GPO_REPARTITION="",GPO_FACTURABLE="X",GPO_TYPEPDR="", GPO_TYPEFOURNI=""  WHERE GPO_TIERS IS NULL');
    //GIGA
    ExecuteSQL('UPDATE FACTAFF SET AFA_DATEPIECE="' + SS + '"');
  end;
  ExecuteSQL('update journal set j_tresovalid="-",j_tresolibre="-"');
  ExecuteSQL('update jalref set jr_tresoimport="",jr_tresochainage="-",jr_tresomontant="-",jr_tresodate="-",jr_tresolibelle="-",jr_tresovalid="-",jr_tresolibre="-"');
End;

Procedure MajVer700 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //DP
   // correction 27092005 MD ExecuteSQL('update DPSOCIAL set DSO_PERURSSAF="", DSO_PERASSEDIC="", DSO_PERIRC="", DSO_REMBULLPAIE="", DSO_DECLHANDICAP="", DSO_DECLMO=""');
    ExecuteSQL('update DPSOCIAL set DSO_PERURSSAF="", DSO_PERASSEDIC="", DSO_PERIRC="", DSO_REMBULLPAIE=0, DSO_DECLHANDICAP="-", DSO_DECLMO="-"');
    ExecuteSQL('update DPORGA set DOR_WMODETRAITCPT="", DOR_WTYPEBNC=""');
    ExecuteSQL('UPDATE CONTACT SET C_CLEFAX="",C_CLETELEX=""');
    //GIGA
    executeSQL ('UPDATE AFVALINDICE set AFV_BASE100="-"');
    //GPAO
    if (V_PGI.Driver <> dbORACLE7 ) then // pas supporté en DB2
        MajPoidsTarifs;
    //SAV
    executeSQL ('update article set ga_articleparc="" ');
    // MNG 07092005 W_SAV_ArticlesParc;
    executeSQL ('update actions set rac_identparc=0 ');
    RT_InsertInfoCompl;
   end;
  //PAIE
  AglNettoieListes('PGGENECRCOMPTA', 'PSA_DATESORTIE',nil);  // Rajout du champ PSA_DATESORTIE comme champ obligatoire

End;

Procedure MajVer701 ;
var SS : string;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SS := UsDateTime(IDate1900);
  if not IsDossierPCL then
  begin
    executeSQL ('update RTINFOS008 set RD8_RD8LIBTABLE5="",RD8_RD8LIBTABLE6="",RD8_RD8LIBTABLE7="",RD8_RD8LIBTABLE8="",RD8_RD8LIBTABLE9="",RD8_RD8LIBMUL5="",RD8_RD8LIBMUL6="",RD8_RD8LIBMUL7="",RD8_RD8LIBMUL8="",RD8_RD8LIBMUL9=""');
    executeSQL ('update RTINFOS004 set RD4_RD4LIBTABLE5="",RD4_RD4LIBTABLE6="",RD4_RD4LIBTABLE7="",RD4_RD4LIBTABLE8="",RD4_RD4LIBTABLE9="",RD4_RD4LIBMUL5="",RD4_RD4LIBMUL6="",RD4_RD4LIBMUL7="",RD4_RD4LIBMUL8="",RD4_RD4LIBMUL9="" ');
    RT_InsertLibelleInfoComplDixArtLigne;
    //CHR
    ExecuteSQL('UPDATE PAYS SET PY_CODEINSEE="" WHERE (PY_CODEINSEE IS NULL)');
    //DP
    ExecuteSQL('UPDATE YMESSAGES SET YMS_INBOX="-"');
    ExecuteSQL('UPDATE DOSSIER SET DOS_WINSTALL="", DOS_TYPEMAJ="", DOS_DETAILMAJ=""');
    ExecuteSQL('UPDATE YMESSAGES SET YMS_INBOX="X" WHERE YMS_BROUILLON="X"');

    // GPAO
    ExecuteSQL('UPDATE PIEDPORT SET GPT_FACTURABLE="X" WHERE GPT_FACTURABLE IS NULL');
    ExecuteSQL('UPDATE WPARAMFONCTION SET WPF_USER="", WPF_CONTEXTE="" WHERE WPF_USER IS NULL');
    ExecuteSQL('UPDATE STKMOUVEMENT SET GSM_NUMPREPAORI=0 WHERE GSM_NUMPREPAORI IS NULL');
    ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA="NON" WHERE SOC_NOM="SO_FRAISINDPR" AND SOC_DATA="-"');
    ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA="THE" WHERE SOC_NOM="SO_FRAISINDPR" AND SOC_DATA="X"');
    ExecuteSQL('UPDATE WORDRERES SET WOR_RESSOURCESCM="-" WHERE WOR_RESSOURCESCM IS NULL');
    ExecuteSQL('UPDATE WGAMMERES SET WGR_RESSOURCESCM="-" WHERE WGR_RESSOURCESCM IS NULL');
    ExecuteSQL('UPDATE WGAMMELIG SET WGL_RESSOURCESCM="-",WGL_TYPEOPE="" WHERE WGL_RESSOURCESCM IS NULL');
    ExecuteSQL('UPDATE WORDREGAMME SET WOG_RESSOURCESCM="-",WOG_TYPEOPE="" WHERE WOG_RESSOURCESCM IS NULL');
    ExecuteSQL('UPDATE ARTICLE SET GA_CODEFORME="PIE" WHERE ISNULL(GA_CODEFORME,"") = ""');
    //GC
    ExecuteSQL('UPDATE COMMERCIAL SET GCL_COMMENCAISS ="-" ');
    // Noyau
    YY_InitChampsSystemes;
    //GIGA
    ExecuteSql ('Update Tache set ata_nbpartprevu=0,ata_nbpartinscrit=0');
 end;
// Noyau
ExecuteSql ('update UTILISAT SET US_NBTENTATIVE=0,US_NBCONNEXION=0,US_DATECHANGEPWD="'+SS+'"'
    +',US_OLDPWD1="",US_OLDPWD2="",US_OLDPWD3="",US_DESACTIVE="-",US_MULTICONNEXION="X"'
    +',US_DOMAINEAUTHENT="",US_AUTHENTNT="-",US_SSO="-",US_AUTHENTFORTE="-",US_AUTHENTPARAM=""');
 // PAIE
ExecuteSql('UPDATE MOTIFABSENCE SET PMA_PGCOLORVAL = "",PMA_PGCOLORATT = "",PMA_PGCOLORREF = "",PMA_PGCOLORANN = ""');

End;

Procedure MajVer702 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //giGA
    ExecuteSql('update factaff set afa_etatvisa="VIS" where afa_typeche="NOR" and afa_etatvisa<>"VIS"');
    //GRC
    ExecuteSQL('UPDATE PARPIECE SET GPP_INFOSCPLPIECE="-"');
    ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA = "<<aucun>>" WHERE SOC_NOM = "SO_QUALIFMVTWPC"');
    InsertChoixCode('RGD', 'RD1', '.-Table libre 1', 'Table libre 1', '');
    InsertChoixCode('RGD', 'RD2', '.-Table libre 2', 'Table libre 2', '');
    InsertChoixCode('RGD', 'RD3', '.-Table libre 3', 'Table libre 3', '');
    //GPAO
//    ExecuteSQL('UPDATE CATALOGU SET GCA_QUALIFUNITEACH="PIE" WHERE ISNULL(GCA_QUALIFUNITEACH, "") = ""');
    if ExisteSQL('SELECT 1 FROM CATALOGU WHERE ISNULL(GCA_QUALIFUNITEACH, "") = ""') then
    begin
      if ExisteSQL('SELECT 1 FROM MEA WHERE GME_MESURE = "UNI"') then
        ExecuteSQL('UPDATE CATALOGU SET GCA_QUALIFUNITEACH="UNI" WHERE ISNULL(GCA_QUALIFUNITEACH, "") = ""')
      else if ExisteSQL('SELECT 1 FROM MEA WHERE GME_MESURE = "U"') then
        ExecuteSQL('UPDATE CATALOGU SET GCA_QUALIFUNITEACH="U" WHERE ISNULL(GCA_QUALIFUNITEACH, "") = ""')
      else
      begin
        if not ExisteSQL('SELECT 1 FROM MEA WHERE GME_MESURE = "UNI"') then
          ExecuteSQL('INSERT INTO MEA (GME_QUALIFMESURE,GME_MESURE,GME_LIBELLE,GME_QUOTITE) VALUES ("PIE", "UNI", "Unité", 1)');
        ExecuteSQL('UPDATE CATALOGU SET GCA_QUALIFUNITEACH="UNI" WHERE ISNULL(GCA_QUALIFUNITEACH, "") = ""')
      end;
    end;
    MaJdesPrixGSMFromBlocNote;
    AjouteNewTarifs;
    //GC
    ExecuteSQL('update parpiece set gpp_qualifmvt="RCC" where gpp_naturepieceg="CC" ');
    ExecuteSQL('update parpiece set gpp_qualifmvt="SCC" where gpp_naturepieceg="BLC" ');
    ExecuteSQL('update parpiece set gpp_qualifmvt="SCC" where gpp_naturepieceg="FAC" ');
    ExecuteSQL('update parpiece set gpp_qualifmvt="ECC" where gpp_naturepieceg="BLF" ');
    ExecuteSQL('update parpiece set gpp_qualifmvt="ECC" where gpp_naturepieceg="FF"  ');
    ExecuteSQL('update parpiece set gpp_qualifmvt="ACC" where gpp_naturepieceg="CF"  ');
    ExecuteSQL('update parpiece set gpp_qualifmvt="RCC" where gpp_naturepieceg="PRE" ');
    ExecuteSQL('update parpiece set gpp_qualifmvt="ACC" where gpp_naturepieceg="PRF" ');
    //PAIE
    ExecuteSql('UPDATE ABSENCESALARIE SET PCN_REPRISEARRET="'+UsDateTime (IDate1900)+'", PCN_DATEATTEST="'+UsDateTime (IDate1900)+'"');
    AglNettoieListes('PGBILANSOCIAL','PBC_DATEDEBUT;PBC_DATEFIN;PBC_BSPRESENTATION;PBC_ETABLISSEMENT;PBC_BSINDTHEME;(PBC_INDICATEURBS);PBC_INDICATEURBS;PBC_CATBILAN;PBC_VALCAT;PBC_PGINDICATION;PBC_NBELEMENT;PBC_NODOSSIER;PBC_INCREM',nil);
    AglNettoieListes('PGBSINDSELECT','PID_PREDEFINI;PID_NODOSSIER;PID_BSPRESENTATION;PID_LIBELLE;PID_PGPERIODICITE',nil);
    AglNettoieListes('PBSINDICATEURS','PBI_TYPINDICATBS;PBI_INDICATEURBS;PBI_NODOSSIER;PBI_LIBELLE;PBI_PREDEFINI;PBI_BSINDTHEME;PBI_FOURCHETTEMINI;PBI_FOURCHETTEMAXI',nil);
    ExecuteSql('UPDATE DUCSENTETE SET PDU_MASSEANNUEL = 0, PDU_REMUNDADS = 0');
    ExecuteSql('UPDATE CONVENTIONCOLL SET PCV_IDCC="",PCV_NUMERODIV="",PCV_LIENCCN=""');
    //CHR
    AglNettoieListes('HLMULREGLEMENT','', Nil);
  end;
End;

Procedure MajVer703 ;
var stCle : string;
    i : integer;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //RT
    if GetParamSoc ('SO_RTPROPTABHIE') then
    begin
      for i := 1 to 3 do
      begin
        stCle := 'RTRPRLIBPERSPECTIVE';
        ExecuteSQL('UPDATE YDATATYPETREES SET YDT_CODEHDTLINK ="'+'GC'+stCle+IntToStr(i)+'" WHERE YDT_CODEHDTLINK="'+stCle+IntToStr(i)+'"');
        ExecuteSQL('UPDATE YDATATYPELINKS SET YDL_CODEHDTLINK ="'+'GC'+stCle+IntToStr(i)+'",YDL_PREDEFINI="CEG" WHERE YDL_CODEHDTLINK="'+stCle+IntToStr(i)+'"');
      end;
    end;
    if GetParamSoc ('SO_RTACTTABHIE') then
    begin
      for i := 1 to 3 do
      begin
        stCle := 'RTRPRLIBACTION';
        ExecuteSQL('UPDATE YDATATYPETREES SET YDT_CODEHDTLINK ="'+'GC'+stCle+IntToStr(i)+'" WHERE YDT_CODEHDTLINK="'+stCle+IntToStr(i)+'"');
        ExecuteSQL('UPDATE YDATATYPELINKS SET YDL_CODEHDTLINK ="'+'GC'+stCle+IntToStr(i)+'",YDL_PREDEFINI="CEG" WHERE YDL_CODEHDTLINK="'+stCle+IntToStr(i)+'"');
      end;
    end;
    //GPAO
    ExecuteSQL('UPDATE ARTICLECOMPL SET GA2_DATEMODIFBT="'+UsDateTime(iDate1900)+'" WHERE GA2_DATEMODIFBT IS NULL');
    ExecuteSQL('UPDATE USERGRP SET UG_ENVIRONNEMENT="" WHERE UG_ENVIRONNEMENT IS NULL');
    ExecuteSQL('UPDATE DISPO SET GQ_DELAIMOYEN=0 WHERE GQ_DELAIMOYEN IS NULL');
    ExecuteSQL('UPDATE QSIMULATION SET QSM_MODESIMU="" WHERE QSM_MODESIMU IS NULL');
    ExecuteSQL('UPDATE QARTTECH SET QAR_INACTIF="-" WHERE QAR_INACTIF IS NULL');
    ExecuteSQL('UPDATE QBPSESSIONBP SET QBS_OKOBJPREV="-", QBS_OBJPREV="X" WHERE QBS_OBJPREV IS NULL');
    ExecuteSQL('UPDATE WPDRLIG SET WPL_CONSOLIDE="-" WHERE WPL_CONSOLIDE IS NULL');
    ExecuteSQL('UPDATE STKNATURE SET GSN_MAJPRIXVALO="" WHERE GSN_MAJPRIXVALO IS NULL');
    ExecuteSQL('UPDATE STKNATURE SET GSN_MAJPRIXVALO="DPA;DPR;PPA;PPR" WHERE GSN_QUALIFMVT IN ("CLO","EAC","ECP","EDM","EEX","EIN","EIT","EPR","ETR","000")');
    ExecuteSQL('UPDATE STKNATURE SET GSN_MAJPRIXVALO="DPA;DPR" WHERE GSN_QUALIFMVT IN ("AAC","APR","ATD")');
    ExecuteSQL(' UPDATE RESSOURCE SET ARS_RESSOURCESCM="-",ARS_GRP="" '
          + ',ARS_CHAINEORDO="",ARS_PHASE=""'
          + ',ARS_GESTGROUPE="",ARS_EFFECTIF=0'
          + ',ARS_ORDREAFF=0,ARS_ORDRECALC=0'
          + ',ARS_DUREELUN="",ARS_DUREEMAR=""'
          + ',ARS_DUREEMER="",ARS_DUREEJEU=""'
          + ',ARS_DUREEVEN="",ARS_DUREESAM=""'
          + ',ARS_DUREEDIM="",ARS_PROFILLUN=""'
          + ',ARS_PROFILMAR="",ARS_PROFILMER=""'
          + ',ARS_PROFILJEU="",ARS_PROFILVEN=""'
          + ',ARS_PROFILSAM="",ARS_PROFILDIM=""'
          + ',ARS_OUVRELUN="-",ARS_OUVREMAR="-"'
          + ',ARS_OUVREMER="-",ARS_OUVREJEU="-"'
          + ',ARS_OUVREVEN="-",ARS_OUVRESAM="-"'
          + ',ARS_OUVREDIM="-"'
          +' WHERE ARS_RESSOURCESCM IS NULL');
    AGLNettoieListes('YTARIFSMULART101','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULART201','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULART211','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULART301','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULART401','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULART501','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);

    AGLNettoieListes('YTARIFSMULTIE101','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULTIE201','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULTIE211','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULTIE301','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULTIE401','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
    AGLNettoieListes('YTARIFSMULTIE501','YTS_ACTIF;YTS_DATEREFERENCE;YTS_SITE;YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT',nil);
  end;
//PAIE
ExecuteSql('UPDATE HISTOBULLETIN SET PHB_ORDREETAT = 3 WHERE PHB_NATURERUB <> "AAA"');
ExecuteSql('UPDATE HISTOANALPAIE SET PHA_ORDREETAT = 3 WHERE PHA_NATURERUB <> "AAA"');
ExecuteSql('UPDATE histobulletin set phb_omtsalarial = phb_ordreetat');
ExecuteSql('UPDATE histoanalpaie set pha_omtsalarial = pha_ordreetat');

//PDumet 12102005
v_pgi.enableDEShare := True;
V_PGI.StandardSurDP := True;

ExecuteSql('UPDATE HISTOANALPAIE SET PHA_OMTSALARIAL=(SELECT PCT_ORDREETAT FROM COTISATION WHERE ##PCT_PREDEFINI## PCT_RUBRIQUE=PHA_RUBRIQUE) '
       +' WHERE PHA_NATURERUB<>"AAA" AND PHA_RUBRIQUE IN (SELECT PCT_RUBRIQUE FROM COTISATION'
                            +' COT WHERE ##COT.PCT_PREDEFINI## COT.PCT_THEMECOT="RDC")');
ExecuteSql('UPDATE HISTOBULLETIN SET PHB_OMTSALARIAL=(SELECT PCT_ORDREETAT FROM COTISATION '
                +' WHERE ##PCT_PREDEFINI## PCT_RUBRIQUE=PHB_RUBRIQUE) WHERE PHB_NATURERUB<>"AAA" '
                +' AND PHB_RUBRIQUE IN (SELECT PCT_RUBRIQUE FROM COTISATION COT '
                +'  WHERE ##COT.PCT_PREDEFINI## COT.PCT_THEMECOT="RDC")');

//PDumet 12102005
v_pgi.enableDEShare := False;
V_PGI.StandardSurDP := False;

End;

Procedure MajVer704 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GIGA
    AglNettoieListes('AFMULPLANCHARGE', 'AFF_TIERS',nil);
    //GPAO
    AGLNettoieListes('GCSTKDISPODETAIL','GQD_CONTREMARQUE;GQD_REFERENCE;GQD_FOURNISSEUR',nil);
    AGLNettoieListes('YTARIFSMULART101;YTARIFSMULART201;YTARIFSMULART211;YTARIFSMULART301;YTARIFSMULART401','YTS_CODEPORT',nil);
    AGLNettoieListes('YTARIFSMULTIE101;YTARIFSMULTIE201;YTARIFSMULTIE211;YTARIFSMULTIE301;YTARIFSMULTIE401','YTS_CODEPORT',nil);
    ExecuteSql('UPDATE STKREGLEGESTION SET GSR_ORDRESELECT="-GQD_REFAFFECTATION;+GQD_DATEENTREELOT;+GQD_DATEPEREMPTION;+GQD_INDICEARTICLE;+GQD_LOTINTERNE;+GQD_SERIEINTERNE;+GQD_EMPLACEMENT" WHERE GSR_REGLEGESTION="FIFO"');
    ExecuteSql('UPDATE STKREGLEGESTION SET GSR_ORDRESELECT="-GQD_REFAFFECTATION;-GQD_DATEENTREELOT;+GQD_DATEPEREMPTION;+GQD_INDICEARTICLE;+GQD_LOTINTERNE;+GQD_SERIEINTERNE;+GQD_EMPLACEMENT" WHERE GSR_REGLEGESTION="LIFO"');
    ExecuteSql('UPDATE STKREGLEGESTION SET GSR_ORDRESELECT="-GQD_REFAFFECTATION;+GQD_DATEPEREMPTION;+GQD_DATEENTREELOT;+GQD_INDICEARTICLE;+GQD_LOTINTERNE;+GQD_SERIEINTERNE;+GQD_EMPLACEMENT" WHERE GSR_REGLEGESTION="FEFO"');

    if (V_PGI.driver <> dbORACLE7) then RecupParamCBN;

    //GC
    AglNettoieListes('GCMULARCHIVAGEMVT','',nil);

  end;
End;

Procedure MajVer705 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GIGA
    AglNettoieListes('YYMULSELDOSS', '',nil);
    AglNettoieListes('AFMULPROPOSITION','AFF_AVENANT',nil);
    AglNettoieListes('AFMULLIGNEACOMPTE','GL_NUMORDRE',nil);
    AglNettoieListes('GCSTKDISPODETAIL','GA_QUALIFUNITESTO;GA_QUALIFUNITEVTE;GA_UNITEQTEVTE;GA_UNITEQTEACH;GA_UNITEPROD;GA_UNITECONSO',nil);
    //GRC
    InsertChoixCode ('R5Z','BLO','Bloc-notes ','','');
    InsertChoixCode ('R7Z','BLO','Bloc-notes ','','');
    // pour modifier un paramètrage par défaut trop serré (timer)
    If IsMonoOuCommune and (V_PGI.ModePCL='1') then
       if GetParamSoc('SO_MDMESSAGERIETIMER')=10 then SetParamSoc('SO_MDMESSAGERIETIMER', 120);
    // GC
    ExecuteSql('UPDATE CHOIXCOD SET CC_TYPE="ZLA" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "A%"' );
    // changé DBR le 29082005
    ExecuteSql('UPDATE CHOIXCOD SET CC_TYPE="ZLT" WHERE CC_TYPE="ZLI" AND ((CC_CODE LIKE "C%" AND CC_CODE NOT LIKE "CP%") OR CC_CODE LIKE "F%")');

  end;

End;

Procedure MajVer706 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    AGLNettoieListes('WPDRTET','WPE_UNPDRPAR',nil);
    //GPAO
    InsertChoixCode ('GFT','TRA','Transporteur','Transporteur','');
  end;
End;

Procedure MajVer707 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GIGA
      AglNettoieListesPlus('AFMULPROPOSITION', 'AFF_AVENANT',nil,true);
      ExecuteSql ('DELETE FROM CHOIXCOD WHERE CC_TYPE="ATU" AND CC_CODE in ("030","031","032","033")');
      InsertChoixCode('ATU', '030', 'opérations commerciales', '', 'opérations marketing');
      InsertChoixCode('ATU', '031', 'opération commerciale', '', 'opération marketing');
    //fin GIGA
    //GPAO
    ExecuteSql('UPDATE STKNATURE SET GSN_SDISPOPICKING="LBR;BLQ;PER;" WHERE GSN_QUALIFMVT="SEX"');
    //GRC
    RT_InsertTablettesSAV;
    ExecuteSQL('DELETE FROM CHOIXCOD WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "W%"');
    W_SAV_ArticlesParc;
  end;

  //SYSTEM
  ExecuteSql('UPDATE UTILISAT SET US_GRPSDELEGUES="" where US_GRPSDELEGUES IS NULL');

End;

Procedure MajVer708 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    // mise a jour des listes du suivi de contremarque
    // DBR 21092005 AGLNettoieListes ('GCMULCONTREAVTGD;GCMULCONTREAVTGT', '', nil);
    //initialisation nouveaux champs STKNATURE
    ExecuteSql('UPDATE STKNATURE SET GSN_SDISPOPICKING="AFF;", GSN_SFLUXPICKING="STD;PVE;" WHERE GSN_QUALIFMVT="SCC"');
    //GC
    ExecuteSql('UPDATE STKNATURE SET GSN_LIBELLE="Sortie exceptionnelle" where gsn_qualifmvt="SEX"');
    ExecuteSql('UPDATE STKNATURE SET GSN_LIBELLE="Entrée exceptionnelle" where gsn_qualifmvt="EEX"');
  end;
End;

Procedure MajVer709 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    AGLNettoieListes('WCBNPREVTET','WPT_AFFAIRE',nil);
    ExecuteSql('UPDATE CHOIXCOD SET CC_TYPE="GED" WHERE CC_TYPE="GCE"');
  end;
End;

Procedure MajVer710 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
       AGLNettoieListes('WCBNPREVTET','WPT_AFFAIRE',nil);
       //GPAO
      ExecuteSQL('DELETE FROM PARPIECE WHERE GPP_NATUREPIECEG LIKE "W%"');
      ExecuteSQL('UPDATE PARPIECE SET GPP_INSERTLIG="X",GPP_ACTIONFINI="ENR",GPP_STKQUALIFMVT="STR",GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG = "TEM"');
      ExecuteSQL('UPDATE PARPIECE SET GPP_INSERTLIG="-",GPP_ACTIONFINI="ENR",GPP_STKQUALIFMVT="ETR",GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG = "TRE"');
      ExecuteSQL('UPDATE PARPIECE SET GPP_INSERTLIG="-",GPP_ACTIONFINI="TRA",GPP_STKQUALIFMVT="ETR",GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG = "TRV"');
      ExecuteSQL('UPDATE PARPIECE SET GPP_CONTEXTES="GP" WHERE GPP_NATUREPIECEG IN ("BSA","BSP","CSA","CSP")');
  end;
//TRESO
AglNettoieListes('TRECRITURERAPPRO', 'TE_NATURE', nil);
//

End;

Procedure MajVer711 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GP GC
    if (V_PGI.driver <> dbORACLE7) and not (V_PGI.ModePCL='1') then
    begin
      if V_PGI.SAV then LogAGL('Début Init Stock ContreMarque' + DateTimeToStr(Now));
      CallSTKMoulinetteContreMarque(True);
      if V_PGI.SAV then LogAGL('Fin Init Stock ContreMarque' + DateTimeToStr(Now));
    end;

    AGLNettoieListesPlus('GCMULMODIFDOCACH','GP_CDETYPE',nil, True);
    AGLNettoieListesPlus('GCMULMODIFDOCVEN','GP_CDETYPE',nil, True);
    AGLNettoieListesPlus('GCMULLIGNEACH','GL_TYPECADENCE',nil, True);
    AGLNettoieListesPlus('GCDUPLICPIECE','GP_CDETYPE',nil, True);

    {-- MISE A JOUR DE STKNATURE --}
    ExecuteSQL('DELETE FROM STKNATURE WHERE GSN_QUALIFMVT="TAV"');
    ExecuteSQL('UPDATE STKNATURE SET GSN_SIGNEMVT="SOR" WHERE GSN_QUALIFMVT="STR"');
    //Compta
    // RR 11102005 AglNettoieListes('SSAISODA','Y_EXERCICE;Y_NUMEROPIECE;Y_DATECOMPTABLE;Y_NUMVENTIL;Y_AXE,Y_SECTION;Y_REFINTERNE;Y_LIBELLE;Y_DEBIT;Y_CREDIT;');
  end;
End;


Procedure MajVer712 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
      //GIGA
      // MCD 19102005  AglNettoieListes('AFALIGNAFF', 'YTC_REPRESENTANT2;YTC_REPRESENTANT3',nil);
      //GC
      ExecuteSql('UPDATE parpiece set gpp_contextes="AFF;GCA;" where gpp_naturepieceg in ("AFF","PAF","FPR","APR")');
      ExecuteSql('UPDATE parpiece set gpp_contextes="AFF;GC;" where gpp_naturepieceg in ("FF","AF","FAC","CF","BLF","AVC")');
      ExecuteSql('UPDATE parpiece set gpp_contextes="GC;" where gpp_naturepieceg in ("AFP","AFS","AVS", "BFA","BLC","BRC","CC","CCE","DE","FFA","FFO","FRA","LCE","PRE","PRF","PRO","REA","TEM","TRE","TRV")');

      If IsMonoOuCommune and (V_PGI.ModePCL='1') then
      begin // mise à jour table tiers si base 00 et PCL
        ExecuteSql('update tiers set t_tiers=t_auxiliaire where t_tiers=""');
        ExecuteSql('update tierscompl set ytc_tiers=ytc_auxiliaire where ytc_tiers=""');
      end;
  end;

End;

Procedure MajVer713 ;
var
  Q: tQuery;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
  	//GIGA
    AGLNettoieListesPlus('AFMUL_EACTIVITE2','EAC_TOTPRCHARGE',nil, True);
    AGLNettoieListesPlus('AFMULFRAISCOMPTA','ACT_TOTPRCHARGE',nil, True);
    //GC
    { Formatage des champs lien pièce }
    Q := OpenSQL('SELECT COUNT(*)  FROM LIGNE WHERE GL_ARTICLE<>""', True);
    try
      if (not Q.Eof) and (Q.Fields[0].AsInteger > 0) and (Q.Fields[0].AsInteger < 2000000) then
        { On lance le reformatage des ligne uniquement si on a moins de 2millions de lignes}
        MoulinetteFormatageChampsLienPiece(True);
    finally
      Ferme(Q);
    end;
    ExecuteSQL('UPDATE ANNULIEN SET ANL_RACINE="PFC" WHERE ANL_FONCTION="PFC"');
  end;
//PAIE
AglNettoieListes('PGMULSALARIE', 'PSA_ETABLISSEMENT;PSA_SALARIE;PSA_CONFIDENTIEL;PSA_LIBELLE;PSA_PRENOM;PSA_DATEENTREE;PSA_DATESORTIE',Nil);
ExecuteSql('UPDATE SALARIES SET PSA_CATBILAN="000" WHERE (PSA_CATBILAN="" OR PSA_CATBILAN IS NULL)');
v_pgi.enableDEShare := True;
V_PGI.StandardSurDP := True;
ExecuteSQL ('UPDATE HISTOBULLETIN set phb_ordreetat=(Select prm_ordreetat from remuneration where ##Prm_predefini## prm_rubrique= substring(phb_rubrique,1,4)) where phb_naturerub="AAA"  and (phb_rubrique like "%.%" OR (PHB_ORDREETAT = 0 or PHB_ORDREETAT is null)) ') ;
// Dumet 08/12/2005 ExecuteSQL ('UPDATE HISTOBULLETIN set phb_ordreetat=(Select pct_ordreetat from cotisation where ##PCT_predefini## pct_rubrique=substring(phb_rubrique,1,4)) '
ExecuteSQL ('UPDATE HISTOBULLETIN set phb_omtsalarial=(Select pct_ordreetat from cotisation where ##PCT_predefini## pct_rubrique=substring(phb_rubrique,1,4)) '
          +' where phb_naturerub<>"AAA" and substring(phb_rubrique,1,4) in (select pct_rubrique from cotisation cot where ##cot.pct_predefini## cot.pct_themecot="RDC")') ;
ExecuteSQL ('UPDATE HISTOANALPAIE set pha_ordreetat=(Select prm_ordreetat from remuneration where ##Prm_predefini## prm_rubrique= pha_rubrique) where pha_naturerub="AAA"  and (PHA_ORDREETAT = 0 or PHa_ORDREETAT is null)') ;
ExecuteSQL ('UPDATE HISTOBULLETIN SET PHB_OMTSALARIAL = PHB_ORDREETAT WHERE (PHB_OMTSALARIAL IS NULL OR PHB_OMTSALARIAL=0)') ;
v_pgi.enableDEShare := false;
V_PGI.StandardSurDP := False;
//CPMPTA
ExecuteSql('UPDATE STDCPTA SET STC_DATEMODIF="' + UsdateTime(iDate1900) + '" WHERE STC_DATEMODIF IS NULL')
End;



Procedure MajVer714 ;
var S : String;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSQL('update parpiece set gpp_qualifmvt="ERC" where gpp_naturepieceg="AVS"');
    ExecuteSQL('update parpiece set gpp_qualifmvt="ERC" where gpp_naturepieceg="BRC"');
    ExecuteSQL('update parpiece set gpp_qualifmvt="" where gpp_naturepieceg="FRA"');
    ExecuteSQL('update parpiece set gpp_qualifmvt="" where gpp_naturepieceg="LCE"');
    ExecuteSQL('update parpiece set gpp_qualifmvt="SRC" where gpp_naturepieceg="BFA"');
    //
    //ExecuteSQL('UPDATE DECHAMPS SET DH_CONTROLE=DH_CONTROLE||"C" WHERE DH_PREFIXE = "GA" AND DH_CONTROLE LIKE "LD%"');
    //ExecuteSQL('UPDATE DECHAMPS SET DH_CONTROLE=DH_CONTROLE||"C" WHERE DH_PREFIXE = "ARS" AND DH_CONTROLE LIKE "LD%"');
    //ExecuteSQL('UPDATE DECHAMPS SET DH_CONTROLE=DH_CONTROLE||"C" WHERE DH_PREFIXE = "GL" AND DH_CONTROLE LIKE "LD%"');
  end;
//Compta
S := 'E_AUXILIAIRE;E_AFFAIRE;T_LIBELLE;T_SCORERELANCE;T_RELANCETRAITE;T_RELANCEREGLEMENT;E_NUMEROPIECE;E_MODEPAIE;E_DATEECHEANCE;E_DATERELANCE;E_QUALIFORIGINE;E_COUVERTUREDEV;E_DEBITDEV;E_CREDITDEV;';
S := S + 'E_DEBIT;E_CREDIT;E_COUVERTURE;E_DATECOMPTABLE;E_JOURNAL;E_EXERCICE;E_QUALIFPIECE;E_NATUREPIECE;';
S := S + 'E_NUMLIGNE;E_NUMECHE;E_DATEMODIF;E_ETABLISSEMENT;E_DEVISE;E_DATETAUXDEV;E_TAUXDEV;E_VALIDE;E_MODESAISIE;E_NIVEAURELANCE;E_BLOCNOTE;';

AglNettoieListes('CPRELANCECLIENT', S);

ExecuteSQL('update menu set mn_versiondev="-" where mn_tag=26095');    // activation lien OLE forcé à autorisé 28102005


  if not (V_PGI.Driver in [dbORACLE7, dbORACLE8 ,dbORACLE9]) then
  begin
  DupliqueFiltre ('SUPPRECR','CPSUPECRN');
  DupliqueFiltre ('SUPPRECR','CPSUPECRS');
  DupliqueFiltre ('SUPPRECR','CPSUPECRR');
  DupliqueFiltre ('MULMVT4H','CPVISECRH');
  DupliqueFiltre ('MULMVT4N','CPVISECRN');
  DupliqueFiltre ('MULMVT4R','CPVISECRR');
  DupliqueFiltre ('MULMVT4S','CPVISECRS');
  DupliqueFiltre ('MULMVT3H','CPMODECRH');
  DupliqueFiltre ('MULMVT3N','CPMODECRN');
  DupliqueFiltre ('MULMVT3R','CPMODECRR');
  DupliqueFiltre ('MULMVT3S','CPMODECRS');
  end;

End;

Procedure MajVer720 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
  //GIGA
  Executesql ('delete from choixcod where cc_type="AFP"');
  AglNettoieListes('AFLIGPLANNING', 'APL_NUMPIECE',nil);
  //GPAO
  ExecuteSQL('UPDATE QMATIERE SET QMT_FAMILLENIV1="",QMT_FAMILLENIV2="",QMT_FAMILLENIV3="" WHERE QMT_FAMILLENIV1 IS NULL');
  ExecuteSQL('UPDATE QARTTECH SET QAR_FAMILLENIV1="",QAR_FAMILLENIV2="",QAR_FAMILLENIV3="",QAR_FAMILLEMAT="" WHERE QAR_FAMILLENIV1 IS NULL');
  ExecuteSQL('UPDATE QMAGAPP  SET QMP_MATIERE="", QMP_FAMILLENIV1="",QMP_FAMILLENIV2="",QMP_FAMILLENIV3="" WHERE QMP_MATIERE IS NULL');
  ExecuteSQL('UPDATE QSTKCOLT SET QSK_STOCKRECOMPL=0,QSK_UNITEMIN="",QSK_UNITEALERTE="",QSK_UNITEMAX="",QSK_UNITERECOMPL="",QSK_DELAIMOYAPPRO=0,QSK_TAILLELOT=0,QSK_QTEMINFAB=0 WHERE QSK_STOCKRECOMPL IS NULL');
  ExecuteSQL('UPDATE QAPPROS SET QA_REFCOMMANDE="" WHERE QA_REFCOMMANDE IS NULL');
  ExecuteSQL('UPDATE QSTKDISPO SET QSD_REFCOMMANDE="" WHERE QSD_REFCOMMANDE IS NULL');
  ExecuteSQL('UPDATE WORDRELIG SET WOL_LIGNEORDREGEN=0 WHERE WOL_LIGNEORDREGEN IS NULL');
  end;


End;

Procedure MajVer721 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
  //GRC
  AglNettoieListes('RTMULACTIONSRAP','RAC_NUMCHAINAGE',nil);
  end;
End;


Procedure MajVer722 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GPAO
    GPAOMajMES;
    ExecuteSql('UPDATE WORDREBES SET WOB_SURCONSO="X",WOB_CONSOTOT="-" WHERE WOB_SURCONSO IS NULL');

  end;
End;

Procedure MajVer723 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //Maj_GSM_GUIDORI;
    ExecuteSQL('UPDATE LIGNECOMPL SET GLC_NBINTERVENANT=0 WHERE GLC_NBINTERVENANT IS NULL');
  end;
End;

Procedure MajVer724 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //SAV
    ExecuteSQL('UPDATE WPARCVERSION SET WPV_ARTICLEPARC=(SELECT WPN_ARTICLEPARC FROM WPARCNOME WHERE WPN_IDENTIFIANT=WPV_IDENTIFIANT)');
    //GPAO
    ExecuteSQL('UPDATE WNATURETRAVAIL SET WNA_IMPMODELEWOL="OR1", WNA_APERCUWOL="X" where WNA_IMPMODELEWOL IS NULL');
    { Mise à jour des familles de messages EDI }
    { Bons de réception }
    if not ExisteSQL('SELECT 1 FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="BLF"') then
    ExecuteSql('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE       , EFM_NATURESPIECE, EFM_SERIALISE)'
            +                  ' VALUES ("BLF"          , "Bon de réception", "PRF;BLF"       , "X"          )');

    { Commandes d'achat }
    if not ExisteSQL('SELECT 1 FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="CDA"') then
    ExecuteSql('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE     , EFM_NATURESPIECE, EFM_SERIALISE)'
           +                   ' VALUES ("CDA"          , "Commande achat", "CF"            , "X"          )');

    ExecuteSQL('update TIERSCOMPL  set YTC_MODELEBON="" where YTC_MODELEBON is null');
    ExecuteSQL('update TIERSCOMPL  set YTC_EDITRA="" where YTC_EDITRA is null');
    ExecuteSQL('update TIERSCOMPL  set YTC_MODELETXT="" where YTC_MODELETXT is null');
    ExecuteSQL('update TIERSCOMPL  set YTC_TIERSEXPE="" where YTC_TIERSEXPE is null');
    ExecuteSQL('update TIERSCOMPL  set YTC_QUALIFPOIDS="" where YTC_QUALIFPOIDS is null');
  end;
End;


Procedure MajVer725 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSQL('UPDATE PAYS SET PY_LIBELLE = "HAITI", PY_ABREGE = "HAITI" WHERE PY_PAYS = "HAI"');
    ExecuteSQL('UPDATE PAYS SET PY_LIBELLE = "GUADELOUPE", PY_ABREGE = "GUADELOUPE" WHERE PY_PAYS = "GOU"');
    //GP
    ExecuteSQL('UPDATE TIERSCOMPL SET YTC_DEPOT="", YTC_ETABLISSEMENT="" WHERE YTC_DEPOT IS NULL');
  end;
End;


Procedure MajVer726 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    // GCAIX
    ExecuteSQL('update TIERSCOMPL  set YTC_MODELEBON="",YTC_EDITRA="-",YTC_MODELETXT="",YTC_TIERSEXPE="",YTC_QUALIFPOIDS="" where YTC_MODELEBON is null and  YTC_EDITRA is null and YTC_MODELETXT is null and YTC_TIERSEXPE is null and YTC_QUALIFPOIDS is null');
    //GIGA
    ExecuteSql ('UPDATE AFINFOGRCDP set ADP_TYPEINFO="DP"');
  end;
End;

Procedure MajVer727 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSql('INSERT INTO WFORMULECHAMPDEF '
      + ' (WDY_CONTEXTE, WDY_NOMCHAMP, WDY_TYPECHAMP, WDY_DATECREATION, WDY_DATEMODIF, WDY_CREATEUR, WDY_UTILISATEUR) '
      + ' VALUES ("WOR","WOR_COEFPROD","COE","'+ USTime(Nowh)+'","'+UsTime(Nowh)+'","'+V_PGI.User+'","'+V_PGI.User+'")' );
    ExecuteSql('INSERT INTO WFORMULECHAMPDEF (WDY_CONTEXTE, WDY_NOMCHAMP, WDY_TYPECHAMP, '
      + ' WDY_DATECREATION, WDY_DATEMODIF, WDY_CREATEUR, WDY_UTILISATEUR) '
      + ' VALUES ("WOR","WOR_QPRODSAIS","QTE","'+ USTime(Nowh)+'","'+UsTime(Nowh)+'","'+V_PGI.User+'","'+V_PGI.User+'")' );
    ExecuteSql('INSERT INTO WFORMULECHAMPDEF '
      + ' (WDY_CONTEXTE, WDY_NOMCHAMP, WDY_TYPECHAMP, WDY_DATECREATION, WDY_DATEMODIF, WDY_CREATEUR, WDY_UTILISATEUR) '
      + ' Values ("WOR","WOR_QPRODSTOC","QTE","'+ USTime(Nowh)+'","'+UsTime(Nowh)+'","'+V_PGI.User+'","'+V_PGI.User+'")' );
    ExecuteSql('INSERT INTO WFORMULECHAMPDEF '
      + ' (WDY_CONTEXTE, WDY_NOMCHAMP, WDY_TYPECHAMP, WDY_DATECREATION, WDY_DATEMODIF, WDY_CREATEUR, WDY_UTILISATEUR) '
      + ' VALUES ("WOR","WOR_QUALIFUNITESTO","UNI","'+ USTime(Nowh)+'","'+UsTime(Nowh)+'","'+V_PGI.User+'","'+V_PGI.User+'")' );
    ExecuteSql('INSERT INTO WFORMULECHAMPDEF '
      + ' (WDY_CONTEXTE, WDY_NOMCHAMP, WDY_TYPECHAMP, WDY_DATECREATION, WDY_DATEMODIF, WDY_CREATEUR, WDY_UTILISATEUR) '
      + ' VALUES ("WOR","WOR_UNITEPROD","UNI","'+ USTime(Nowh)+'","'+UsTime(Nowh)+'","'+V_PGI.User+'","'+V_PGI.User+'")' );
    ExecuteSql('INSERT INTO WFORMULECHAMP '
      + ' (WDX_CONTEXTE, WDX_QUALIFIANT, WDX_CHAMPSAIS, WDX_UNITECHAMPSAIS,'
      + 'WDX_CHAMPCOEF, WDX_CHAMPCALC, WDX_UNITECHAMPCALC, WDX_DATECREATION, WDX_DATEMODIF, WDX_CREATEUR, WDX_UTILISATEUR) '
      + ' VALUES ("WOR","QPR","WOR_QPRODSTOC","WOR_QUALIFUNITESTO","WOR_COEFPROD",'
      + '"WOR_QPRODSAIS","WOR_UNITEPROD","'
      + USTime(Nowh)+'","'+UsTime(Nowh)+'","'+V_PGI.User+'","'+V_PGI.User+'")' );
    ExecuteSql('UPDATE WORDRERES'
      + ' SET WOR_RESORIGINEMES="", WOR_NBRESSOURCE=1,WOR_UNITEPROD= (SELECT WOG_UNITEACC FROM WORDREGAMME'
      + ' WHERE WOG_NATURETRAVAIL=WOR_NATURETRAVAIL AND WOG_LIGNEORDRE=WOR_LIGNEORDRE'
      + ' AND WOG_OPECIRC=WOR_OPECIRC AND WOG_NUMOPERGAMME=WOR_NUMOPERGAMME),'
      + ' WOR_QUALIFUNITESTO= (SELECT WOG_QUALIFUNITESTO FROM WORDREGAMME'
      + ' WHERE WOG_NATURETRAVAIL=WOR_NATURETRAVAIL AND WOG_LIGNEORDRE=WOR_LIGNEORDRE'
      + ' AND WOG_OPECIRC=WOR_OPECIRC AND WOG_NUMOPERGAMME=WOR_NUMOPERGAMME),WOR_COEFPROD=1,WOR_QPRODSAIS=0,WOR_QPRODSTOC=0');

    ExecuteSql ('UPDATE ARTICLE SET GA_GMARQUE="-", GA_GCHOIXQUALITE="-" WHERE GA_GMARQUE IS NULL');

    ExecuteSql('UPDATE CATALOGU SET GCA_LIBREGCA1="",GCA_LIBREGCA2="",GCA_LIBREGCA3="",GCA_LIBREGCA4="",GCA_LIBREGCA5="",GCA_LIBREGCA6="",GCA_LIBREGCA7="",GCA_LIBREGCA8="",GCA_LIBREGCA9="",GCA_LIBREGCAA="",'
                             + 'GCA_CHARLIBRE1="",GCA_CHARLIBRE2="",GCA_CHARLIBRE3="",'
                             + 'GCA_DATELIBRE1="' + UsDateTime(iDate1900) + '",GCA_DATELIBRE2="' + UsDateTime(iDate1900) + '",GCA_DATELIBRE3="' + UsDateTime(iDate1900) + '",'
                             + 'GCA_VALLIBRE1=0,GCA_VALLIBRE2=0,GCA_VALLIBRE3=0,GCA_BOOLLIBRE1="-",GCA_BOOLLIBRE2="-",GCA_BOOLLIBRE3="-" WHERE GCA_LIBREGCA1 IS NULL');


    ExecuteSql('UPDATE LIGNECOMPL SET GLC_MARQUE="", GLC_CHOIXQUALITE="" WHERE GLC_MARQUE IS NULL');

    MajMarkMESApres;
    //TRESO
    AglNettoieListes('TRLISTEVIR', 'TEQ_PAYSDEST;TEQ_PAYSSOURCE;TEQ_FICVIR;TEQ_DEVISE;', nil);
    //GRC
    AglNettoieListes('WARTICLEPARC_GENE','WAP_DUREEGARANTIE',nil);
    AglNettoieListes('WPARCSELECT','WPC_IDENTIFIANT;WPC_TIERS',nil);
    //GPAO
    ExecuteSQL('UPDATE PARPIECE SET GPP_NATURESUIVANTE="BSA;", GPP_CONTEXTES="GC;"  WHERE GPP_NATUREPIECEG="CSA"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_CONTEXTES="GC;"  WHERE GPP_NATUREPIECEG="BSA"');

    ExecuteSql('UPDATE STKFICHETRACE SET GST_MARQUE = "" WHERE GST_MARQUE IS NULL');
    ExecuteSql('UPDATE STKFICHETRACE SET GST_CHOIXQUALITE = "" WHERE GST_CHOIXQUALITE IS NULL');

    ExecuteSql('UPDATE LIGNECOMPL SET GLC_MARQUE = "" WHERE GLC_MARQUE IS NULL');
    ExecuteSql('UPDATE LIGNECOMPL SET GLC_CHOIXQUALITE = "" WHERE GLC_CHOIXQUALITE IS NULL');

    if not existeSql('SELECT 1 FROM STKNATURE WHERE GSN_QUALIFMVT="CMQ"') then
    begin
        ExecuteSql('INSERT INTO STKNATURE (GSN_QUALIFMVT, GSN_LIBELLE, GSN_STKTYPEMVT, GSN_QTEPLUS, GSN_QUALIFMVTSUIV, GSN_SIGNEMVT,'
                 + ' GSN_STKFLUX, GSN_GERECOMPTA, GSN_CALLGSL, GSN_CALLGSS, GSN_CONTREMARQUE, GSN_MAJPRIXVALO, GSN_SDISPODISPATCH, GSN_SFLUXDISPATCH, GSN_SDISPOPICKING, GSN_SFLUXPICKING)'
                 + ' VALUES ("CMQ", "Changement de marque", "CET", "", "", "MIX",'
                 + ' "STO", "-", "-", "-", "-", "", "", "", "LBR;", "STD;")');
    end;

    if not existeSql('SELECT 1 FROM STKNATURE WHERE GSN_QUALIFMVT="CCQ"') then
    begin
        ExecuteSql('INSERT INTO STKNATURE (GSN_QUALIFMVT, GSN_LIBELLE, GSN_STKTYPEMVT, GSN_QTEPLUS, GSN_QUALIFMVTSUIV, GSN_SIGNEMVT,'
                 + ' GSN_STKFLUX, GSN_GERECOMPTA, GSN_CALLGSL, GSN_CALLGSS, GSN_CONTREMARQUE, GSN_MAJPRIXVALO, GSN_SDISPODISPATCH, GSN_SFLUXDISPATCH, GSN_SDISPOPICKING, GSN_SFLUXPICKING)'
                 + ' VALUES ("CCQ", "Changement de choix de qualité", "CET", "", "", "MIX",'
                 + ' "STO", "-", "-", "-", "-", "", "", "", "LBR;", "STD;")');
    end;


    AGLNettoieListes('UORDRELIG','WOL_ORDREPERE',nil);
    AglNettoieListes('GCSTKDISPATCH'  , 'GQD_MARQUE;GQD_CHOIXQUALITE',nil);
    AglNettoieListes('GCSTKPICKING'   , 'GQD_MARQUE;GQD_CHOIXQUALITE',nil);
    AglNettoieListes('GCLISTEINVLIGNE', 'GIL_MARQUE;GIL_CHOIXQUALITE',nil);
     //giga

    if vers_afftiers<108 then ExecuteSql ('update afftiers set aft_affairesession="",aft_inscrit="-"');
    AglNettoieListes('AFLIGPLANNING', 'APL_NUMEROTACHE;APL_QTEPLANIFIEE;APL_NUMPIECE;APL_HEUREDEBPLA;APL_HEUREFINPLA',nil);
    AglNettoieListes('AFTACHE_MUL', 'ATA_TYPEPLANIF',nil);
  end;
End;


Procedure MajVer728 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //  ExecuteSQL('UPDATE LIGNECOMPL SET GLC_DATEDEB="'+UsDateTime(iDate1900)+'",GLC_DATEFIN="'+UsDateTime(iDate1900)+'",GLC_NUMEROMARCHE=0,GLC_CODEMARCHE="" WHERE GLC_CODEMARCHE IS NULL');
    AglNettoieListes('GCLISTINVLIGNE'  , 'GM_LIBLLE;GCQ_LIBELLE;GIL_REFAFFECATION',nil);
    //GIGA
    AglNettoieListes('AFPLALIGNE_MUL', 'GLC_AFFAIRELIE',nil);
  end;

ExecuteSQL('UPDATE UTILISAT SET US_EMAILSMTPLOGIN="", US_EMAILSMTPPWD ="", US_POPPORT="", US_SMTPPORT=""');
//ajout 01/02/2006
ExecuteSQL('UPDATE IMMO SET I_OPEDEPREC="-",I_REPRISEDEP=0.00,I_ECCLEECR="",I_DOCGUID="",I_DATEFINCB="'+UsDateTime(iDate1900)+'"');

End;


Procedure MajVer729 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    InsertChoixCode('CVU', '003', 'Conversion d''unité', 'AdvConvU', '');
    ExecuteSQL('UPDATE ARTICLE SET GA_UNITEPRIXACH=ISNULL(GA_QUALIFUNITEVTE,""),GA_COEFCONVQTEACH=1,GA_COEFCONVQTEVTE=1,GA_COEFCONVCONSO=1 WHERE GA_COEFCONVQTEACH IS NULL');
    ExecuteSQL('UPDATE CATALOGU SET GCA_UNITEQTEACH=GCA_QUALIFUNITEACH,GCA_COEFCONVQTEACH=1 WHERE GCA_UNITEQTEACH IS NULL');
    UpdateDecoupeLigne('GL_UNITEPRIX=GL_QUALIFQTESTO', 'AND (ISNULL(GL_ARTICLE, " ") <> " ")');
    ExecuteSQL('UPDATE WCBNEVOLUTION SET WEV_MARQUE="",WEV_CODEMARCHE="",WEV_NUMEROMARCHE=0,WEV_QUOTAPCTDEM=0,WEV_QUOTAPCTAPP=0 WHERE WEV_MARQUE IS NULL');
    ExecuteSQL('UPDATE PARPIECE SET GPP_TARIFGENSAISIE="010", GPP_TARIFGENTRANSF="010", GPP_TARIFGENDATE="010", GPP_TARIFGENDEPOT="-", GPP_TARIFGENSPECIA="010" WHERE GPP_TARIFGENSAISIE IS NULL');
    ExecuteSQL('UPDATE PARPIECE SET GPP_TARIFGENSAISIE="020", GPP_TARIFGENTRANSF="010", GPP_TARIFGENDATE="020", GPP_TARIFGENDEPOT="-", GPP_TARIFGENSPECIA="040" WHERE GPP_NATUREPIECEG="CMF"');
    ExecuteSQL('UPDATE YTARIFSPARAMETRES SET YFO_OKTARIFMARCHE = "----" WHERE YFO_OKTARIFMARCHE IS NULL');
    ExecuteSQL('UPDATE YTARIFS SET YTS_TARIFMARCHE="" WHERE YTS_TARIFMARCHE IS NULL');
    ExecuteSQL('UPDATE YTARIFS SET YTS_FORMULEMT="" WHERE YTS_FORMULEMT IS NULL');
    ExecuteSQL('UPDATE WPARAMFONCTION SET WPF_FORMULEMT="" WHERE  WPF_FORMULEMT IS NULL');
    { Nouvelle nature de stock permettant de gérer les unités logistiques }
    if not ExisteSQL('SELECT 1 FROM STKNATURE WHERE GSN_QUALIFMVT="SCO"') then
    begin
      ExecuteSQL('INSERT INTO STKNATURE (GSN_CALLGSL, GSN_CALLGSS, GSN_CONTREMARQUE, GSN_GERECOMPTA, GSN_LIBELLE, '
                +                       'GSN_MAJPRIXVALO, GSN_QTEPLUS, GSN_QUALIFMVT, GSN_QUALIFMVTSUIV, '
                +                       'GSN_SDISPODISPATCH, GSN_SDISPOPICKING, GSN_SFLUXDISPATCH, GSN_SFLUXPICKING, '
                +                       'GSN_SIGNEMVT, GSN_STKFLUX, GSN_STKTYPEMVT)'
                + ' VALUES ("-", "-", "-", "-", "Sortie de colisage", "", "GQ_PHYSIQUE;", "SCO", "", "", "", "", "", "MIX", "STO", "PHY")');
    end;
    //GIGA
    ExecuteSql ('UPDATE AFREVISION set AFR_LASTPRIXCALC=0,AFR_LASTPRIXCALCD=0');
    ExecuteSql ('Update AFFAIRE set aff_revlastprix="-"');
    if vers_tache < 124 then
      ExecuteSql ('Update Tache set ata_typeadresse="",ata_numeroadresse=0');
    if Vers_afplanning < 113 then
      ExecuteSql ('Update Afplanning set apl_typeadresse="",apl_numeroadresse=0');
    //GC
    AglNettoieListes('GCLISTINVENT', 'GIE_CONTREMARQUE', nil);
    ExecuteSQL('UPDATE LIGNECOMPL SET GLC_MODIFPRIXACHAT="-" WHERE GLC_MODIFPRIXACHAT IS NULL');
  end;
End;

Procedure MajVer730 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GPAO
    ExecuteSQL('UPDATE PARPIECE SET GPP_NATURESUIVANTE="" WHERE GPP_NATUREPIECEG="CMF"');
    ExecuteSQL('UPDATE WORDRELAS SET WLS_PHASE="" WHERE WLS_PHASE IS NULL');
    ExecuteSQL('UPDATE QCIRCUIT SET QCI_COMAGENT=0 WHERE QCI_COMAGENT IS NULL');
    ExecuteSQL('UPDATE STKVALOTYPE SET GVT_VALOPRIXCOMP = "" WHERE  GVT_VALOPRIXCOMP IS NULL');
    ExecuteSQL('UPDATE CATALOGU SET GCA_QECOACH=0,GCA_QPCBACH=0 WHERE GCA_QECOACH IS NULL');
    //AIX
    { Init Nouvelles Zones }
    ExecuteSQL('update TIERSCOMPL  set YTC_MODELE="",YTC_AUTOEDITETIQ="-"'
             +',YTC_AUTOEDITOT ="-",YTC_TYPEECHANGE="",YTC_PATHFICPRIVE=""'
             +',YTC_AUTOGENEPRIVE="-",YTC_SOUCHEOT="",YTC_CODEPRODTRA=""'
             +',YTC_QUALIFLINEAIRE="",YTC_QUALIFVOLUME="" '
             +' where YTC_MODELE is null and YTC_AUTOEDITETIQ is null '
             +'       and YTC_AUTOEDITOT is null and YTC_TYPEECHANGE is null '
             +'       and YTC_PATHFICPRIVE is null and YTC_AUTOGENEPRIVE is null '
             +'       and YTC_SOUCHEOT is null and YTC_CODEPRODTRA is null '
             +'       and YTC_QUALIFLINEAIRE is Null and YTC_QUALIFVOLUME is Null');
    {Maj YTC_TYPEECHANGE Si YTC_EDITRA est True}
    ExecuteSQL('update TIERSCOMPL  set YTC_TYPEECHANGE="PRI",YTC_AUTOGENEPRIVE="X" where YTC_EDITRA = "X" ');
    { Init Nouvelle Zone Régime transport }
    ExecuteSQL('update gcproduittra set gtc_regimetra="" ');
    //DP
    if IsMonoOuCommune and (Not ExisteSQL('SELECT 1 FROM DOSSIERGRP')) then
        ExecuteSQL ('INSERT INTO DOSSIERGRP (DOG_NODOSSIER,DOG_GROUPECONF) SELECT DOS_NODOSSIER AS DOG_NODOSSIER,DOS_GROUPECONF AS DOG_GROUPECONF FROM DOSSIER')
  end;

ExecuteSQL ('UPDATE SECTION SET S_LIBELLE = "FRANCE METRO. - TAUX 19,6 %" WHERE S_SECTION = "0206" AND S_CREERPAR ="TVA"');
if (GetParamSocSecur('SO_CPPCLSAISIETVA', '')) and (GetParamSocSecur('SO_CPPCLAXETVA', '') = '') then
    begin
          SetParamSoc('SO_CPPCLAXETVA', GetColonneSQL('AXE', 'X_AXE', 'X_LIBELLE = "TVA"'));
          InsertChoixCode('ATU', '038', 'des tiers', '', 'des clients');
    end;
ExecuteSQL('update dechamps set dh_controle="L" where dh_nomchamp="GU_REFRELEVE" or dh_nomchamp="GU_TRESORERIE"');

End;

Procedure MajVer734 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    // MES
    if not ExisteSQL('SELECT 1 from QOFAMILLEALEA where QOF_FAMILLEALEA="OP"') then
      ExecuteSql('INSERT INTO QOFAMILLEALEA '
            + ' (QOF_FAMILLEALEA, QOF_LIBELLE, QOF_ALEADEDUCTIBLE, QOF_ALEAPRIME, QOF_ALEARATIO) '
            + ' VALUES ("OP","Opérations","X","X","-")' );
    if not ExisteSQL('SELECT 1 from QOALEA where QOA_ALEA="ALEAOP"') then
      ExecuteSql('INSERT INTO QOALEA '
            + ' (QOA_ALEA,QOA_FAMILLEALEA,QOA_LIBELLE,QOA_ACTIF,QOA_ALEADEDUCTIBLE,QOA_DATEDEBCONST,QOA_DATEFINCONST, '
            +   'QOA_DATECREATION,QOA_DATEMODIF,QOA_CREATEUR,QOA_UTILISATEUR) '
            + ' VALUES ("ALEAOP","OP","Aléas sur Opérations","X","X","'+UsTime(iDate1900)+'","'+UsTime(iDate1900)+'",'
            + ' "'+UsTime(Nowh)+'","'+UsTime(Nowh)+'","'+V_PGI.User+'","'+V_PGI.User+'")' );
    ExecuteSql('Update QWHISTORES SET QWH_ORDREWOR=0');
    //GPAO
    ExecuteSql('UPDATE WORDRELIG SET WOL_DATEEDITFIC = "' + UsDateTime(iDate1900) + '" WHERE WOL_DATEEDITFIC IS NULL');
    ExecuteSql('UPDATE WORDRELIG SET WOL_DATEEDITBON = "' + UsDateTime(iDate1900) + '" WHERE WOL_DATEEDITBON IS NULL');
    ExecuteSql('UPDATE WORDRELIG SET WOL_USEREDITFIC = "" WHERE WOL_USEREDITFIC IS NULL');
    ExecuteSql('UPDATE WORDRELIG SET WOL_USEREDITBON = "" WHERE WOL_USEREDITBON IS NULL');
    //DP
    AGLNettoieListes ('DPDOCUMENTS_MUL','DPD_UTILISATEUR',nil);
    //GEP
  end;

ExecuteSql('update scriptsp set YSC_QUALIFIANT="SG" where ysc_script not in (select YFM_FORME from formesp) and ysc_typescript="YY"');

End;

Procedure MajVer735 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GPAO
    ExecuteSQL('UPDATE LIGNEFRAIS SET LF_POURCENT1=LF_POURCENT, LF_POURCENT2=0, LF_POURCENT3=0 WHERE LF_POURCENT1 IS NULL');
    //GIGA
    AglNettoieListes('AFCTRL_ANNTIERS', 'ANN_GUIDPER;ANN_NOMPER',nil);
    AglNettoieListes('AFDOUBLONS_ANNTIE', 'ANN_GUIDPER;ANN_NOMPER',nil);
    AglNettoieListes('AFDOUBLONS_ANNU', 'ANN_GUIDPER',nil);
    // fin giga
  end;
  ExecuteSQL ('update journal set j_ouvrirlett="-"');
  ExecuteSQL ('update jalref set jr_ouvrirlett="-"');
End;

Procedure MajVer736 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    // Mise à jour des pièces jointes
    ExecuteSQL('update YMSGFILES set YMG_ATTACHED="X"');
    // Transfert du contenu Html des mails reçus
    ExecuteSQL('insert into YMSGFILES(YMG_MSGGUID, YMG_MSGID, YMG_FILEGUID, YMG_FILEID, YMG_ATTACHED)'
     + ' select YMS_MSGGUID, YMS_MSGID, YMF_FILEGUID, YMF_FILEID, "-"'
     + ' from YMAILFILES, YMESSAGES'
     + ' where YMS_USERMAIL=YMF_UTILISATEUR AND YMS_MAILID=YMF_MAILID and YMF_ATTACHED="-"');

    //GPAO
    ExecuteSQL('UPDATE PARPIECE SET GPP_MAJPRIXVALO="" WHERE GPP_NATUREPIECEG IN ("CSP","BSP") AND GPP_MAJPRIXVALO LIKE "DPA;DPR;PPA;PPR;"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_MAJPRIXVALO="DPA;DPR;" WHERE GPP_NATUREPIECEG = "CSA" AND GPP_MAJPRIXVALO LIKE "DPA;DPR;PPA;PPR;"');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="ASA" WHERE GPP_NATUREPIECEG IN ("CSA")');
    ExecuteSQL('UPDATE PARPIECE SET GPP_STKQUALIFMVT="ESA" WHERE GPP_NATUREPIECEG IN ("BSA")');
    //
    AGLNettoieListes('GCCATALOGUE','GCA_TIERS',nil);

    ExecuteSQL('UPDATE ETABLISS SET ET_EAN="" WHERE ET_EAN IS NULL');

    if not ExisteSQL('SELECT 1 FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="DEL"') then
      ExecuteSQL('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE, EFM_NATURESPIECE, EFM_SERIALISE) VALUES ("DEL", "Plan d''approvisionnement DELFOR", "", "X")');

    if not ExisteSQL('SELECT 1 FROM STKVALOTYPE where GVT_VALOTYPE="SAA"') then
    begin
      ExecuteSQL('INSERT INTO STKVALOTYPE (GVT_VALOTYPE, GVT_LIBELLE, GVT_VALOPRIX, GVT_VALOPRIXCOMP'
                +', GVT_TYPEPDR, GVT_TEMPSREEL, GVT_RECALCUL, GVT_ECARTVALO, GVT_AVECINDIRECT, GVT_CREATEUR'
                +', GVT_DATECREATION, GVT_UTILISATEUR, GVT_DATEMODIF)'
                +' VALUES ("SAA", "Valorisation au dernier prix de sous traitance d''achat", "04", "41", "ACT", "-", "-", "-", "-", "CEG", "", "000", "")');
    end;
    if not ExisteSQL('SELECT 1 FROM STKNATURE WHERE GSN_QUALIFMVT="ASA"') then
    begin
      ExecuteSQL('INSERT INTO STKNATURE (GSN_CALLGSL, GSN_CALLGSS, GSN_CONTREMARQUE, GSN_GERECOMPTA, GSN_LIBELLE, '
                +                       'GSN_MAJPRIXVALO, GSN_QTEPLUS, GSN_QUALIFMVT, GSN_QUALIFMVTSUIV, '
                +                       'GSN_SDISPODISPATCH, GSN_SDISPOPICKING, GSN_SFLUXDISPATCH, GSN_SFLUXPICKING, '
                +                       'GSN_SIGNEMVT, GSN_STKFLUX, GSN_STKTYPEMVT)'
                + ' VALUES ("-", "-", "-", "-", "Attendu de sous traitance d''achat", "DPA;DPR,", "GQ_RESERVEFOU;", "ASA", "", "", "", "", "", "ENT", "ACH", "ATT")');
    end;

    if not ExisteSQL('SELECT 1 FROM STKNATURE WHERE GSN_QUALIFMVT="ESA"') then
    begin
      ExecuteSQL('INSERT INTO STKNATURE (GSN_CALLGSL, GSN_CALLGSS, GSN_CONTREMARQUE, GSN_GERECOMPTA, GSN_LIBELLE, '
                +                       'GSN_MAJPRIXVALO, GSN_QTEPLUS, GSN_QUALIFMVT, GSN_QUALIFMVTSUIV, '
                +                       'GSN_SDISPODISPATCH, GSN_SDISPOPICKING, GSN_SFLUXDISPATCH, GSN_SFLUXPICKING, '
                +                       'GSN_SIGNEMVT, GSN_STKFLUX, GSN_STKTYPEMVT)'
                + ' VALUES ("-", "-", "-", "-", "Entréee de sous traitance d''achat", "DPA;DPR;PPA;PPR;", "GQ_PHYSIQUE;", "ESA", "", "", "", "", "", "MIX", "ACH", "PHY")');
    end;


    ExecuteSQL('UPDATE STKMOUVEMENT SET GSM_QUALIFMVT="ASA" WHERE GSM_NATUREORI="CSA"');
    ExecuteSQL('UPDATE STKMOUVEMENT SET GSM_QUALIFMVT="ESA" WHERE GSM_NATUREORI="BSA"');

    if not ExisteSQL('SELECT 1 FROM STKVALOPARAM WHERE GVP_QUALIFMVT="ASA"') then
    begin
      ExecuteSQL('INSERT INTO STKVALOPARAM (GVP_IDENTIFIANT, GVP_QUALIFMVT, GVP_NATURETRAVAIL, GVP_STKFLUX, GVP_FAMILLENIV1, GVP_FAMILLENIV2, GVP_FAMILLENIV3, GVP_FAMILLEVALO, GVP_VALOTYPE, GVP_CREATEUR, GVP_DATECREATION, GVP_UTILISATEUR, GVP_DATEMODIF)'
                + ' VALUES ('+inttostr(wSetJeton('GVP'))+', "ASA", "", "", "", "", "", "", "SAA", "CEG", "", "000", "")');
    end;

    if not ExisteSQL('SELECT 1 FROM STKVALOPARAM WHERE GVP_QUALIFMVT="ESA"') then
    begin
      ExecuteSQL('INSERT INTO STKVALOPARAM (GVP_IDENTIFIANT, GVP_QUALIFMVT, GVP_NATURETRAVAIL, GVP_STKFLUX, GVP_FAMILLENIV1, GVP_FAMILLENIV2, GVP_FAMILLENIV3, GVP_FAMILLEVALO, GVP_VALOTYPE, GVP_CREATEUR, GVP_DATECREATION, GVP_UTILISATEUR, GVP_DATEMODIF)'
                + ' VALUES ('+inttostr(wSetJeton('GVP'))+', "ESA", "", "", "", "", "", "", "SAA", "CEG", "", "000", "")');
    end;

    ExecuteSQL('UPDATE WORDREBES SET WOB_DPA=0, WOB_DPR=0 WHERE WOB_DPA IS NULL');

    ExecuteSql('UPDATE STKFORMULE SET GSF_COMPOSANT="",GSF_LIBREART1="",GSF_LIBREART2="",GSF_LIBREART3="",GSF_LIBREART4="",GSF_LIBREART5="",GSF_LIBREART6="",GSF_LIBREART7="",GSF_LIBREART8="",GSF_LIBREART9="",GSF_LIBREARTA="" WHERE GSF_COMPOSANT IS NULL');

    ExecuteSql('UPDATE STKFICHETRACE SET GST_VALLIBRE1=0,GST_VALLIBRE2=0,GST_VALLIBRE3=0 WHERE GST_VALLIBRE1 IS NULL');

    AglNettoieListes('GCSTKDISPATCH'   , 'GQD_VALLIBRE1;GQD_VALLIBRE2,GQD_VALLIBRE3',nil);
    AglNettoieListes('GCSTKPICKING'    , 'GQD_VALLIBRE1;GQD_VALLIBRE2,GQD_VALLIBRE3',nil);
    AglNettoieListes('GCSTKDISPODETAIL', 'GQD_VALLIBRE1;GQD_VALLIBRE2,GQD_VALLIBRE3',nil);
    AglNettoieListes('GCLISTINVLIGNECTM', 'GIL_REFAFFECTATION',nil);
    //giga
    ExecuteSQL('Update CONTACT set C_GUIDPER='' where C_GUIDPER is null');
    If IsMonoOuCommune then SuppAnnuinterloc; //pasage table annuinterloc dans contact dans base commune
    //fin giga
    //DP
    If IsMonoOuCommune then
    begin
      ExecuteSQL('update YMESSAGES set YMS_STRDATE=""');
     //  ExecuteSQL('update YMESSAGES set YMS_STRDATE=YMA_STRDATE from YMAILS where YMS_USERMAIL=YMA_UTILISATEUR and YMS_MAILID=YMA_MAILID');
      ExecuteSQL('update YMESSAGES set YMS_STRDATE=(select YMA_STRDATE from YMAILS where YMS_USERMAIL=YMA_UTILISATEUR and YMS_MAILID=YMA_MAILID)');
    end;
    //GC
    ExecuteSql('UPDATE PARPIECE SET GPP_GEREARTICLELIE = "AUT" WHERE GPP_GEREARTICLELIE IS NULL');
    ExecuteSql('update articlelie set gal_articlelie=isnull((select ga_article from article where ga_codearticle=gal_articlelie and ga_statutart in ("UNI","GEN")),gal_articlelie) where len(gal_articlelie)<34');

    ExecuteSQL('UPDATE PIECE SET GP_OPERATION = ""');

    // transfert elements de GCZONELIBRE vers GCZONELIBRECON
    ExecuteSQL('UPDATE CHOIXCOD SET CC_TYPE="ZLC" WHERE CC_TYPE="ZLI" AND (CC_CODE LIKE "BT%" OR CC_CODE LIKE "BB%" or CC_CODE LIKE "BC%" or CC_CODE LIKE "BD%" or CC_CODE LIKE "BM%")');


  end;
End;

Procedure MajVer737 ;
var
  st: string ; // XP 08.03.2006 Oublie de PC
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GIGA
    AglNettoieListes('AFMUL_EACTIVITE1', 'EAC_UTILISATEUR',nil);
    AglNettoieListes('AFMUL_EACTIVITE2', 'EAC_UTILISATEUR',nil);
    // JURI
    ExecuteSQL('update JUSYSCLE set jsc_guidx = jsc_codex WHERE JSC_GUIDX = "" OR JSC_GUIDX IS NULL');
    //gp
    AglNettoieListes('GCSTKDISPODETAIL', 'GQD_DATEDISPO',nil);
    AglNettoieListes('GCSTKPICKING', 'GQD_DATEDISPO',nil);
    AglNettoieListes('GCSTKDISPATCH', 'GQD_DATEDISPO',nil);

    ExecuteSql('update stknature set gsn_ctrldispo="000" where gsn_ctrldispo is null');
    ExecuteSql('update stknature set gsn_ctrlperemption="000" where gsn_ctrlperemption is null');
  end;
  //PAIE
  ExecuteSql('UPDATE DECLARATIONS SET PDT_AUTREVICT="-"');
  ExecuteSql('UPDATE SALARIES SET PSA_ETATBULLETIN=(SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_PGBULDEFAUT")');
  ExecuteSql('UPDATE SALARIES SET PSA_NUMEROBL=""');
  ExecuteSql('UPDATE ORGANISMEPAIE SET  POG_CONDSPEC="-"');
  ExecuteSql('UPDATE DUCSPARAM SET PDP_CONDITION="" WHERE PDP_PREDEFINI <> "CEG"');
  ExecuteSql('UPDATE MAINTIEN SET PMT_HISTOMAINT ="-"  ');
  ExecuteSql('UPDATE MOTIFABSENCE SET PMA_CARENCEIJSS=0');
  ExecuteSql('UPDATE MOTIFABSENCE SET PMA_CARENCEIJSS=3 WHERE PMA_TYPEABS = "MAN"');
  ExecuteSql('UPDATE MOTIFABSENCE SET PMA_CARENCEIJSS=1 WHERE PMA_TYPEABS = "MAP" OR PMA_TYPEABS = "ATJ" OR PMA_TYPEABS = "ATR"');
  ExecuteSql('UPDATE MAINTIEN SET PMT_TYPEMAINTIEN="MAI" WHERE PMT_CARENCE <> 9999');
  ExecuteSql('UPDATE MAINTIEN SET PMT_TYPEMAINTIEN="GAR" WHERE PMT_CARENCE = 9999');
  ExecuteSql('UPDATE MAINTIEN SET PMT_HISTOMAINT ="-"');
  ExecuteSql('UPDATE CRITMAINTIEN SET PCM_RUBGARANTIE=""');
  ExecuteSql('UPDATE REMUNERATION SET PRM_EXCLMAINT="-", PRM_EXCLENVERS="-" WHERE PRM_PREDEFINI <> "CEG"');
  AglNettoieListes('PGMAINTIEN', 'PMT_HISTOMAINT', nil);
  AglNettoieListes('PGMULMVTIJSS', 'PSA_LIBELLE', nil);
  AglNettoieListes('PGDUCSINIT', '', nil);
  st := 'update absencesalarie set pcn_dateannulation="' + UsDateTime(Idate1900) +
    '",pcn_etatpostpaie="VAL",pcn_motifannul="",pcn_annulepar=""';
  ExecuteSql(st);
  st := 'update motifabsence set pma_annulable="-", pma_controlmotif="X",pma_createur="",pma_utilisateur="",' +
    'pma_pgcoloractif="10944422",pma_pgcolorpay="16053248",pma_pgcolorval="10944422",  pma_pgcoloratt="10210815"' +
    ', pma_pgcolorann="8421631",pma_pgcolorref="8421631",PMA_EDITPLANPAIE="X",PMA_EDITPLANABS="-",PMA_SSJOURFERIE="-"';
    ExecuteSql(St);
  ExecuteSql('update motifabsence set PMA_EDITPLANABS="X" WHERE PMA_MOTIFEAGL="X"');
  ExecuteSql('update RECAPSALARIES set PRS_ACQUISANC=0, PRS_ACQUISCPSUIV=0, PRS_ACQUISRTTSUIV=0');
  ExecuteSql('Update CONTRATTRAVAIL set PCI_FINTRAVAIL = PCI_FINCONTRAT');
  ExecuteSql('UPDATE ENVOISOCIAL SET PES_CTRLENVOI="X" WHERE PES_CTRLENVOI<>"X" AND PES_CTRLENVOI<>"-"');
  ExecuteSql('UPDATE DADSPERIODES SET PDE_EXERCICEDADS="2005", PDE_DECALEE="01"');
  ExecuteSql('UPDATE DADSDETAIL SET PDS_EXERCICEDADS="2005"');
  ExecuteSql('UPDATE DADS2SALARIES SET PD2_CHQEMPLOI="-", PD2_INDEMIMPATRI=0, PD2_AUTREREVENUS=0');
  ExecuteSql('UPDATE ATTESTATIONS SET PAS_RECLASS="-", PAS_INDCNEB="-", PAS_INDCNEM=0, PAS_INDSINISTREB="-", PAS_INDSINISTREM=0, PAS_INDSPCFIQB="-", PAS_INDSPCFIQM=0, PAS_INDCPEB="-", PAS_INDCPEM=0, PAS_INDAPPRENTIB="-", PAS_INDAPPRENTIM=0');
  st := 'UPDATE MAINTIEN SET PMT_HISTOMAINT ="X" WHERE (PMT_DATEDEBUT = "'
        + usdatetime(strtodate('01/01/1959')) + '" AND PMT_DATEFIN = "'
        + usdatetime(strtodate('01/01/1959')) + '") OR  (PMT_DATEDEBUT = "' + usdatetime(strtodate('31/12/1900')) + '" AND PMT_DATEFIN = "' + usdatetime(strtodate('31/12/1900')) + '")';
  ExecuteSql(St);
  st := 'UPDATE MAINTIEN SET PMT_DATEDEBUT = "' + usdatetime(strtodate('31/12/1900')) + '", PMT_DATEFIN = "' + usdatetime(strtodate('31/12/1900')) + '" WHERE PMT_DATEDEBUT = "' + usdatetime(strtodate('01/01/1959')) + '" AND PMT_DATEFIN = "' + usdatetime(strtodate('01/01/1959')) + '"';
  ExecuteSql(St);
  ExecuteSql('UPDATE MOTIFABSENCE SET PMA_RUBRIQUEJ="",PMA_ALIMENTJ="",PMA_PROFILABSJ="" WHERE PMA_PREDEFINI <> "CEG"');
  ExecuteSql('UPDATE MOTIFABSENCE SET PMA_RUBRIQUEJ=PMA_RUBRIQUE,PMA_ALIMENTJ=PMA_ALIMENT,PMA_PROFILABSJ=PMA_PROFILABS WHERE PMA_JOURHEURE="JOU" AND PMA_PREDEFINI <> "CEG"');
  ExecuteSql('UPDATE MOTIFABSENCE SET PMA_RUBRIQUE="",PMA_ALIMENT="",PMA_PROFILABS="" WHERE PMA_JOURHEURE="JOU" AND PMA_PREDEFINI <> "CEG" ');

End;


Procedure MajVer738 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GPAO
    InsertChoixCode('YTN','M','Des montants','Montants','');
    // GC
    //Gestion FAR FAE :
    ExecuteSql('UPDATE TIERSPIECE SET GTP_FAR_FAE=""');
    ExecuteSql('UPDATE TXCPTTVA SET TV_CPTACHFARFAE="",  TV_CPTVTEFARFAE=""');
    ExecuteSql('UPDATE PARPIECE SET GPP_FAR_FAE="",GPP_PIECESAV="-"');
    // Gestion mode dégradé de la caisse :
    ExecuteSql('UPDATE PARCAISSE SET GPK_AUTDEGRADE="-", GPK_REPTICKETS="", GPK_ARTICLEDEGRADE=""');

    AglNettoieListes('RTRECHDOCGED', 'RTD_DOCGUID;YDF_FILEGUID',nil, 'RTD_DOCID;YDF_FILEID');
    AglNettoieListes('RTRECHDOCGEDGLO', 'RTD_DOCGUID;YDF_FILEGUID',nil, 'RTD_DOCID;YDF_FILEID');


  end;

  ExecuteSQL('UPDATE RUBRIQUE SET RB_DATEVALIDITE="' + USDATETIME(iDate2099) + '"');
  AGLNettoieListes('CPRUBRIQUE;RUBRIQUE','RB_NODOSSIER;RB_PREDEFINI;RB_DATEVALIDITE');

  //PAIE
  ExecuteSql('UPDATE FORMATIONS SET PFO_CENTREFORMGU="" WHERE PFO_CENTREFORMGU IS NULL');
  ExecuteSql('UPDATE FORMATIONS SET PFO_ORGCOLLECTGU="" WHERE PFO_ORGCOLLECTGU IS NULL');
  ExecuteSql('UPDATE EMPLOIINTERIM SET PEI_CENTREFORMGU="" WHERE PEI_CENTREFORM IS NULL');
  ExecuteSql('UPDATE EMPLOIINTERIM SET PEI_AGENCEINTGU="" WHERE PEI_AGENCEINTGU IS NULL');
  ExecuteSql('UPDATE FRAISSALPLAF SET PFP_ORGCOLLECTGU="" WHERE PFP_ORGCOLLECTGU IS NULL');
  ExecuteSql('UPDATE INVESTFORMATION SET PIF_ORGCOLLECTGU="" WHERE PIF_ORGCOLLECTGU IS NULL');
  ExecuteSql('UPDATE SESSIONSTAGE SET PSS_CENTREFORMGU="" WHERE PSS_CENTREFORMGU IS NULL');
  ExecuteSql('UPDATE SESSIONSTAGE SET PSS_ORGCOLLECTSGU="" WHERE PSS_ORGCOLLECTSGU IS NULL');
  ExecuteSql('UPDATE SESSIONSTAGE SET PSS_ORGCOLLECTPGU="" WHERE PSS_ORGCOLLECTPGU IS NULL');
  ExecuteSql('UPDATE STAGE SET PST_CENTREFORMGU="" WHERE PST_CENTREFORMGU IS NULL');
  ExecuteSql('UPDATE STAGE SET PST_ORGCOLLECTPGU="" WHERE PST_ORGCOLLECTPGU IS NULL');
  ExecuteSql('UPDATE STAGE SET PST_ORGCOLLECTSGU="" WHERE PST_ORGCOLLECTSGU IS NULL');
  ExecuteSql('UPDATE ENVOIFORMATION SET PVF_ORGCOLLECTGU="" WHERE PVF_ORGCOLLECTGU IS NULL');
  ExecuteSql('UPDATE VISITEMEDTRAV SET PVM_MEDTRAVGU="" WHERE PVM_MEDTRAVGU IS NULL');
  ExecuteSql('UPDATE RETENUESALAIRE SET PRE_BENEFRSGU="" WHERE PRE_BENEFRSGU IS NULL');
  ExecuteSql('UPDATE ETABCOMPL SET ETB_CODEDDTEFPGU="" WHERE ETB_CODEDDTEFPGU IS NULL');
  ExecuteSql('UPDATE ETABCOMPL SET ETB_MEDTRAVGU="" WHERE ETB_MEDTRAVGU IS NULL');
  AglNettoieListes('PGANNUAIRE', 'ANN_GUIDPER',nil);


End;

Procedure MajVer739 ;
var st: string ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSQL('update JUEVENEMENT set JEV_OCCURENCEEVT="SIN", JEV_FOREIGNID="", JEV_FOREIGNAPP=""');
    // DP
    If IsMonoOuCommune then
    begin
      ExecuteSQL('update ANNUBIS set ANB_GUIDPERGRPINT=""');
      ExecuteSQL('update DPTABCOMPTA set DTC_TOTACTIFBILAN=0, DTC_TOTPASSIFBILAN=0');
      ExecuteSQL('update DOSSIER set DOS_VERSIONSOC=0');
      ExecuteSQL('update DOSSIER set DOS_ABSENT="X" where DOS_VERROU="ABS"');
      ExecuteSQL('update DOSSIER set DOS_ABSENT="-" where DOS_VERROU<>"ABS"');
    end;

    ExecuteSQL('update YGEDDICO set YGD_EWSREGLEPUB=""');
    // GC
    ExecuteSQL('UPDATE PARPIECE SET GPP_GEREARTICLELIE="AUT" ');
    ExecuteSQL('update articlelie set gal_article=isnull((select ga_article from article where ga_codearticle=gal_article and ga_statutart in ("UNI","GEN")),gal_article) where len(gal_article)<34 ');
    //GPAO

    ExecuteSQL('UPDATE PORT SET GPO_PDRNIV="" WHERE GPO_PDRNIV IS NULL');
    ExecuteSQL('UPDATE PORT SET GPO_PDRNIV="002" WHERE GPO_TYPEPDR<>""');
    ExecuteSQL('UPDATE PORT SET GPO_PDRMTBASE="" WHERE GPO_PDRMTBASE IS NULL');
    ExecuteSQL('UPDATE PORT SET GPO_PDRMTBASE="TOT" WHERE GPO_TYPEPDR<>""');
    ExecuteSQL('UPDATE WPDRTET SET WPE_VALEURPDRD=0'
     + ', WPE_VALPDRLIBRE01=0, WPE_VALPDRLIBRE02=0, WPE_VALPDRLIBRE03=0, WPE_VALPDRLIBRE04=0, WPE_VALPDRLIBRE05=0, WPE_VALPDRLIBRE06=0,'
     + ' WPE_VALPDRLIBRE07=0, WPE_VALPDRLIBRE08=0, WPE_VALPDRLIBRE09=0, WPE_VALPDRLIBRE10=0, WPE_VALPDRLIBRE11=0, WPE_VALPDRLIBRE12=0'
     + ' WHERE WPE_VALEURPDRD IS NULL');
    ExecuteSQL('UPDATE WPDRTYPE SET WRT_MAJDPADPR="-" WHERE WRT_MAJDPADPR IS NULL');
    ExecuteSQL('UPDATE WPARAM SET WPA_BOOLEAN10="-"');
    ExecuteSQL('UPDATE PORT SET GPO_PDRNIV="" WHERE GPO_PDRNIV IS NULL');
    ExecuteSQL('UPDATE PORT SET GPO_PDRNIV="002" WHERE GPO_TYPEPDR<>""');
    ExecuteSQL('UPDATE PORT SET GPO_PDRMTBASE="" WHERE GPO_PDRMTBASE IS NULL');
    ExecuteSQL('UPDATE PORT SET GPO_PDRMTBASE="TOT" WHERE GPO_TYPEPDR<>""');

    ExecuteSQL('UPDATE WPDRTET SET WPE_VALEURPDRD=0'
     + ', WPE_VALPDRLIBRE01=0, WPE_VALPDRLIBRE02=0, WPE_VALPDRLIBRE03=0, WPE_VALPDRLIBRE04=0, WPE_VALPDRLIBRE05=0, WPE_VALPDRLIBRE06=0,'
     + ' WPE_VALPDRLIBRE07=0, WPE_VALPDRLIBRE08=0, WPE_VALPDRLIBRE09=0, WPE_VALPDRLIBRE10=0, WPE_VALPDRLIBRE11=0, WPE_VALPDRLIBRE12=0'
     + ' WHERE WPE_VALEURPDRD IS NULL');


    ExecuteSQL('UPDATE WPDRTYPE SET WRT_MAJDPADPR="-" WHERE WRT_MAJDPADPR IS NULL');
    ExecuteSQL('UPDATE WPARAM SET WPA_BOOLEAN10="-"');
    ExecuteSQL('UPDATE PARACTIONS set RPA_QNATUREACTION="DIV", RPA_NATURETRAVAIL="" where RPA_QNATUREACTION is null');

   { ExecuteSQL('UPDATE ACTIONS SET RAC_IDENTIFIANT=0, RAC_QORIGINERQ="",RAC_QUALITETYPE="",RAC_TYPELOCA="",'
                            +'RAC_QDEMDEROGNUM=0,RAC_QPLANCORRNUM=0,RAC_QNCNUM=0,RAC_NATURETRAVAIL="",'
                            +'RAC_LIGNEORDRE=0,RAC_AREALISERPAR="",RAC_QPCTAVANCT=0,RAC_REALISEELE="'+string(UsDateTime(iDate1900))+'",'
                            +'RAC_REALISEEPAR="",RAC_VERIFIEELE="'+ string(UsDateTime(iDate1900))+'",'
                            +'RAC_VERIFIEEPAR="",RAC_EFFICACITE="",RAC_EFFJUGEELE="'+string(UsDateTime(iDate1900))+'",'
                            +'RAC_EFFJUGEEPAR="",RAC_CLOTUREELE="'+string(UsDateTime(iDate1900))+'",RAC_CLOTUREEPAR="" '
                            +'WHERE RAC_IDENTIFIANT is null');    }

    { Mise à jour du contexte de jetons pour le colisage }
    ExecuteSQL('UPDATE WJETONS'
             + ' SET WJT_CONTEXTE="UL"||(SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_SSCCCHAREXT")'
             + ' WHERE WJT_PREFIXE="GCS" AND WJT_CONTEXTE="ULI"');

    { Mise à jour du jeton en reformant le code unité logistique le plus grand }
    if ExisteSQL('SELECT 1 FROM COLISAGE WHERE GCS_ESTUL="X" AND GCS_ORDRE>0') then
    begin
      ExecuteSQL('UPDATE WJETONS'
               + ' SET WJT_JETON=(SELECT ##TOP 1## CAST(SUBSTRING(GCS_SSCC, 3+LEN(TRIM((SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_CNUF"))),'
               +                                                         ' 15-LEN(TRIM((SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_CNUF")))'
               +                                      ') AS INTEGER) + 1'
               +                ' FROM COLISAGE'
               +                ' WHERE GCS_ESTUL="X"'
               +                ' AND GCS_ORDRE>0'
               +                ' ORDER BY CAST(SUBSTRING(GCS_SSCC, 3+LEN(TRIM((SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_CNUF"))), '
               +                                                 ' 15-LEN(TRIM((SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_CNUF")))'
               +                              ') AS INTEGER) DESC'
               +                ')'
               + ' WHERE WJT_PREFIXE="GCS" AND WJT_CONTEXTE="UL"||(SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_SSCCCHAREXT")');
    end;

    { Mise à jour table EDILIGNE }
    ExecuteSQL('UPDATE EDILIGNE SET ELI_LIGNEORILU="", ELI_LIGNEORI="", ELI_QUALIFQTELU=""'
             + ' WHERE ELI_LIGNEORILU IS NULL'
             + ' AND ELI_LIGNEORI IS NULL'
             + ' AND ELI_QUALIFQTELU IS NULL');

    { Mise à jour table EDIMESSAGES }
    ExecuteSQL('UPDATE EDIMESSAGES SET'
              + ' EMG_BLOREFEXTGPVID=IIF((EMG_TYPETRANS="REC" AND EMG_CODEFAMILLE<>"DEL"), "NON", ""),'
              + ' EMG_BLOREFEXTGLVID=IIF((EMG_TYPETRANS="REC" AND EMG_CODEFAMILLE<>"DEL"), "NON", ""),'
              + ' EMG_BLOAJUSTEQTE=IIF((EMG_TYPETRANS="REC" AND EMG_CODEFAMILLE<>"DEL"), "NON", ""),'
              + ' EMG_BLOARTICLE=IIF((EMG_TYPETRANS="REC" AND EMG_CODEFAMILLE<>"DEL"), "NON", ""),'
              + ' EMG_BLOFORMULE=IIF((EMG_TYPETRANS="REC" AND EMG_CODEFAMILLE<>"DEL"), "NON", ""),'
              + ' EMG_BLOLIGNEORI=IIF((EMG_TYPETRANS="REC" AND EMG_CODEFAMILLE<>"DEL"), "NON", ""),'
              + ' EMG_BLOPRIXZERO=IIF((EMG_TYPETRANS="REC" AND EMG_CODEFAMILLE<>"DEL"), "NON", ""),'
              + ' EMG_BLOSERIELOTEXT=IIF((EMG_TYPETRANS="REC" AND EMG_CODEFAMILLE<>"DEL"), "NON", ""),'
              + ' EMG_ALERTEMAIL="",'
              + ' EMG_EMAIL="",'
              + ' EMG_DATECONFIG=EMG_DATECREATION,'
              + ' EMG_ALIMAPPRO="",'
              + ' EMG_CMDLINEBEFORE="",'
              + ' EMG_CMDBEFOREMODAL="-",'
              + ' EMG_CMDBEFOREHIDE="-",'
              + ' EMG_CMDLINEAFTER="",'
              + ' EMG_CMDAFTERMODAL="-",'
              + ' EMG_CMDAFTERHIDE="-",'
              + ' EMG_AUTOPRINT="-",'
              + ' EMG_PRINTERSLIST="",'
              + ' EMG_NBCOPY=1,'
              + ' EMG_AUTOMATEACTIVE="-",'
              + ' EMG_PERIODE=1,'
              + ' EMG_UNITETEMPS="H",'
              + ' EMG_NEXTDATEPLANIF="' + UsDateTime(iDate2099) + '"'
              + ' WHERE EMG_ALERTEMAIL IS NULL'
              + ' AND EMG_BLOREFEXTGPVID IS NULL'
              + ' AND EMG_BLOREFEXTGLVID IS NULL'
              + ' AND EMG_BLOAJUSTEQTE IS NULL'
              + ' AND EMG_BLOARTICLE IS NULL'
              + ' AND EMG_BLOFORMULE IS NULL'
              + ' AND EMG_BLOLIGNEORI IS NULL'
              + ' AND EMG_BLOPRIXZERO IS NULL'
              + ' AND EMG_BLOSERIELOTEXT IS NULL'
              + ' AND EMG_EMAIL IS NULL'
              + ' AND EMG_DATECONFIG IS NULL'
              + ' AND EMG_ALIMAPPRO IS NULL'
              + ' AND EMG_CMDLINEBEFORE IS NULL'
              + ' AND EMG_CMDBEFOREMODAL IS NULL'
              + ' AND EMG_CMDBEFOREHIDE IS NULL'
              + ' AND EMG_CMDLINEAFTER IS NULL'
              + ' AND EMG_CMDAFTERMODAL IS NULL'
              + ' AND EMG_CMDAFTERHIDE IS NULL'
              + ' AND EMG_AUTOPRINT IS NULL'
              + ' AND EMG_PRINTERSLIST IS NULL'
              + ' AND EMG_NBCOPY IS NULL'
              + ' AND EMG_AUTOMATEACTIVE IS NULL'
              + ' AND EMG_PERIODE IS NULL'
              + ' AND EMG_UNITETEMPS IS NULL'
              + ' AND EMG_NEXTDATEPLANIF IS NULL');

    { Suppression des éventuels paramètres d'automate EDI
     (ancienne version - se trouve maintenant au niveau du message)  }
    ExecuteSQL('DELETE FROM WPARAM WHERE WPA_CODEPARAM="EMG_AUTOMATE"');

    { Mise à jour table LIGNECOMPL }
    {1. 1er passage}
    ExecuteSQL('UPDATE LIGNECOMPL'
             + ' SET GLC_AFFCONTREM=(SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_GCAFFCONTREM"), GLC_AUTOGENCONTREM="-"'
             + ' WHERE TRIM(GLC_NUMERO)||"~"||TRIM(STR(GLC_NUMERO))||"~"||TRIM(GLC_SOUCHE)||"~"||TRIM(STR(GLC_INDICEG))||"~"||TRIM(STR(GLC_NUMORDRE))'
             +        '=(SELECT TRIM(GL_NUMERO)||"~"||TRIM(STR(GL_NUMERO))||"~"||TRIM(GL_SOUCHE)||"~"||TRIM(STR(GL_INDICEG))||"~"||TRIM(STR(GL_NUMORDRE))'
             +         ' FROM LIGNE'
             +         ' WHERE GL_NATUREPIECEG=GLC_NATUREPIECEG'
             +         ' AND GL_NUMERO=GLC_NUMERO'
             +         ' AND GL_SOUCHE=GLC_SOUCHE'
             +         ' AND GL_INDICEG=GLC_INDICEG'
             +         ' AND GL_NUMORDRE=GLC_NUMORDRE'
             +         ' AND GL_ENCONTREMARQUE="X")'
             + ' AND GLC_AUTOGENCONTREM IS NULL'
             + ' AND GLC_AFFCONTREM IS NULL');
    {2. 2ème passage}
    ExecuteSQL('UPDATE LIGNECOMPL'
             + ' SET GLC_AFFCONTREM="", GLC_AUTOGENCONTREM="-"'
             + ' WHERE GLC_AUTOGENCONTREM IS NULL'
             + ' AND GLC_AFFCONTREM IS NULL');

    { Modif. de listes }
    AglNettoieListesPlus('EDIMESSAGES', 'EMG_IDENTIFIANT;EMG_TYPETRANS;EMG_CODEMESSAGE', nil, True, '');
    // GC
    ExecuteSQL('UPDATE PARPIECE SET GPP_GEREARTICLELIE="AUT"');
    ExecuteSQL('update articlelie set gal_article=isnull((select ga_article from article where ga_codearticle=gal_article and ga_statutart in ("UNI","GEN")),gal_article) where len(gal_article)<34');
  end;

End;

Procedure MajVer740 ;
var st: string ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    RT_InsertTablettesWIV;
    ExecuteSQL('UPDATE PARAMSOC SET SOC_DESIGN="36;Service après vente;0;W" WHERE SOC_NOM="SCO_SAV"');
    //GPAO
    ExecuteSQL('UPDATE TIERSCOMPL SET YTC_PALMARESTRA = "-", YTC_CODEPORT="" WHERE YTC_PALMARESTRA IS NULL');
    ExecuteSQL('UPDATE WPARAMFONCTION SET WPF_SECTEURGEO="", WPF_PAYS="", WPF_CODEPOSTAL="",WPF_REGION="",WPF_VENTEACHAT="" WHERE WPF_SECTEURGEO IS NULL');
    ExecuteSQL('UPDATE PIEDPORT SET GPT_TIERSFRAIS="", GPT_ORIGINE="" WHERE GPT_TIERSFRAIS IS NULL');
    ExecuteSQL('UPDATE LIGNEFRAIS SET LF_TIERSFRAIS="" WHERE LF_TIERSFRAIS IS NULL');
    AglNettoieListesPlus('GCSTKDISPO','GA_COEFCONVQTEACH;GA_COEFCONVQTEVTE;GA_COEFPROD;GA_COEFCONVCONSO' ,nil);
    //MES
    ExecuteSql('UPDATE WORDREGAMME SET WOG_AUTOMATISME="" WHERE WOG_AUTOMATISME IS NULL');
    MajMarkMESApres;
	  //GIGA
    ExecuteSql('UPDATE PROFILGENER SET APG_TRIPROFIL=""');
    ExecuteSql('UPDATE PARPIECE SET GPP_REPRISEENTAFF="",GPP_REPRISELIGAFF=""');
    AglNettoieListes('AFMULEAFFAIRE', 'EAF_TIERS',nil);
    //
    ExecuteSql ('Update activite set act_modesaisfrais="NOR"');
    ExecuteSql ('Update Eactivite set EAC_modesaisfrais="NOR"');
	  //fin GIGA
    If IsMonoOuCommune then
    begin
      if not (V_PGI.Driver in [dbORACLE7, dbORACLE8 ,dbORACLE9,dbDB2]) then
      begin
        ExecuteSQL( 'UPDATE ANNUAIRE SET'
           +' ANN_CODEINSEE     = (SELECT TOP 1 YFJ_CODEINSEE FROM YFORMESINSEE'
                              +' WHERE ANN_FORME = YFJ_FORME AND YFJ_FORME <> "" AND YFJ_NONACTIF = "-"'
                              +' ORDER BY YFJ_FORME, YFJ_NIVEAU),'
           +' ANN_COOP          = (SELECT TOP 1 YFJ_COOP FROM YFORMESINSEE'
                              +' WHERE ANN_FORME = YFJ_FORME AND YFJ_FORME <> "" AND YFJ_NONACTIF = "-"'
                              +' ORDER BY YFJ_FORME, YFJ_NIVEAU),'
           +' ANN_FORMEGRPPRIVE = (SELECT TOP 1 YFJ_FORMEGRPPRIVE FROM YFORMESINSEE'
                              +' WHERE ANN_FORME = YFJ_FORME AND YFJ_FORME <> "" AND YFJ_NONACTIF = "-"'
                              +' ORDER BY YFJ_FORME, YFJ_NIVEAU),'
           +' ANN_FORMESTE      = (SELECT TOP 1 YFJ_FORMESTE  FROM YFORMESINSEE'
                              +' WHERE ANN_FORME = YFJ_FORME AND YFJ_FORME <> "" AND YFJ_NONACTIF = "-"'
                              +' ORDER BY YFJ_FORME, YFJ_NIVEAU),'
           +' ANN_FORMESCI      = (SELECT TOP 1 YFJ_FORMESCI FROM YFORMESINSEE'
                              +' WHERE ANN_FORME = YFJ_FORME AND YFJ_FORME <> "" AND YFJ_NONACTIF = "-"'
                              +' ORDER BY YFJ_FORME, YFJ_NIVEAU)' );
      end;
      ExecuteSQL( 'UPDATE ANNUAIRE SET ANN_CODEINSEE     = ""  WHERE ANN_CODEINSEE     IS NULL');
      ExecuteSQL( 'UPDATE ANNUAIRE SET ANN_COOP          = "-" WHERE ANN_COOP          IS NULL');
      ExecuteSQL( 'UPDATE ANNUAIRE SET ANN_FORMEGRPPRIVE = ""  WHERE ANN_FORMEGRPPRIVE IS NULL');
      ExecuteSQL( 'UPDATE ANNUAIRE SET ANN_FORMESTE      = ""  WHERE ANN_FORMESTE      IS NULL');
      ExecuteSQL( 'UPDATE ANNUAIRE SET ANN_FORMESCI      = ""  WHERE ANN_FORMESCI      IS NULL');

      if not (V_PGI.Driver in [dbORACLE7, dbORACLE8 ,dbORACLE9,dbDB2]) then
      begin
        ExecuteSQL( 'UPDATE DPORGA SET'
         +' DOR_REGLEFISC = (SELECT TOP 1 YFJ_REGLEFISC FROM YFORMESINSEE, ANNUAIRE'
                                +' WHERE ANN_GUIDPER = DOR_GUIDPER AND ANN_FORME = YFJ_FORME'
                                +' AND YFJ_FORME <> ""  AND YFJ_NONACTIF = "-"'
                                +' ORDER BY YFJ_FORME, YFJ_NIVEAU),'
         +' DOR_SECTIONBNC = (SELECT TOP 1 YFJ_SECTIONBNC FROM YFORMESINSEE, ANNUAIRE'
                                +' WHERE ANN_GUIDPER = DOR_GUIDPER AND ANN_FORME = YFJ_FORME'
                                +' AND YFJ_FORME <> ""  AND YFJ_NONACTIF = "-"'
                                +' ORDER BY YFJ_FORME, YFJ_NIVEAU),'
         +' DOR_FORMEASSO = (SELECT TOP 1 YFJ_FORMEASSO FROM YFORMESINSEE, ANNUAIRE'
                                +' WHERE ANN_GUIDPER = DOR_GUIDPER AND ANN_FORME = YFJ_FORME'
                                +' AND YFJ_FORME <> ""  AND YFJ_NONACTIF = "-"'
                                +' ORDER BY YFJ_FORME, YFJ_NIVEAU)' );
      end;
      //CHR
      ExecuteSQL('UPDATE HRCAISSE SET HRC_GESTIONPDA="-" WHERE HRC_GESTIONPDA IS NULL');
      ExecuteSQL('UPDATE HRPARAMPLANNING SET HPP_LPARAMPLANNING="" WHERE HPP_LPARAMPLANNING IS NULL');

      ExecuteSQL('UPDATE HRDOSSIER SET HDC_CBINTERNET="",HDC_CBLIBELLE="",HDC_DATEEXPIRE="",HDC_TYPECARTE="",HDC_CBNUMCTRL="",HDC_CBNUMAUTOR="",HDC_NUMCHEQUE=""  WHERE HDC_CBINTERNET IS NULL');

      //GC
      ExecuteSQL('update etables set edt_nbenreg=0, edt_taillemoyenne=0, edt_tempstraite=0, edt_active="X"');
      ExecuteSQL('UPDATE PARCAISSE SET GPK_IP="", GPK_IMPMODTICDGD=""');
      //GPAO
      ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA="CMC" WHERE SOC_NOM="SO_CDEMARCHEVTE"');
      ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA="CC;" WHERE SOC_NOM="SO_CDEMARCHEAPPELVTE"');
      ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA="GL_DATELIVRAISON" WHERE SOC_NOM="SO_CDEMARCHEAPPELDATEVTE"');
      //SERVANTISSIMO
      ExecuteSQL('UPDATE IMOREF SET IRF_CONFIDENTIEL ="0"');
    end;

  end;

// TRESO
AglNettoieListes('TRECRITURERAPPRO', 'TE_NODOSSIER;',nil);
AglNettoieListes('TRMODIFIERUBRIQUE', 'TE_NODOSSIER;',nil);
AglNettoieListes('TRECRITURE', 'TE_NODOSSIER;',nil);
AglNettoieListes('TRPREVISIONNELLES', 'TE_NODOSSIER;',nil);
AglNettoieListes('TRLISTESYNCHRO', 'TE_NODOSSIER;TE_NUMTRANSAC;TE_NUMLIGNE;TE_DATEVALEUR;TE_GENERAL;TE_MONTANTDEV;TE_DEVISE;TE_JOURNAL;TE_QUALIFORIGINE;TE_LIBELLE;TE_NUMEROPIECE;TE_CODECIB;TE_CODEFLUX;TE_DATECOMPTABLE;TE_NATURE;TE_EXERCICE;TE_DATECREATION; ',nil);
AglNettoieListes('TRSUPPRECRITURE', 'TE_NODOSSIER;',nil);
ExecuteSQL('UPDATE COURTSTERMES SET TCT_NODOSSIER = "000000"');
ExecuteSQL('UPDATE EQUILIBRAGE SET TEQ_SNODOSSIER = "000000"');
ExecuteSQL('UPDATE EQUILIBRAGE SET TEQ_DNODOSSIER = "000000"');
ExecuteSQL('UPDATE TRECRITURE SET TE_NODOSSIER = "000000"');
ExecuteSQL('UPDATE TROPCVM SET TOP_NODOSSIER = "000000"');
ExecuteSQL('UPDATE TRVENTEOPCVM SET TVE_NODOSSIER = "000000"');
ExecuteSQL('UPDATE BANQUECP SET BQ_NODOSSIER = "000000"');
//
ExecuteSQL('UPDATE SUIVCPTA SET SC_CONTROLEQTE = "-"' ) ;
ExecuteSQL('UPDATE SECTION SET S_FINCHANTIER = "' + UsDateTime(iDate2099) + '" WHERE S_FINCHANTIER = "' + UsDateTime(iDate1900) + '"');
//Compta
ExecuteSQL('UPDATE CBALSIT SET BSI_DATEMODIF="' + USDATETIME(iDate1900) + '",BSI_TYPEECR="N"');
ExecuteSQL('UPDATE CBALSITECR SET BSE_QUALIFPIECE="",BSE_ETABLISSEMENT="",BSE_DEVISE=""');
//PAie

  St := 'UPDATE MASQUESAISRUB SET PMR_AIDECOL1="",PMR_AIDECOL2="",PMR_AIDECOL3="",PMR_AIDECOL4="",'+
        'PMR_AIDECOL5="",PMR_AIDECOL6="",PMR_AIDECOL7="",PMR_UTILISATEUR="",PMR_DATECREATION="'+
        UsDateTime(Idate1900)+'",PMR_DATEMODIF="'+UsDateTime(Idate1900)+'"';
  ExecuteSql(St);

  ExecuteSql('UPDATE CRITMAINTIEN SET PCM_VALCATEGORIE = PCM_VALCATEG');
  ExecuteSql('UPDATE FORMATIONS SET PFO_SIGNIFICATIF="-"');
  ExecuteSql('UPDATE RHCOMPETRESSOURCE SET PCH_PGPROVCOMP="",PCH_RHPROFIL=""');
  ExecuteSql('UPDATE RHCOMPETENCES SET PCO_RHNATURECOMP="",PCO_DUREEACQUIS=0,'+
  'PCO_RHVALIDEURCOMP="",PCO_RHMODEVALID="",PCO_RHNIVEAUCOMP="",PCO_RHMACROCOMP=""');
  ExecuteSql('UPDATE REFERENTIELMETIER SET PRF_CODEEMPLOI=""');

  St:= 'UPDATE EMETTEURSOCIAL SET PET_CRECIVILITE="", PET_CRENOM="", PET_CIVIL1DADSU="", PET_TEL1DADSU="",'+
       ' PET_FAX1DADSU="", PET_CIVIL2DADSU="", PET_TEL2DADSU="", PET_FAX2DADSU="", PET_CIVIL3DADSU="",'+
       ' PET_TEL3DADSU="", PET_FAX3DADSU=""';
  ExecuteSql(St);
  ExecuteSql('UPDATE EMETTEURSOCIAL SET PET_TEL1DADSU=PET_APPEL1DUDS WHERE PET_MEDIADUDS1="01"');
  ExecuteSql('UPDATE EMETTEURSOCIAL SET PET_APPEL1DUDS="", PET_MEDIADUDS1="" WHERE PET_MEDIADUDS1="01"');
  ExecuteSql('UPDATE EMETTEURSOCIAL SET PET_FAX1DADSU=PET_APPEL1DUDS WHERE PET_MEDIADUDS1="02"');
  ExecuteSql('UPDATE EMETTEURSOCIAL SET PET_APPEL1DUDS="", PET_MEDIADUDS1="" WHERE PET_MEDIADUDS1="02"');
  ExecuteSql('UPDATE EMETTEURSOCIAL SET PET_MEDIADUDS1="" WHERE PET_MEDIADUDS1="03"');
  ExecuteSql('UPDATE EMETTEURSOCIAL SET PET_TEL2DADSU=PET_APPEL2DUDS WHERE PET_MEDIADUDS2="01"');
  ExecuteSql('UPDATE EMETTEURSOCIAL SET PET_APPEL2DUDS="", PET_MEDIADUDS2="" WHERE PET_MEDIADUDS2="01"');
  ExecuteSql('UPDATE EMETTEURSOCIAL SET PET_FAX2DADSU=PET_APPEL2DUDS WHERE PET_MEDIADUDS2="02"');
  ExecuteSql('UPDATE EMETTEURSOCIAL SET PET_APPEL2DUDS="", PET_MEDIADUDS2="" WHERE PET_MEDIADUDS2="02"');
  ExecuteSql('UPDATE EMETTEURSOCIAL SET PET_MEDIADUDS2="" WHERE PET_MEDIADUDS2="03"');
  ExecuteSql('UPDATE EMETTEURSOCIAL SET PET_TEL3DADSU=PET_APPEL3DUDS WHERE PET_MEDIADUDS3="01"');
  ExecuteSql('UPDATE EMETTEURSOCIAL SET PET_APPEL3DUDS="", PET_MEDIADUDS3="" WHERE PET_MEDIADUDS3="01"');
  ExecuteSql('UPDATE EMETTEURSOCIAL SET PET_FAX3DADSU=PET_APPEL3DUDS WHERE PET_MEDIADUDS3="02"');
  ExecuteSql('UPDATE EMETTEURSOCIAL SET PET_APPEL3DUDS="", PET_MEDIADUDS3="" WHERE PET_MEDIADUDS3="02"');
  ExecuteSql('UPDATE EMETTEURSOCIAL SET PET_MEDIADUDS3="" WHERE PET_MEDIADUDS3="03"');

  ExecuteSql('UPDATE MOTIFABSENCE SET PMA_ALIMNETJ="-",PMA_ALIMNETH="-"');
  ExecuteSql('UPDATE MOTIFABSENCE SET PMA_ALIMNETJ="X" WHERE PMA_RUBRIQUEJ <> ""');
  ExecuteSql('UPDATE MOTIFABSENCE SET PMA_ALIMNETH="X" WHERE PMA_RUBRIQUE <> ""');
  ExecuteSql('UPDATE REMUNERATION SET PRM_METHODARRONDI="",PRM_REMLIEREGUL="" WHERE PRM_PREDEFINI<> "CEG"');

  ExecuteSql('UPDATE ETABCOMPL SET ETB_PAIEVALOMS="",ETB_MSAUNITEGES="",ETB_RATTACHEHANDI=""');
  ExecuteSql('UPDATE ETABCOMPL SET ETB_PAIEVALOMS="ANT" WHERE ETB_CONGESPAYES="X"');
  ExecuteSql('UPDATE SALARIES SET PSA_TYPPAIEVALOMS="ETB", PSA_PAIEVALOMS=""');
  St:= 'UPDATE SALARIES SET PSA_TYPPAIEVALOMS="ETB", PSA_PAIEVALOMS="ANT" WHERE PSA_ETABLISSEMENT IN'+
	' (SELECT ETB_ETABLISSEMENT FROM ETABCOMPL WHERE ETB_CONGESPAYES="X")';
  ExecuteSql(St);
  St:= 'update motifabsence set PMA_TYPEPERMAXI="",PMA_PERMAXI=0,PMA_OUVRES="-",PMA_OUVRABLE="-",PMA_CALENDSAL="-"';
  ExecuteSql(St);
  ExecuteSql('UPDATE MOTIFABSENCE SET PMA_CALENDSAL="X" WHERE PMA_CALENDCIVIL="-"');
  St:= 'UPDATE SALARIES SET PSA_DATEVIEACTIVE="'+UsDateTime(IDate1900)+'"';
  ExecuteSql(St);
  St:= 'UPDATE DEPORTSAL SET PSE_MSATYPUNITEG="ETB", PSE_MSAUNITEGES="", PSE_RESPONSAUG=PSE_RESPONSVAR,'+
	' PSE_ASSISTAUG=PSE_ASSISTVAR';
  ExecuteSql(St);
  St:= 'UPDATE SERVICES SET PGS_RESPONSAUG=PGS_RESPONSVAR, PGS_SECRETAIREAUG=PGS_SECRETAIREVAR,'+
	' PGS_ADJOINTAUG=PGS_ADJOINTVAR';
  ExecuteSql(St);
  ExecuteSql('UPDATE FRAISSALPLAF SET PFP_CATEGFRAISFOR=""');
  ExecuteSql('UPDATE SESSIONSTAGE SET PSS_NUMSESSION=-1, PSS_TYPESESSSTAGE="IND"');
  St:= 'UPDATE PGHISTODETAIL SET PHD_DATEDEBVALID="'+UsDateTime(IDate1900)+'", PHD_RUBRIQUE=""';
  ExecuteSql(St);
  St:= 'UPDATE SALARIES SET PSA_TYPCONVENTION="ETB" WHERE PSA_CONVENTION IN (SELECT ETB_CONVENTION FROM ETABCOMPL'+
	' WHERE ETB_ETABLISSEMENT=PSA_ETABLISSEMENT)';
  ExecuteSql(St);
  St:= 'UPDATE SALARIES SET PSA_TYPCONVENTION="PER" WHERE PSA_CONVENTION NOT IN (SELECT ETB_CONVENTION FROM ETABCOMPL'+
	' WHERE ETB_ETABLISSEMENT=PSA_ETABLISSEMENT)';
  ExecuteSql(St);

  //IMMO
  ExecuteSQL('UPDATE IMMO SET I_SUSDEF="A",I_REGLECESSION="NOR",I_NONDED="-",I_REPRISEINT=0.00,I_REPRISEDEPCEDEE=0.00,I_REPRISEINTCEDEE=0.00,I_DPI="-",I_DPIEC="-",I_CORRECTIONVR=0.00,I_CORVRCEDDE=0.00,I_SUBVENTION="NON",I_SBVPRI=0.00');
  ExecuteSQL('UPDATE IMMO SET I_SBVMTC=0.00,I_SBVPRIC=0.00,I_SBVEC="C",I_SBVDATE="'+UsDateTime(iDate1900)+'",I_CPTSBVR="",I_CPTSBVB="",I_PFR="-",I_COEFDEG=0.00,I_AMTFOR=0.00,I_AMTFORC=0.00,I_ACHFOR="'+UsDateTime(iDate1900)+'"');
  ExecuteSQL('UPDATE IMMO SET I_PRIXACFORC=0.00,I_VNCFOR=0.00,I_DURRESTFOR=0,I_DATEDEBECO=I_DATEAMORT,I_DATEDEBFIS=I_DATEAMORT');

  ExecuteSQL('UPDATE IMMO SET I_DATEDEBECO=I_DATEPIECEA WHERE I_METHODEECO="DEG"');
  ExecuteSQL('UPDATE IMMO SET I_DATEDEBFIS=I_DATEPIECEA WHERE I_METHODEFISC="DEG"');

  ExecuteSQL('UPDATE IMMOLOG SET IL_TAUX=0.0');

  ExecuteSQL('UPDATE IMMOAMOR SET IA_NONDEDUCT=0.00,IA_CESSIONND=0.00,IA_MONTANTSBV=0.00,IA_CESSIONSBV=0.00,IA_MONTANTDPI=0.00,IA_CESSIONDPI=0.00,IA_MONTANTARD=0.00,IA_CESSIONARD=0.00');



End;


// =====   MAJAPRES ================================================

function MAJHalleyApres(VSoc: Integer; MajLab1, MajLab2: TLabel; MajJauge: TProgressBar): boolean;
begin
  L1 := MAJLab1;
  L2 := MAjLab2;
  J := MajJauge;
  if V_PGI.SAV then LogAGL('Début MajApres ' + DateTimeToStr(Now));
  L1.Caption:=TraduireMemoire('Initialisation des nouvelles données');
  L2.Caption:='';
  if V_PGI.NumVersionBase-Vsoc >0 then
     J.Max :=V_PGI.NumVersionBase-Vsoc
  else
     J.Max := 100;
  J.Position:=0;

  //if TOB(V_PGI.TOBSOC)<>Nil then BEGIN TOB(V_PGI.TOBSOC).Free ; V_PGI.TOBSOC:=NIL ; END ;
  //V_PGI.DEDejaCharge:= False;
  //ChargeTablePrefixe(FALSE, TRUE);  // ajout PCS pour s'assurer du chargement en mémoire du nouveau dictionnaire pour utiliser les BOBs

  if VSoc <= 541 then MajVer541;
  if VSoc <= 542 then MajVer542;
  if VSoc <= 543 then MajVer543;
  if VSoc <= 544 then MajVer544;
  if VSoc <= 545 then MajVer545;
  if VSoc <= 546 then MajVer546;
  if VSoc <= 547 then MajVer547;
  if VSoc <= 548 then MajVer548;
  if VSoc <= 549 then MajVer549;
  if VSoc <= 550 then MajVer550;
  if VSoc <= 560 then MajVer560;
  if VSoc <= 561 then MajVer561;
  if VSoc <= 562 then MajVer562;
  if VSoc <= 564 then MajVer564;
  if VSoc <= 565 then MajVer565;
  if VSoc <= 567 then MajVer567;
  if VSoc <= 568 then MajVer568;
  if VSoc <= 569 then MajVer569;
  if VSoc <= 570 then MajVer570;
  if VSoc <= 574 then MajVer574;
  if VSoc <= 580 then MajVer580;
  if VSoc <= 581 then MajVer581(vsoc); // pour version 582
  if VSoc <= 589 then MajVer589;
  if VSoc <= 590 then MajVer590;
  if VSoc <= 594 then MajVer594;
  {$IFDEF CEGID}
  if VSoc <= 595 then MajVer595; // 596 specif SIC CEGID
  {$ENDIF}
  if VSoc <= 601 then MajVer601;
  if VSoc <  603 then MajVer603; // PCS ATTENTION Changement dans la numerotation pour refletter le vrai numero de socref
  if VSoc <  604 then MajVer604;
  if VSoc <  605 then MajVer605;
  if VSoc <  606 then MajVer606;
  if VSoc <  607 then MajVer607;
  if VSoc <  608 then MajVer608;
  if VSoc <  609 then MajVer609;
  if VSoc <  610 then MajVer610;
  if VSoc <  612 then MajVer612;
  if VSoc <  613 then MajVer613;
  if VSoc <  614 then MajVer614;
  if VSoc <  615 then MajVer615;
  if VSoc <  616 then MajVer616;
  if VSoc <  617 then MajVer617;
  if VSoc <  618 then MajVer618;
  if VSoc <  619 then MajVer619;
  if VSoc <  620 then MajVer620;  // diffusion 5.0
  if VSoc <  625 then MajVer625;
  if VSoc <  626 then MajVer626;
  if VSoc <  627 then MajVer627;
  if VSoc <  628 then MajVer628;
  if VSoc <  630 then MajVer630;
  if VSoc <  631 then MajVer631;
  if VSoc <  633 then MajVer633;
  if VSoc <  634 then MajVer634;
  if VSoc <  635 then MajVer635;
  if VSoc <  636 then MajVer636;
  if VSoc <  637 then MajVer637;
  if VSoc <  639 then MajVer639;
  if VSoc <  640 then MajVer640;
  if VSoc <  641 then MajVer641;
  if VSoc <  642 then MajVer642;
  if VSoc <  644 then MajVer644;
  if VSoc <  645 then MajVer645;
  if VSoc <  646 then MajVer646;
  if VSoc <  647 then MajVer647;
  if VSoc <  648 then MajVer648;
  if VSoc <  649 then MajVer649;
  if VSoc <  650 then MajVer650;
  if VSoc <  651 then MajVer651;
  if VSoc <  652 then MajVer652;
  if VSoc <  653 then MajVer653;
  if VSoc <  654 then MajVer654;
  if VSoc <  655 then MajVer655;
  if VSoc <  656 then MajVer656;
  if VSoc <  657 then MajVer657;
  if VSoc <  658 then MajVer658;
  if VSoc <  659 then MajVer659;
  if VSoc <  660 then MajVer660;
  if VSoc <  661 then MajVer661;
  if VSoc <  662 then MajVer662;
  if VSoc <  663 then MajVer663;  //V6.00
  if VSoc <  667 then MajVer667;
  if VSoc <  668 then MajVer668;
  if VSoc <  669 then MajVer669;
  if VSoc <  670 then MajVer670;
  if VSoc <  671 then MajVer671;
  if VSoc <  673 then MajVer673;
  if VSoc <  674 then MajVer674;
  if VSoc <  675 then MajVer675;
  if VSoc <  676 then MajVer676;
 //  if VSoc <  677 then MajVer677;  // V6.20
  if VSoc <  680 then MajVer680;
  if VSoc <  681 then MajVer681;
  if VSoc <  683 then MajVer683;
  if VSoc <  684 then MajVer684;
  if VSoc <  685 then MajVer685;
  if VSoc <  686 then MajVer686;
  if VSoc <  687 then MajVer687;
  if VSoc <  688 then MajVer688;
  if VSoc <  689 then MajVer689;
  if VSoc <  690 then MajVer690;
  if VSoc <  691 then MajVer691;
  if VSoc <  692 then MajVer692;
  if VSoc <  694 then MajVer694;
  if VSoc <  695 then MajVer695;
  if VSoc <  696 then MajVer696;
  if VSoc <  697 then MajVer697;
  if VSoc <  698 then MajVer698;
  if VSoc <  699 then MajVer699;
  if VSoc <  700 then MajVer700;
  if VSoc <  701 then MajVer701;
  if VSoc <  702 then MajVer702;
  if VSoc <  703 then MajVer703;
  if VSoc <  704 then MajVer704;
  if VSoc <  705 then MajVer705;
  if VSoc <  706 then MajVer706;
  if VSoc <  707 then MajVer707;
  if VSoc <  708 then MajVer708;
  if VSoc <  709 then MajVer709;
  if VSoc <  710 then MajVer710;
  if VSoc <  711 then MajVer711;
  if VSoc <  712 then MajVer712;
  if VSoc <  713 then MajVer713;
  if VSoc <  714 then MajVer714;  // Diffusion 6.50
  if VSoc <  720 then MajVer720;
  if VSoc <  721 then MajVer721;
  if VSoc <  722 then MajVer722;
  if VSoc <  723 then MajVer723;
  if VSoc <  724 then MajVer724;
  if VSoc <  725 then MajVer725;
  if VSoc <  726 then MajVer726;
  if VSoc <  727 then MajVer727;
  if VSoc <  728 then MajVer728;
  if VSoc <  729 then MajVer729;
  if VSoc <  730 then MajVer730; // pilotage 6.90
  if VSoc <  734 then MajVer734;
  if VSoc <  735 then MajVer735;
  if VSoc <  736 then MajVer736;
  if VSoc <  737 then MajVer737;
  if VSoc <  738 then MajVer738;
  if VSoc <  739 then MajVer739;
  if VSoc <  740 then MajVer740;


  if not IsDossierPCL then
  begin
    AddNewNaturesGC;
    GPMajSTKNature;
  end;
  MajSouche;
  MAJStandardPaie;
  MAJListeProduits;
  ForceConfidentialiteMenu;
  Result := True;
  if V_PGI.SAV then LogAGL('Fin Mise à jour à ' + DateTimeToStr(Now));
end;

// Recupération des elements nationaux de type PREDEFINI=STD uniquement pour la paie

procedure MAJStandardPaie;
var DD: TDateTime;
  TSoc, TRef: THTable;
begin
  if IsMonoOuCommune then
  begin
    DD := Encodedate(1998, 01, 01);
    // suppression des elemenents nationaux std sans valorisation et à la date du 01/01/1998 dont le theme est Retraite
    ExecuteSQL('DELETE FROM ELTNATIONAUX WHERE PEL_PREDEFINI="STD" AND PEL_MONTANT=0 AND PEL_MONTANTEURO=0  AND PEL_THEMEELT="RET" AND PEL_DATEVALIDITE = "' +
      UsDateTime(DD) + '"');
    TSoc := THTable.Create(Application);
    TSoc.DatabaseName := DBSoc.DatabaseName;
    TSoc.Tablename := 'ELTNATIONAUX';
    TSoc.Indexname := 'PEL_CLE1';
    TSoc.Open;
    TRef := OpenTableRef('ELTNATIONAUX', 'PEL_CLE1');
    while not TRef.EOF do
    begin
      if (not TSoc.Findkey([TRef.FindField('PEL_PREDEFINI').AsString, TRef.FindField('PEL_NODOSSIER').AsString, TRef.FindField('PEL_CODEELT').AsString,
        TRef.FindField('PEL_DATEVALIDITE').AsDateTime])) and (TRef.FindField('PEL_PREDEFINI').AsString = 'STD') then
        AddParamGC(TSoc, TRef);
      TRef.Next;
    end;
    Ferme(TRef);
    TSoc.Close;
    TSoc.Free;
  end;
end;



//============= MAJ PENDANT ===================================================


procedure Pendant568;
begin
  ExecuteSQL('UPDATE RESSOURCEPR SET ARP_TYPEVALO="R" where ARP_TYPEVALO is NULL OR ARP_TYPEVALO=""');
end;

procedure Pendant569;
begin
  ExecuteSql('UPDATE HISTOBULLETIN SET PHB_ORGANISME="...." WHERE PHB_ORGANISME=""');
end;

procedure Pendant582(NomTable: string);
begin
  if (NomTable = 'DUCSENTETE') then ExecuteSQL('update DUCSENTETE SET PDU_NUM=0');
  if (NomTable = 'DUCSDETAIL') then ExecuteSQL('update DUCSDETAIL SET PDD_NUM=0');
end;

procedure Pendant595;
begin
  //Compta
  ExecuteSQL('UPDATE TRFDOSSIER SET TRD_DOSSIERINI=TRD_NODOSSIER WHERE TRD_DOSSIERINI is NULL OR TRD_DOSSIERINI=""');
end;

procedure Pendant602;
Begin
  ExecuteSql('UPDATE JOURFERIE SET AJF_PREDEFINI="DOS",AJF_NODOSSIER="' + V_PGI.Nodossier + '"');
End;

procedure Pendant674(NomTable: string);
var Q : TQuery;
    i : integer;
Begin
  if (NomTable = 'FPBIENS') then
  begin
    // Table des BIENS fiscalite (FPBIENS)
    i := 1;
    Q := OpenSQL('select FPB_N01281 from fpbiens order by fpb_n01200', false);
    while not Q.Eof do
    begin
      Q.Edit;
      Q.FindField('FPB_N01281').AsInteger := i;
      Q.Post;
      Q.Next;
      inc(i);
    end;
    Ferme(Q);
  end;
  if (NomTable = 'FPBIENS_VE') then
  begin
    // Table des Histo.valeurs des biens  (FPBIENS_VE)
    i := 1;
    Q := OpenSQL('select fve_n03705 from fpbiens_ve order by fve_ncode', false);
    while not Q.Eof do
    begin
      Q.Edit;
      Q.FindField('FVE_N03705').AsInteger := i;
      Q.Post;
      Q.Next;
      inc(i);
    end;
    Ferme(Q);
  end;
End;

function MAJHalleyPendant(VSoc: Integer; NomTable: string): boolean;
Begin
  Result := True;
  if ((VSOC < 568) and (NomTable = 'RESSOURCEPR')) then Pendant568;
  if ((VSOC < 569) and (NomTable = 'HISTOBULLETIN')) then Pendant569;
  if ((VSOC < 582) and ((NomTable = 'DUCSENTETE') or (NomTable = 'DUCSDETAIL'))) then Pendant582(NomTable);
  if ((VSOC < 595) and (NomTable = 'TRFDOSSIER')) then Pendant595;
  if ((VSOC < 602) and (NOMTABLE = 'JOURFERIE')) then Pendant602; //mcd 032003
  if ((VSOC < 614) and ((NOMTABLE = 'F2725_SOC' )OR (NOMTABLE = 'F2725_ENF'))) then MAJF2725Pendant(NomTable); // Fiscalité personnelle;
  if ((VSOC < 650  ) and (NOMTABLE = 'LISTEINVLIG') and (V_PGI.driver <> dbORACLE7)) then
    //Changement de clé Table LISTEINVLIG
    ExecuteSQL('UPDATE LISTEINVLIG SET GIL_GUID=PGIGUID WHERE GIL_GUID IS NULL');
  if ((VSOC < 674) and ((NOMTABLE = 'FPBIENS' )OR (NOMTABLE = 'FPBIENS_VE'))) then Pendant674(NomTable); // Fiscalité personnelle;
  {Maj PENDANT}
  if ((VSOC < 676) and (NOMTABLE = 'WPDRLIG' )and (V_PGI.driver <> dbORACLE7)) then ExecuteSQL('UPDATE WPDRLIG SET WPL_GUID=PGIGUID WHERE WPL_GUID IS NULL');
    //Changement de clé Table QMAGAPP
//  if ((VSOC < 720) and (NOMTABLE = 'QMAGAPP' )) then ExecuteSQL('UPDATE QMAGAPP SET QMP_GUID=PGIGUID WHERE QMP_GUID IS NULL');

  // XP 06.03.2006 Car ECRCOMPL est amenée par la 714
  if (VSOC < 715) and (NomTable = 'ECRCOMPL') then
    TraitementTableEcriture() ;

  if (VSOC < 740) and (NomTable = 'LIGNEFRAIS') then
  begin
     ExecuteSQL('UPDATE LIGNEFRAIS SET LF_RANG = 1+(SELECT Count(*) FROM LIGNEFRAIS L2 WHERE L2.LF_NATUREPIECEG=LIGNEFRAIS.LF_NATUREPIECEG AND'
       +' L2.LF_SOUCHE=LIGNEFRAIS.LF_SOUCHE AND L2.LF_NUMERO=LIGNEFRAIS.LF_NUMERO AND L2.LF_INDICEG=LIGNEFRAIS.LF_INDICEG AND '
       +' L2.LF_NUMORDRE=LIGNEFRAIS.LF_NUMORDRE AND L2.LF_FONCTIONNALITE||L2.LF_CODEPORT<LIGNEFRAIS.LF_FONCTIONNALITE||LIGNEFRAIS.LF_CODEPORT)');
  end;
End;

function MAJHalleyIsPossible(VSoc: Integer; MajLab1, MajLab2: TLabel; MajJauge: TProgressBar): boolean;
Begin
  Result:=True;
  if Vsoc < 540 then
  begin
        Result:=False ;
        PGIBox('La version de la base ('+intToStr(vsoc)+') est inférieure à 540. ' + #13#10
             + 'Vous devez au préalable faire une mise à jour intermédiaire en structure 540 .','Attention');
  end;
  If (V_PGI.Driver=dbINTRBASE) then
  begin
        Result:=False ;
        PGIBox('Le SGBDR Interbase n''est plus supporté à partir de cette version ','Attention');
  end;
  // Pas possible si base CHR, BTP ou Mode
  // or (GetParamSoc('SO_DISTRIBUTION')='016') Pas CHR à partir de 699

  if ((GetParamSoc('SO_DISTRIBUTION')='014')  or (GetParamSoc('SO_DISTRIBUTION')='030') ) then
  begin
        Result:=False ;
        PGIBox('Vous ne pouvez pas traiter une base sectorielle ' + #13#10
             + '( Code distribution '+ GetParamSoc('SO_DISTRIBUTION')+ ' )'+ #13#10
             + 'avec ce programme. ','Attention');
  end;
End;

end.
