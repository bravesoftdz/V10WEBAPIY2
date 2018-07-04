unit MajHalleyProc;

interface

procedure Recup_UserPrefGc;
procedure MajVer_DP_530;
procedure MHCopyPlanRefToGenerauxRef;
procedure MoveChoixcodeFamRubToChoixExt;
procedure RTMoveTypeAction;
procedure RecupTabletteChoixCodeGCTYPEMASQUE;
procedure InstalleLesCoefficientsDegressifs;
procedure MajJurPredef;
procedure MajAnnulienLoiNRE;
procedure MajApresForTheTox;
procedure MajApresForTheTox2;
procedure RT_InsertLibelleInfoComplAction;
procedure CorrectionSiren;
procedure MajPays;
procedure RTMoveInfosData;
procedure RTMoveConcurrent;
procedure GIGAChangeUnite;
procedure GIGAChangeNomAffaire;
procedure RT_InsertLibelleInfoComplOperationetChain;
(*
procedure MajTableDureeTaux; // G.Piot
*)
procedure MajJuridiqueTypePer; // jerome Lefevre Paie 12/2002
procedure GIGAIndexActivite; //mcd 032003
Procedure MajGPAOAvant;
Procedure MajGPAOApres;
Procedure MajGPAOAvant612;
Procedure MajGPAOApres612;
Procedure CreationEnrTYPEREMISE;
Procedure PGMinConvPaie;
Procedure InitTablePaie605;
procedure ConvertirDateHeureEvenement;
procedure MAJF2725Pendant(NomTable:string);
Procedure MajImoAvantPourSynchro ;
procedure RTMajDureeAction;
Procedure Init_WInitChampLig;
Procedure GPAOMajApres_617;
procedure CorrigeRTTiers ;
procedure MajFluxTreso;
procedure RT_InsertLibelleInfoComplContact;
procedure W_SAV_COMPETENCE;
procedure W_SAV_COMPL;
Procedure GIGAAvant644;
Procedure GIGAApres644;
procedure GPMajTypeTarif;
procedure GPMajSTKNature;
procedure GPMajSTKValoParam;
Procedure MajImo646;
procedure RT_InsertLibelleTablettesFour;
procedure MajEvtPCP_V654;

function wExistGA(Article: string; WithAlert: Boolean = false; Where : string = ''): Boolean;
procedure GIGAAvant658;
procedure W_SAV_COMPETENCE_UPDATE;
procedure W_SAV_INTER_COMPL;
procedure GPMajPdrType;
Procedure GPMajxWizards;
procedure RT_InsertLibelleInfoComplArticle;
procedure RT_InsertLibelleInfoComplLigne;
procedure CFE5_RegulLiacodedo;
procedure CPMajTLVersCHOIXCOD;
procedure UpdateNatureSuivante;
procedure GPCopyWPRINTTV;
procedure RT_InsertInfoCompl;
procedure W_SAV_ArticlesParc;
procedure RT_InsertLibelleInfoComplDixArtLigne;
procedure YY_InitChampsSystemes;
procedure MaJdesPrixGSMFromBlocNote;
{Passage de la V6.0 à V7.0}
//Mise à jour QSIMULATION en fonction de WCBNEVOLUTION et WPARAM
procedure RecupParamCBN;
procedure RT_InsertTablettesSAV;
Procedure GPAOMajMES;
procedure Maj_GSM_GUIDORI;
procedure SuppAnnuInterloc;
procedure MajMarkMESApres;
procedure RT_InsertTablettesWIV;
procedure GCTraite_TIERSIMPACTPIECE;
procedure InsertTablettesRQ; //kb270706
procedure wConsolideCalculParam;
function UpdateAcpteCptaDiff2 : boolean ;
procedure PgMajChampSWS ();
procedure MajEngagement;
procedure UpdateColRTYPEIDBQ;
procedure CPActivationPointageAvance ;
procedure CPInitialisationEtablissementEtAgencesBancaires ;


implementation

uses HEnt1, HCtrls, MajTable, StdCtrls, ComCtrls, DB,
{$IFDEF VER150} {D7}
  Variants,
{$ENDIF}
  {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
  HQry, SysUtils, Classes,
  DBCtrls,
  TiersUtil,      //mcd 12/2005
  MajHalleyUtil, UTOB, UtilPgi, ParamSoc, Forms, HStatus, MajStruc,CbpTrace,
  UtilRessource, Controls, Math, wcommuns,  HMsgBox,UtilOutils,CPOBJITRTPARAM,CPOBJBANQUESH,uInitBanque ;

var  TobTacheJour,TobAfPlanningEtat : TOB ;

procedure PgMajChampSWS ();
var st, LePrefixe : String;
TS : TOB;
i : Integer;
begin
  St := 'SELECT DH_PREFIXE,DT_NOMTABLE FROM DECHAMPS LEFT JOIN DETABLES ON DT_PREFIXE=DH_PREFIXE WHERE DH_NOMCHAMP like "P%_PGSWS"';
  TS := TOB.Create('LES CHAMPS SWS', nil, -1);
  TS.LoadDetailDBFROMSQL('DECHAMPS', st);
  for i := 0 to TS.detail.count - 1 do
  begin
    LePrefixe := TS.detail[i].getvalue('DH_PREFIXE');
    St := 'UPDATE LATABLE SET LEPREF_PGSWS="000" WHERE LEPREF_PREDEFINI="CEG" AND (LEPREF_PGSWS="" OR LEPREF_PGSWS IS NULL)';
    St := StringReplace (St, 'LATABLE', TS.detail[i].getvalue('DT_NOMTABLE'), [rfIgnoreCase,rfReplaceAll]);
    St := StringReplace (St, 'LEPREF', LePRefixe, [rfIgnoreCase,rfReplaceAll]);
    ExecuteSQL (st);
    St := StringReplace (St, '"000"', '"001"', [rfIgnoreCase,rfReplaceAll]);
    St := StringReplace (St, '"CEG"', '"DOS"', [rfIgnoreCase,rfReplaceAll]);
    ExecuteSQL (st);
    St := StringReplace (St, '"001"', '"051"', [rfIgnoreCase,rfReplaceAll]);
    St := StringReplace (St, '"DOS"', '"STD"', [rfIgnoreCase,rfReplaceAll]);
    ExecuteSQL (st);
  end;
  FreeAndNil (TS);
end;


procedure Recup_UserPrefGc;
var QQ: TQuery;
  iNature, pos: integer;
  req, sql, stVal, stChamp: string;
begin
  if not TableExiste('USERPREFGC') then exit;
  QQ := OpenSQL('select GUP_UTILISATEUR,GUP_ARTDIMDETAIL,GUP_ARTDIMDEFAUT,GUP_ARTDIMVAL,GUP_STODIMDEFAUT,GUP_STODIMVAL from USERPREFGC', True);
  while not QQ.EOF do
  begin
    for iNature := 0 to 1 do
    begin
      req := 'insert into USERPREFDIM (GUD_NATUREPIECE,GUD_UTILISATEUR,GUD_DETAIL,GUD_DEFAUT,' +
        'GUD_POSITION,GUD_CHAMP,GUD_LIBELLE) values (';
      if iNature = 0 then
      begin
        req := req + '"art","' + QQ.Findfield('GUP_UTILISATEUR').AsString + '","' + QQ.Findfield('GUP_ARTDIMDETAIL').AsString + '","' +
          QQ.Findfield('GUP_ARTDIMDEFAUT').AsString + '",';
        stVal := QQ.Findfield('GUP_ARTDIMVAL').AsString;
      end else
      begin
        req := req + '"sto","' + QQ.Findfield('GUP_UTILISATEUR').AsString + '","","' + QQ.Findfield('GUP_STODIMDEFAUT').AsString + '",';
        stVal := QQ.Findfield('GUP_STODIMVAL').AsString;
      end;
      pos := 0;
      while stVal <> '' do
      begin
        stChamp := Trim(ReadTokenSt(stVal));
        if stChamp <> '' then
        begin
          inc(pos);
          sql := req + IntToStr(pos) + ',"' + stChamp + '","' + RechDom('GCDIMCHAMP', stChamp, False) + '")';
          ExecuteSQL(sql);
        end;
      end;
    end; // for iNature:=0 to 1 do
    QQ.next;
  end; // while not QQ.EOF do
  Ferme(QQ);
  //PGIBox('Récupération ok','USERPREFGC') ;
end;

procedure MajVer_DP_530;
var
  sql: string;
begin
   // if not TableSurAutreBase('JURIDIQUE') then  by PCS and MD 09112004 DESHARE désactivé dans MajStruct
   if IsMonoOuCommune then
   begin
    // normalement, ANN_CAPNBTITRE bien que Double contient uniqumt des Entier
    // sinon, le résultat se trouvera arrondi dans JUR_NBTITRESCLOT (entier)
    sql := 'UPDATE JURIDIQUE SET JUR_NBTITRESCLOT=ANN_CAPNBTITRE '
      + 'FROM JURIDIQUE, ANNUAIRE WHERE JURIDIQUE.JUR_CODEPERDOS=ANNUAIRE.ANN_CODEPER ';
    try
      ExecuteSQL(sql);
    except;
    end;
    sql := 'UPDATE JURIDIQUE SET JUR_VALNOMINCLOT=ANN_CAPVN '
      + 'FROM JURIDIQUE, ANNUAIRE WHERE JURIDIQUE.JUR_CODEPERDOS=ANNUAIRE.ANN_CODEPER ';
    try
      ExecuteSQL(sql);
    except;
    end;
    // normalement ANN_CAPNBTITRE est entier
    sql := 'UPDATE JURIDIQUE SET JUR_NBDROITSVOTE=ANN_CAPNBTITRE '
      + 'FROM JURIDIQUE, ANNUAIRE WHERE JURIDIQUE.JUR_CODEPERDOS=ANNUAIRE.ANN_CODEPER ';
    try
      ExecuteSQL(sql);
    except;
    end;
  end;
   // if not TableSurAutreBase('ANNULIEN') then by PCS and MD 09112004 DESHARE désactivé dans MajStruct
   if IsMonoOuCommune then
  begin

    // attention, quand un des champs contient Null, la somme est Null
    try
      ExecuteSQL('UPDATE ANNULIEN SET ANL_TTNBTOTUS=ANL_TTNBPP+ANL_TTNBUS');
    except
    end;
    try
      ExecuteSQL('UPDATE ANNULIEN SET ANL_TTNBTOT=ANL_TTNBPP+ANL_TTNBNP');
    except
    end;
    // 1.0 obligé !!! pour convertir en rate avant division
    // (sinon le ratio d'entiers reste entier donc perte des décimales)
    sql := 'UPDATE ANNULIEN SET ANL_TTPCTBENEF=1.0 * (ANL_TTNBTOTUS * 10000/JUR_NBTITRESCLOT) / 10000 '
      + 'FROM JURIDIQUE, ANNULIEN WHERE JUR_CODEPERDOS=ANL_CODEPERDOS AND JUR_NBTITRESCLOT<>0';
    try
      ExecuteSQL(sql);
    except;
    end;
    // idem
    sql := 'UPDATE ANNULIEN SET ANL_TTPCTCAP=1.0 *  (ANL_TTNBTOT * 10000/JUR_NBTITRESCLOT ) / 10000 '
      + 'FROM JURIDIQUE, ANNULIEN WHERE JUR_CODEPERDOS=ANL_CODEPERDOS AND JURIDIQUE.JUR_NBTITRESCLOT<>0';
    try
      ExecuteSQL(sql);
    except;
    end;
    // idem rate := 1.0 * double / entier
    sql := 'UPDATE ANNULIEN SET ANL_TTPCTVOIX=1.0 *  (ANL_VOIXAGE * 10000/JUR_NBDROITSVOTE) / 10000 '
      + 'FROM JURIDIQUE, ANNULIEN WHERE JUR_CODEPERDOS=ANL_CODEPERDOS AND JUR_NBDROITSVOTE<>0 AND ANNULIEN.ANL_VOIXAGE<>0';
    try
      ExecuteSQL(sql);
    except;
    end;
  end;
end;

procedure MHCopyPlanRefToGenerauxRef;
var QEcr, QLu: TQuery;
begin
  QLu := OpenSQL('SELECT * FROM PLANREF WHERE PR_PREDEFINI="STD"', True);
  QEcr := OpenSQL('SELECT * FROM GENERAUXREF WHERE GER_PREDEFINI="STD"', False);
  if QEcr.eof then
  begin
    while not QLu.EOF do
    begin
      QEcr.Insert;
      InitNew(QEcr);
      QEcr.FindField('GER_NUMPLAN').AsInteger := QLu.FindField('PR_NUMPLAN').AsInteger;
      QEcr.FindField('GER_GENERAL').AsString := QLu.FindField('PR_COMPTE').AsString;
      QEcr.FindField('GER_LIBELLE').AsString := QLu.FindField('PR_LIBELLE').AsString;
      QEcr.FindField('GER_ABREGE').AsString := QLu.FindField('PR_ABREGE').AsString;
      QEcr.FindField('GER_NATUREGENE').AsString := QLu.FindField('PR_NATUREGENE').AsString;
      QEcr.FindField('GER_CENTRALISABLE').AsString := QLu.FindField('PR_CENTRALISABLE').AsString;
      QEcr.FindField('GER_SOLDEPROGRESSIF').AsString := QLu.FindField('PR_SOLDEPROGRESSIF').AsString;
      QEcr.FindField('GER_SAUTPAGE').AsString := QLu.FindField('PR_SAUTPAGE').AsString;
      QEcr.FindField('GER_TOTAUXMENSUELS').AsString := QLu.FindField('PR_TOTAUXMENSUELS').AsString;
      QEcr.FindField('GER_COLLECTIF').AsString := QLu.FindField('PR_COLLECTIF').AsString;
      QEcr.FindField('GER_SENS').AsString := QLu.FindField('PR_SENS').AsString;
      QEcr.FindField('GER_LETTRABLE').AsString := QLu.FindField('PR_LETTRABLE').AsString;
      QEcr.FindField('GER_POINTABLE').AsString := QLu.FindField('PR_POINTABLE').AsString;
      QEcr.FindField('GER_VENTILABLE1').AsString := QLu.FindField('PR_VENTILABLE1').AsString;
      QEcr.FindField('GER_VENTILABLE2').AsString := QLu.FindField('PR_VENTILABLE2').AsString;
      QEcr.FindField('GER_VENTILABLE3').AsString := QLu.FindField('PR_VENTILABLE3').AsString;
      QEcr.FindField('GER_VENTILABLE4').AsString := QLu.FindField('PR_VENTILABLE4').AsString;
      QEcr.FindField('GER_VENTILABLE5').AsString := QLu.FindField('PR_VENTILABLE5').AsString;
      QEcr.FindField('GER_PREDEFINI').AsString := QLu.FindField('PR_PREDEFINI').AsString;
      QEcr.Post;
      QLu.Next;
    end;
  end;
  Ferme(QLu);
  Ferme(QEcr);
end;

procedure MoveChoixcodeFamRubToChoixExt;
var QCC, QYX: TQuery;
  CcTyp, YxTyp: string;
begin
  if (V_PGI.ModePCL = '1') then Exit;
  CcTyp := 'RBF';
  YxTyp := 'RBF';
  QCC := OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="' + CcTyp + '"', True);
  while not QCC.EOF do
  begin
    QYX := OpenSQL('SELECT * FROM CHOIXDPSTD WHERE YDS_TYPE="' + YxTyp + '" AND YDS_CODE="' + QCC.FindField('CC_CODE').AsString + '" ', False);
    if QYX.EOF then
    begin
      QYX.Insert;
      InitNew(QYX);
      QYX.FindField('YDS_TYPE').AsString := YxTyp;
      QYX.FindField('YDS_CODE').AsString := QCC.FindField('CC_CODE').AsString;
      QYX.FindField('YDS_LIBELLE').AsString := QCC.FindField('CC_LIBELLE').AsString;
      QYX.FindField('YDS_ABREGE').AsString := QCC.FindField('CC_ABREGE').AsString;
      QYX.FindField('YDS_LIBRE').AsString := '';
      QYX.FindField('YDS_PREDEFINI').AsString := 'DOS';
      QYX.FindField('YDS_NODOSSIER').AsString := V_PGI.NoDossier;
    end;
    Ferme(QYX);
    QCC.Next;
  end;
  Ferme(QCC);
end;

procedure RTMoveTypeAction;
var QCC, QYX: TQuery;
  CcTyp, YxTyp: string;
begin
  CcTyp := 'RTA';
  QCC := OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="' + CcTyp + '"', True);
  while not QCC.EOF do
  begin
    YxTyp := QCC.FindField('CC_CODE').AsString;
    QYX := OpenSQL('SELECT * FROM PARACTIONS WHERE RPA_TYPEACTION="' + YxTyp + '"', False);
    if QYX.EOF then
    begin
      QYX.Insert;
      InitNew(QYX);
      QYX.FindField('RPA_TYPEACTION').AsString := YxTyp;
      QYX.FindField('RPA_LIBELLE').AsString := QCC.FindField('CC_LIBELLE').AsString;
      QYX.Post;
    end;
    Ferme(QYX);
    QCC.Next;
  end;
  Ferme(QCC);
end;

procedure RecupTabletteChoixCodeGCTYPEMASQUE;
var QCC, QTM: TQuery;
  ccCod: string;
begin
  if ExisteSQL('select GMQ_TYPEMASQUE from TYPEMASQUE') then exit;
  QCC := OpenSQL('select * from CHOIXCOD where CC_TYPE="GTQ"', True);
  while not QCC.EOF do
  begin
    ccCod := QCC.FindField('CC_CODE').AsString;
    QTM := OpenSQL('select * from TYPEMASQUE where GMQ_TYPEMASQUE="' + ccCod + '"', False);
    if QTM.EOF then
    begin
      QTM.Insert;
      InitNew(QTM);
      QTM.FindField('GMQ_TYPEMASQUE').AsString := ccCod;
      QTM.FindField('GMQ_LIBELLE').AsString := QCC.FindField('CC_LIBELLE').AsString;
      QTM.FindField('GMQ_ABREGE').AsString := QCC.FindField('CC_ABREGE').AsString;
      QTM.FindField('GMQ_MULTIETAB').AsString := CheckToString(ExisteSql('select GDM_MASQUE from DIMMASQUE where GDM_TYPEMASQUE="' + ccCod +
        '" and GDM_POSITION6<>""'));
      QTM.Post;
    end;
    Ferme(QTM);
    QCC.Next;
  end;
  Ferme(QCC);
end;

procedure InstalleLesCoefficientsDegressifs;
begin
  ExecuteSQL('DELETE FROM CHOIXCOD WHERE CC_TYPE="ICD"');
  ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) VALUES ("ICD","001","01/01/1900;1.5;2;2.5","","")');
  ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) VALUES ("ICD","002","01/02/1996;2.5;3;3.5","","")');
  ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) VALUES ("ICD","003","01/02/1997;1.5;2;2.5","","")');
  ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) VALUES ("ICD","004","01/01/2001;1.25;1.75;2.25","","")');
end;

procedure MajJurPredef;
var
  lstTable, laTable, rac: string;
begin
  //Tables à mettre à jour
  lstTable := 'JUBIBMODULE;JUBIBMODULE2;JUBIBACTION;JUBIBACTION2;JUBIBRUBRIQUE;JUBIBRUBRIQUE2';
  while lstTable <> '' do
  begin
    LaTable := ReadTokenSt(lstTable);
    if TableExiste(LaTable) then
    begin
      rac := TableToPrefixe(LaTable);
      if not ChampPhysiqueExiste(LaTable, rac + '_PREDEFMAJ') then
      begin
        //On ajoute le champ de substitution
        AddChamp(LaTable, rac + '_PREDEFMAJ', 'COMBO', rac, False, 'CEG', 'MAJ bibles', '', False);
        //Récupération des données contenues dans le XXX_PREDEFINI
        ExecuteSQL('UPDATE ' + LaTable + ' SET ' + Rac + '_PREDEFMAJ=' + Rac + '_PREDEFINI');
        //La suppression du champ XXX_PREDEFINI sera fait par mise à jour de structure.
        //On fournit les tables sans ce champ, mais avec le nouveau( précédement renseigné et créé).
      end;
    end;
  end;
end;

procedure MajAnnulienLoiNRE;
// Juridique et DP (MD/JB)
begin

  if GetParamSoc('SO_JUDOSNOMIL') = False then
  begin
    BeginTrans;

    // Transformer tous les liens DGR (ancienne fonction) non déjà traités, en liens DGD (nouvelle fonction)
    ExecuteSQL('update ANNULIEN set ANL_DATEMODIF="' + UsDateTime(Date) + '", ANL_APPMODIF="JURI NRE", '
      + 'ANL_FONCTION="DGD" where (ANL_FORME="SA" or ANL_FORME="SELAFA") and ANL_FONCTION="DGR" and '
      + 'ANL_APPMODIF<>"JURI NRE"');

    // Créer un lien DGA (pour tous les enregistrements de type PCA non déjà traités)
    ExecuteSQL('insert into ANNULIEN '
      + '(ANL_CODEPER, ANL_CODEPERDOS, ANL_TYPEDOS, ANL_NOORDRE, ANL_CODEDOS, ANL_DATECREATION, '
      + 'ANL_DATEMODIF, ANL_UTILISATEUR, ANL_APPMODIF, ANL_FONCTION, ANL_FORME, ANL_NOMPER, '
      + 'ANL_RACINE, ANL_TIERS, ANL_INFO, ANL_AFFICHE, ANL_TRI, ANL_EFDEB, ANL_EFFIN, ANL_TTNBPP, '
      + 'ANL_TTNBUS, ANL_TTNBNP, ANL_TTNBTOT, ANL_TTNBTOTUS, ANL_TTMONTANT, ANL_TTPCTCAP, '
      + 'ANL_TTPCTBENEF, ANL_TTPCTVOIX, ANL_VOIXAGO, ANL_VOIXAGE, ANL_INDIVIS,  '
      + 'ANL_PERASS1CODE, ANL_PERASS2CODE, ANL_PERASS3CODE, ANL_CONV, '
      + 'ANL_CONVDATE, ANL_CONVDATEAPP, ANL_TRAVCONT,   '
      + 'ANL_TXDETDIRECT, ANL_TXDETINDIRECT, '
      + 'ANL_TXDETTOTAL, ANL_GRPFISCAL,  ANL_BANQTXDEB, ANL_BANQDECOUVERT, '
      + 'ANL_NOMDATE, ANL_RENDATE, ANL_EXPDATE, ANL_EXPEXE, ANL_COOPTPROV, ANL_COOPTCODE, '
      + 'ANL_MDEXP, ANL_MDRN, ANL_MDRP, ANL_MDNRNR, ANL_MDRPNO, ANL_MDRPCODE, ANL_CRSPNO, '
      + 'ANL_CRSPCODE, ANL_CJASS, ANL_APPNAT, ANL_APPNATM, ANL_APPNATT, ANL_APPNUMM, ANL_APPNUMT, '
      + 'ANL_APPNUML) '
      + 'select ANL_CODEPER, ANL_CODEPERDOS, ANL_TYPEDOS, ANL_NOORDRE, ANL_CODEDOS, "' + UsDateTime(Date) + '", '
      + '"' + UsDateTime(Date) + '", ANL_UTILISATEUR, "JURI NRE", "DGA", ANL_FORME, ANL_NOMPER, "DGA", '
      + 'ANL_TIERS, ANL_INFO, "DIR GEN", 42, ANL_EFDEB, ANL_EFFIN, ANL_TTNBPP, ANL_TTNBUS, '
      + 'ANL_TTNBNP, ANL_TTNBTOT, ANL_TTNBTOTUS, ANL_TTMONTANT, ANL_TTPCTCAP, ANL_TTPCTBENEF, '
      + 'ANL_TTPCTVOIX, ANL_VOIXAGO, ANL_VOIXAGE, ANL_INDIVIS,  '
      + 'ANL_PERASS1CODE, ANL_PERASS2CODE, ANL_PERASS3CODE, ANL_CONV, ANL_CONVDATE, '
      + 'ANL_CONVDATEAPP, ANL_TRAVCONT, '
      + 'ANL_TXDETDIRECT, ANL_TXDETINDIRECT, ANL_TXDETTOTAL, '
      + 'ANL_GRPFISCAL,  ANL_BANQTXDEB, ANL_BANQDECOUVERT, ANL_NOMDATE, ANL_RENDATE, '
      + 'ANL_EXPDATE, ANL_EXPEXE, ANL_COOPTPROV, ANL_COOPTCODE, ANL_MDEXP, ANL_MDRN, ANL_MDRP, '
      + 'ANL_MDNRNR, ANL_MDRPNO, ANL_MDRPCODE, ANL_CRSPNO, ANL_CRSPCODE, ANL_CJASS, ANL_APPNAT, '
      + 'ANL_APPNATM, ANL_APPNATT, ANL_APPNUMM, ANL_APPNUMT, ANL_APPNUML from ANNULIEN '
      + 'where ANL_FONCTION="PCA" and ANL_APPMODIF<>"JURI NRE"');

    // Annoter les liens PCA pour marquer que la modif. a déjà été faite (pour ne pas la refaire lors de la prochaine MAJ SocRef)
    ExecuteSQL('update ANNULIEN set ANL_DATEMODIF="' + UsDateTime(Date) + '", ANL_APPMODIF="JURI NRE" '
      + 'where ANL_FONCTION="PCA" and ANL_APPMODIF<>"JURI NRE"');
    SetParamSoc('SO_JUDOSNOMIL', True);
    CommitTrans;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : DMD
Créé le ...... : 12/12/2001
Modifié le ... :   /  /
Description .. : Cette procédure est appelée par la mise à jour de structure.
Suite ........ : Elle permet de transférer certaines données
Suite ........ : (SSI_CARRIVE,SSI_CREJET,SSI_CTRAITE,SSI_CENVOY
Suite ........ : E,SSI_CDEPART) dans respectivement
Suite ........ : (SSI_CARRIVEFTP,SSI_CREJETFTP,SSI_CTRAITEFTP,S
Suite ........ : SI_CENVOYEFTP,SSI_CDEPARTFTP) pour tous les sites
Suite ........ : distants.
Suite ........ : Ces mêmes champs sont ensuite remis à blanc.
Mots clefs ... :
*****************************************************************}

procedure MajApresForTheTox;
begin
  try
    BeginTrans;

    { Contrôle de la présence des nouveaux champs dans V_PGI.DECHAMPS par le contrôle d'un seul : SSI_CARRIVEFTP }
    if ChampToNum('SSI_CARRIVEFTP') > 0 then
    begin
      { Uniquement dans le cas des sites distants SSI_STYPESITE='001' }
      { Prendre les valeurs de SSI_CARRIVE => SSI_CARRIVEFTP }
      { Prendre les valeurs de SSI_CTRAITE => SSI_CTRAITEFTP }
      { Prendre les valeurs de SSI_CREJET  => SSI_CREJETFTP }
      { Prendre les valeurs de SSI_CDEPART => SSI_CDEPARTFTP }
      { Prendre les valeurs de SSI_CENVOYE => SSI_CENVOYEFTP }
      ExecuteSql('UPDATE STOXSITES SET SSI_CARRIVEFTP=SSI_CARRIVE,SSI_CDEPARTFTP=SSI_CDEPART,SSI_CREJETFTP=SSI_CREJET,SSI_CENVOYEFTP=SSI_CENVOYE,SSI_CTRAITEFTP=SSI_CTRAITE WHERE SSI_STYPESITE="001"');
      { Remettre à "" les champs SSI_CARRIVE,SSI_CDEPART,SSI_CREJET,SSI_CTRAITE,SSI_CENVOYE pour les sites distants }
      ExecuteSql('UPDATE STOXSITES SET SSI_CARRIVE="",SSI_CDEPART="",SSI_CREJET="",SSI_CENVOYE="",SSI_CTRAITE="" WHERE SSI_STYPESITE="001"');
    end
    else
    begin
      { P'tain ou qu'il est mon champ ?? }
    end;

    CommitTrans;
  except
    on E: Exception do
    begin
      RollBack;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : DMD
Créé le ...... : 07/02/2002
Modifié le ... :   /  /
Description .. : Cette procédure est appelée par la mise à jour de structure.
Suite ........ : Elle permet de renseigner un nouveau champ SSI_PORTFTP à 21.
Suite ........ : Version de la table STOXSITES 15
Mots clefs ... :
*****************************************************************}

procedure MajApresForTheTox2;
begin
  try
    BeginTrans;

    { Contrôle de la présence du nouveau champ dans V_PGI.DECHAMPS }
    if ChampToNum('SSI_PORTFTP') > 0 then
    begin
      ExecuteSql('UPDATE STOXSITES SET SSI_PORTFTP=21');
    end
    else
    begin
      { Z'auriez pas vu mon champ ??? }
    end;

    CommitTrans;
  except
    on E: Exception do
    begin
      RollBack;
    end;
  end;
end;

procedure RT_InsertLibelleInfoComplAction;
var i: integer;
begin
  for i := 0 to 4 do InsertChoixCode('R1Z', 'BL' + intTostr(i), 'Décision libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R1Z', 'DL' + intTostr(i), 'Date libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R1Z', 'CL' + intTostr(i), 'Table libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R1Z', 'TL' + intTostr(i), 'Texte libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R1Z', 'VL' + intTostr(i), 'Valeur libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R1Z', 'ML' + intTostr(i), 'Multi-choix libre ' + intTostr(i + 1), '', '');
end;

procedure RT_InsertLibelleInfoComplOperationetChain;
var i: integer;
begin
  for i := 0 to 4 do InsertChoixCode('R2Z', 'BL' + intTostr(i), 'Décision libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R2Z', 'DL' + intTostr(i), 'Date libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R2Z', 'CL' + intTostr(i), 'Table libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R2Z', 'TL' + intTostr(i), 'Texte libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R2Z', 'VL' + intTostr(i), 'Valeur libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R2Z', 'ML' + intTostr(i), 'Multi-choix libre ' + intTostr(i + 1), '', '');
  InsertChoixCode('R2Z', 'BLO', 'Bloc-notes ', '', '');
  for i := 1 to 3 do InsertChoixCode('RCH', 'RH' + intTostr(i), 'Table libre ' + intTostr(i + 1), '', '');
end;

procedure CorrectionSiren;
var
  TOBAnnuaire_l: TOB;
  QRYRequete_l: TQuery;
  nTOBInd_l, nCarInd_l: integer;
  sAnnSiren_l, sNewSiren_l: string;
begin
  // seul cas où on ne traite pas : si mode MULTI et qu'on n'est pas sur la base commune
   // if TableSurAutreBase('ANNUAIRE') then Exit;  by PCS and MD 09112004 DESHARE désactivé dans MajStruct
   if not IsMonoOuCommune then Exit;

  BeginTrans;
  try
    TOBAnnuaire_l := TOB.Create('ANNUAIRE', nil, -1);
    QRYRequete_l := OpenSQL('select ANN_CODEPER, ANN_SIREN from ANNUAIRE where ANN_SIREN <> "" and ANN_SIREN like "% %"', true);
    TOBAnnuaire_l.LoadDetailDB('ANNUAIRE', '', '', QRYRequete_l, false);
    Ferme(QRYRequete_l);
    InitMove(TOBAnnuaire_l.Detail.Count, 'Correction du SIREN');

    for nTOBInd_l := 0 to TOBAnnuaire_l.Detail.Count - 1 do
    begin
      sAnnSiren_l := TOBAnnuaire_l.Detail[nTOBInd_l].GetValue('ANN_SIREN');
      sNewSiren_l := '';
      nCarInd_l := 1;
      while (nCarInd_l < Length(sAnnSiren_l) + 1) do
      begin
        if (sAnnSiren_l[nCarInd_l] <> ' ') then
          sNewSiren_l := sNewSiren_l + sAnnSiren_l[nCarInd_l];
        inc(nCarInd_l);
      end;
      TOBAnnuaire_l.Detail[nTOBInd_l].PutValue('ANN_SIREN', sNewSiren_l);
      MoveCur(false);
    end;
    FiniMove;
    TOBAnnuaire_l.UpdateDB(true);
    TOBAnnuaire_l.Free;
  except
    V_PGI.IoError := oeUnknown;
  end;

  if V_PGI.IoError = oeUnknown then
    RollBack
  else
    CommitTrans;
end;

procedure MajPays;
var TSoc, TRef: THTable;
  PaysRef: string;
begin
  TSoc := THTable.Create(Application);
  TSoc.DatabaseName := DBSoc.DatabaseName;
  {Copie les drapeaux de la table PAYS}
  TSoc.Tablename := 'PAYS';
  TSoc.Indexname := 'PY_CLE1';
  TSoc.Open;

  TRef := OpenTableREF('PAYS', 'PY_CLE1');
  while not TRef.EOF do
  begin
    PaysRef := TRef.FindField('PY_PAYS').AsString;
    if PaysRef > '' then
    begin
      if TSoc.Findkey([PaysRef]) then
      begin
        TSoc.FindField('PY_DRAPEAUX').Value := TRef.FindField('PY_DRAPEAUX').Value;
        TSoc.Post;
      end;
    end;
    TRef.Next;
  end;
  Ferme(TRef);
  TSoc.Close;
  TSoc.Free;
end;

procedure RTMoveConcurrent;
var Q: TQuery;
  StListe, stConc, Sql: string;
begin
  BeginTrans;
  Q := OpenSQL('SELECT RPE_CONCURRENTS,rpe_perspective,rpe_tiers FROM PERSPECTIVES where rpe_concurrents <> ""', True);
  while not Q.EOF do
  begin
    StListe := Q.FindField('RPE_CONCURRENTS').AsString;
    if StListe <> '' then
      repeat
        StConc := Trim(ReadTokenSt(StListe));
        if StConc <> '' then
        begin
          Sql := 'INSERT INTO PERSPECTIVESTIERS (RPT_TIERS,RPT_PERSPECTIVE) VALUES ("';
          Sql := Sql + StConc + '",' + IntToStr(Q.FindField('rpe_perspective').AsInteger) + ')';
          ExecuteSQL(Sql);
        end;
      until StConc = '';
    Q.Next;
  end;
  Ferme(Q);
  ExecuteSql('Update perspectives set rpe_concurrents="" where rpe_concurrents <> ""');
  // SetParamSoc ('SO_RTDEPLACECON',True) ;
  CommitTrans;
end;

procedure RTMoveInfosData;
var SQL: string;
begin
  if (v_pgi.driver in [dbORACLE7, dbORACLE8]) then Exit;
  // cette requete ne marche pas sous Oracle à cause des BLOBs
  //  mais aucun clients concerné sous Oracle
  BeginTrans;
  if TableExiste('RTINFOSDATA') then
  begin
    SQL := 'INSERT INTO RTINFOS001 (';
    SQL := SQL + 'RD1_CLEDATA,RD1_DATECREATION,RD1_DATEMODIF,';
    SQL := SQL + 'RD1_CREATEUR,RD1_UTILISATEUR,RD1_BLOCNOTE,';
    SQL := SQL + 'RD1_RD1LIBTEXTE0,RD1_RD1LIBTEXTE1,RD1_RD1LIBTEXTE2,RD1_RD1LIBTEXTE3,RD1_RD1LIBTEXTE4,';
    SQL := SQL + 'RD1_RD1LIBTABLE0,RD1_RD1LIBTABLE1,RD1_RD1LIBTABLE2,RD1_RD1LIBTABLE3,RD1_RD1LIBTABLE4,';
    SQL := SQL + 'RD1_RD1LIBVAL0,RD1_RD1LIBVAL1,RD1_RD1LIBVAL2,RD1_RD1LIBVAL3,RD1_RD1LIBVAL4,';
    SQL := SQL + 'RD1_RD1LIBMUL0,RD1_RD1LIBMUL1,RD1_RD1LIBMUL2,RD1_RD1LIBMUL3,RD1_RD1LIBMUL4,';
    SQL := SQL + 'RD1_RD1LIBBOOL0,RD1_RD1LIBBOOL1,RD1_RD1LIBBOOL2,RD1_RD1LIBBOOL3,RD1_RD1LIBBOOL4,';
    SQL := SQL + 'RD1_RD1LIBDATE0,RD1_RD1LIBDATE1,RD1_RD1LIBDATE2,RD1_RD1LIBDATE3,RD1_RD1LIBDATE4)';

    SQL := SQL + ' SELECT ';
    SQL := SQL + 'RDA_CLEDATA,RDA_DATECREATION,RDA_DATEMODIF,';
    SQL := SQL + 'RDA_CREATEUR,RDA_UTILISATEUR,RDA_BLOCNOTE,';
    SQL := SQL + 'RDA_RPRLIBTEXTE0,RDA_RPRLIBTEXTE1,RDA_RPRLIBTEXTE2,RDA_RPRLIBTEXTE3,RDA_RPRLIBTEXTE4,';
    SQL := SQL + 'RDA_RPRLIBTABLE0,RDA_RPRLIBTABLE1,RDA_RPRLIBTABLE2,RDA_RPRLIBTABLE3,RDA_RPRLIBTABLE4,';
    SQL := SQL + 'RDA_RPRLIBVAL0,RDA_RPRLIBVAL1,RDA_RPRLIBVAL2,RDA_RPRLIBVAL3,RDA_RPRLIBVAL4,';
    SQL := SQL + 'RDA_RPRLIBMUL0,RDA_RPRLIBMUL1,RDA_RPRLIBMUL2,RDA_RPRLIBMUL3,RDA_RPRLIBMUL4,';
    SQL := SQL + 'RDA_RPRLIBBOOL0,RDA_RPRLIBBOOL1,RDA_RPRLIBBOOL2,RDA_RPRLIBBOOL3,RDA_RPRLIBBOOL4,';
    SQL := SQL + 'RDA_RPRLIBDATE0,RDA_RPRLIBDATE1,RDA_RPRLIBDATE2,RDA_RPRLIBDATE3,RDA_RPRLIBDATE4';
    SQL := SQL + ' FROM RTINFOSDATA WHERE RDA_DESC="1"';

    ExecuteSQL(SQL);
    ExecuteSQL('DELETE from RTINFOSDATA');

  end;
  CommitTrans;
end;

procedure GIGAChangeUnite;
begin
  ChangeUnite;
end;

procedure GIGAChangeNomAffaire;
begin
  if TableExiste('AFPLANNING') and (tableToVersion('AFPLANNING') < 104) then
  begin
    AddChamp('AFPLANNING', 'APL_INITPTVTDEVHT', 'DOUBLE', 'APL', False, 0, '', '', False);
    if ChampPhysiqueExiste('AFPLANNING', 'APL_INITPTVENTEDEVHT') then
      ExecuteSQL('UPDATE AFPLANNING SET APL_INITPTVTDEVHT=APL_INITPTVTDEVHT');
    AddChamp('AFPLANNING', 'APL_REALPTVTDEVHT', 'DOUBLE', 'APL', False, 0, '', '', False);
    if ChampPhysiqueExiste('AFPLANNING', 'APL_REALPTVENTEDEVHT') then
      ExecuteSQL('UPDATE AFPLANNING SET APL_REALPTVTDEVHT=APL_REALPTVENTEDEVHT');
  end;
  if TableExiste('TACHE') and (tableToVersion('TACHE') < 109) and (tableToVersion('TACHE') > 106) then
  begin
    AddChamp('TACHE', 'ATA_INITPTVTDEVHT', 'DOUBLE', 'ATA', False, 0, '', '', False);
    if ChampPhysiqueExiste('TACHE', 'ATA_INITPTVENTEDEVHT') then
      ExecuteSQL('UPDATE TACHE SET ATA_INITPTVTDEVHT=ATA_INITPTVTDEVHT');
    AddChamp('TACHE', 'ATA_RAPPTVTDEVHT', 'DOUBLE', 'ATA', False, 0, '', '', False);
    if ChampPhysiqueExiste('TACHE', 'ATA_RAPPTVENTEDEVHT') then
      ExecuteSQL('UPDATE TACHE SET ATA_RAPPTVTDEVHT=ATA_RAPPTVENTEDEVHT');
  end;
  if TableExiste('TACHERESSOURCE') and (tableToVersion('TACHERESSOURCE') < 103) and (tableToVersion('TACHERESSOURCE') > 101) then
  begin
    AddChamp('TACHERESSOURCE', 'ATR_INITPTVTDEVHT', 'DOUBLE', 'ATR', False, 0, '', '', False);
    if ChampPhysiqueExiste('TACHERESSOURCE', 'ATR_INITPTVENTEDEVHT') then
      ExecuteSQL('UPDATE TACHERESSOURCE SET ATR_INITPTVTDEVHT=ATR_INITPTVTDEVHT');
    AddChamp('TACHERESSOURCE', 'ATR_RAPPTVTDEVHT', 'DOUBLE', 'ATR', False, 0, '', '', False);
    if ChampPhysiqueExiste('TACHERESSOURCE', 'ATR_RAPPTVENTEDEVHT') then
      ExecuteSQL('UPDATE TACHERESSOURCE SET ATR_RAPPTVTDEVHT=ATR_RAPPTVENTEDEVHT');
  end;

end;

//mcd 032003 debut

procedure GIGAIndexActivite;
var
  TobAff, TobAct, TobDet: TOB;
  QQ1, QQ2: Tquery;
  ii, jj, iNumLigne: integer;
  sAncienTypeArticle, sReqQ2, stReq: string;
begin
  TobAff := nil;
  QQ1 := nil;
  try
    SourisSablier;
    InitMove(3, '');

    if not (ExisteSQL('SELECT ACT_TYPEACTIVITE FROM ACTIVITE')) then exit; // rien dans la table activite dans la base, on sort...
    MoveCur(False);

    TobAff := Tob.Create('les affaires', nil, -1);
    // on traite l'activité par affaire pour éviter de récupérer trop d'enrgt à la fois
    QQ1 := OpenSql('SELECT DISTINCT ACT_AFFAIRE FROM ACTIVITE', true);

    if not (QQ1.eof) then
    begin
      TobAff.LoadDetailDB('', '', '', QQ1, False);

      InitMove(TobAff.Detail.count, '');
      for ii := 0 to TobAff.Detail.count - 1 do
      begin
        if V_PGI.IoError <> oeOk then exit;
        QQ2 := nil;
        try
          sReqQ2 := 'SELECT ACT_TYPEACTIVITE,ACT_AFFAIRE,ACT_RESSOURCE,ACT_DATEACTIVITE,ACT_FOLIO,'
            + 'ACT_TYPEARTICLE,ACT_NUMLIGNE FROM ACTIVITE WHERE ';
          if (TobAff.detail[ii].GetValue('ACT_AFFAIRE') <> null) then
            sReqQ2 := sReqQ2 + 'ACT_AFFAIRE="' + TobAff.detail[ii].GetValue('ACT_AFFAIRE') + '" ORDER BY ACT_TYPEACTIVITE'
          else
            sReqQ2 := sReqQ2 + 'ACT_AFFAIRE IS NULL ORDER BY ACT_TYPEACTIVITE';

          QQ2 := Opensql(sReqQ2, True);
          if not (QQ2.Eof) then
          begin
            TobAct := Tob.Create('les Act', nil, -1);
            iNumLigne := 0;
            sAncienTypeArticle := '';
            try
              TobAct.LoadDetailDb('', '', '', QQ2, False);
              for jj := 0 to TobAct.Detail.count - 1 do
              begin
                TOBDet := TObAct.detail[jj];
                if (sAncienTypeArticle <> TobDet.Getvalue('ACT_TYPEACTIVITE')) then
                begin
                  iNumLigne := 1;
                  sAncienTypeArticle := TobDet.Getvalue('ACT_TYPEACTIVITE');
                end
                else
                  Inc(iNumLigne);

                  stReq := 'UPDATE ACTIVITE SET ACT_NUMLIGNEUNIQUE=' + inttostr(iNumLigne) +
                    ' WHERE ACT_TYPEACTIVITE="' + TobDet.Getvalue('ACT_TYPEACTIVITE') + '"'
                    + ' AND ACT_AFFAIRE="' + TobDet.Getvalue('ACT_AFFAIRE') + '"';
                    if   (TobDet.Getvalue('ACT_RESSOURCE') <> null )then
                      stReq:=stReq+ ' AND ACT_RESSOURCE="' + TobDet.Getvalue('ACT_RESSOURCE') + '"'
                    else
                      stReq:=stReq+ ' AND ACT_RESSOURCE IS NULL';
                    stReq:=stReq + ' AND ACT_DATEACTIVITE="' + UsDateTime(TobDet.Getvalue('ACT_DATEACTIVITE')) + '"'
                    + ' AND ACT_FOLIO=' + inttostr(TobDet.Getvalue('ACT_FOLIO'));
                    if   (TobDet.Getvalue('ACT_TYPEARTICLE') <> null )then
                      stReq:=stReq+ ' AND ACT_TYPEARTICLE="' + TobDet.Getvalue('ACT_TYPEARTICLE') + '"'
                    else
                      stReq:=stReq+ ' AND ACT_TYPEARTICLE IS NULL';
                    stReq:=stReq+ ' AND ACT_NUMLIGNE=' + inttostr(TobDet.Getvalue('ACT_NUMLIGNE'));
                try
                  ExecuteSql(stReq);
                except
                  V_PGI.IoError := oeSaisie;
                end;

              end;

            finally
              TobAct.free;
            end;
          end;
        finally
          Ferme(QQ2);
        end;

        MoveCur(False);
      end; // boucle for
    end;

  finally
    Ferme(QQ1);
    TobAff.free;
    FiniMove;
    SourisNormale;
  end;
end;
//mcd 032003  fin

// GP 08/11/2002
(*
function imo_FormaterTauxArrondi(fTaux: double): string;
begin
  Result := FloatToStr(fTaux);
  if ((fTaux >= 33.33) and (fTaux < 33.34)) then Result := '33' + DecimalSeparator + '333' else
    if ((fTaux >= 16.66) and (fTaux < 16.67)) then Result := '16' + DecimalSeparator + '667' else
    if ((fTaux >= 14.99) and (fTaux < 15.01)) then Result := '15' else
    if ((fTaux >= 14.28) and (fTaux < 14.29)) then Result := '14' + DecimalSeparator + '286' else
    if ((fTaux >= 11.11) and (fTaux < 11.12)) then Result := '11' + DecimalSeparator + '11' else
    if ((fTaux >= 8.33) and (fTaux < 8.34)) then Result := '8' + DecimalSeparator + '333' else
    if ((fTaux >= 6.66) and (fTaux < 6.67)) then Result := '6,667' else
    if ((fTaux >= 3.333) and (fTaux < 3.334)) then Result := '3' + DecimalSeparator + '333' else
    if ((fTaux >= 2.85) and (fTaux < 2.86)) then Result := '2' + DecimalSeparator + '857' else
    if ((fTaux >= 2.222) and (fTaux < 2.223)) then Result := '2' + DecimalSeparator + '222' else
    Result := FloatToStr(Arrondi(fTaux, 3));
end;

procedure MajTableDureeTaux;
var TobTravail, TobT: Tob;
  LibSQL, LibDuree: string;
  Q: TQuery;
  idx: integer;
  Duree, Taux, Coef: double;
  TobIDT: Tob;
  iCode: Integer;
  STaux, StCode: string;
begin
  // remplissage de la tob avec les durées existantes dans CHOIXCOD
  TobTravail := Tob.Create('TobDurTau', nil, -1);
  LibSQL := 'Select CC_TYPE, CC_CODE, CC_LIBELLE, CC_ABREGE from CHOIXCOD where CC_TYPE="IDU"';
  Q := openSql(LibSQL, false);
  TobTravail.LoadDetailDB('TobDurTau', '', '', Q, true);
  Ferme(Q);
  Q := OpenSQL('SELECT MAX(IDT_CODEIDT) FROM IDUREETAUX', TRUE);
  if not Q.Eof then StCode := Q.Fields[0].AsString;
  Ferme(Q);
  if StCode <> '' then ICode := StrToInt(StCode) else ICode := 0;
  TobIDT := Tob.Create('IDUREETAUX', nil, -1);
  for idx := 0 to TobTravail.Detail.Count - 1 do
  begin
    TobIDT.ClearDetail;
    TobT := TobTravail.Detail[idx];
    Duree := TobT.GetValue('CC_ABREGE');
    if (Duree <> 0) and (not ExisteSql('select IDT_DUREEAMT from IDUREETAUX where IDT_DUREEAMT=' + StrFPoint(Duree))) then
    begin
      Taux := 100 / Duree;
      if Duree <= 1 then LibDuree := ' an'
      else LibDuree := ' ans';

      // récupération du coef existant dans CHOIXCOD
      Q := OpenSql('Select CC_LIBELLE from CHOIXCOD where CC_TYPE="IDG" and CC_CODE="' + Copy(TobT.GetValue('CC_CODE'), 2, 2) + 'A' + '"', false);
      if not Q.Eof then Coef := Q.fields[0].asFloat else Coef := 1;
      Ferme(Q);
      STaux := imo_FormaterTauxArrondi(Taux);
      TobIDT.AddChampSupValeur('IDT_DUREEAMT', Duree);
      TobIDT.AddChampSupValeur('IDT_LIBDUREE', StrFPoint(Duree) + LibDuree);
      TobIDT.AddChampSupValeur('IDT_TAUXAMT', STaux);
      TobIDT.AddChampSupValeur('IDT_LIBTAUX', STaux + ' %');
      TobIDT.AddChampSupValeur('IDT_COEFDEG1', Coef);
      Inc(ICode);
      TobIDT.AddChampSupValeur('IDT_CODEIDT', FormatFloat('000', Icode));

      TobIDT.InsertOrUpdateDB(FALSE);
    end;

  end;
  TobIDT.Free;
  TobTravail.Free;
end;
// FIN GP 08/11/2002
*)
procedure MajJuridiqueTypePer;
var TobJutypePer, TJ: Tob;
  DateAjoutTypePer: TDateTime;
begin
  //DateAjoutTypePer := StrTodate('23/10/2002');
  DateAjoutTypePer :=   Encodedate(2002, 10, 23);

  TobJutypePer := Tob.Create('Les JuTypePer', nil, -1);
  TJ := Tob.Create('JUTYPEPER', TobJutypePer, -1);
  TJ.PutValue('JTP_AFFICHE', 'AGI');
  TJ.PutValue('JTP_APPMODIF', 'Decla');
  TJ.PutValue('JTP_DATECREATION', DateAjoutTypePer);
  TJ.PutValue('JTP_DATEMODIF', DateAjoutTypePer);
  TJ.PutValue('JTP_DEFAUTLIB1', '');
  TJ.PutValue('JTP_DEFAUTLIB2', '');
  TJ.PutValue('JTP_DEFAUTLIB3', '');
  TJ.PutValue('JTP_DEFAUTNOM', '');
  TJ.PutValue('JTP_FAMPER', 'SOC');
  TJ.PutValue('JTP_PREDEFINI', 'CEG');
  TJ.PutValue('JTP_RACINE', 'AGI');
  TJ.PutValue('JTP_TYPEPER', 'AGI');
  TJ.PutValue('JTP_TYPEPERABREG', 'Agence d''interim');
  TJ.PutValue('JTP_TYPEPERLIB', 'agence d''interim');
  TJ.PutValue('JTP_UTILISATEUR', '000');
  TJ.PutValue('JTP_VERSION', 3);
  TJ := Tob.Create('JUTYPEPER', TobJutypePer, -1);
  TJ.PutValue('JTP_AFFICHE', 'CFO');
  TJ.PutValue('JTP_APPMODIF', 'Decla');
  TJ.PutValue('JTP_DATECREATION', DateAjoutTypePer);
  TJ.PutValue('JTP_DATEMODIF', DateAjoutTypePer);
  TJ.PutValue('JTP_DEFAUTLIB1', '');
  TJ.PutValue('JTP_DEFAUTLIB2', '');
  TJ.PutValue('JTP_DEFAUTLIB3', '');
  TJ.PutValue('JTP_DEFAUTNOM', '');
  TJ.PutValue('JTP_FAMPER', 'SOC');
  TJ.PutValue('JTP_PREDEFINI', 'CEG');
  TJ.PutValue('JTP_RACINE', 'CFO');
  TJ.PutValue('JTP_TYPEPER', 'CFO');
  TJ.PutValue('JTP_TYPEPERABREG', 'Centre de formation');
  TJ.PutValue('JTP_TYPEPERLIB', 'Centre de formation');
  TJ.PutValue('JTP_UTILISATEUR', '000');
  TJ.PutValue('JTP_VERSION', 3);
  TobJutypePer.InsertOrUpdateDB(False);
  TobJutypePer.free;
end;


procedure MajGPAOAvant;
var
  sSQL : String;
  Q   : TQuery;
begin
  BeginTrans ;
  sSQL := '';
  //-- Suppression des vues du domaine "W" --//

  // Suprimé car pbm de synchronisatio avec l'AGL qui lui-même detruit les vues vant mise à jour //

{ sSQL := 'SELECT DV_NOMVUE, DV_LIBELLE ' +
          'FROM DEVUES '+
          'WHERE DV_DOMAINE = "W"';

  Q := OpenSQL(sSQL, True);
  while not Q.Eof do
  begin
    // fait par l'AGL ExecuteSQL('DROP VIEW ' + Q.Fields[0].AsString);
    ExecuteSQL('DELETE FROM DEVUES WHERE DV_DOMAINE = "W"');
    Q.Next
  end;
}
  //-- Suppression des tables du domaine "W" --//
  sSQL :='SELECT DT_NOMTABLE, DT_NUMVERSION, DT_PREFIXE, DT_BLOCNOTE '+
         'FROM DETABLES '+
         'WHERE DT_DOMAINE = "W" AND DT_PREFIXE <> "WTV" ' +
         'ORDER BY DT_NOMTABLE';
  Q := OpenSQL(sSQL, True);
  while not Q.Eof do
  begin
    TestDropTable(Q.Fields[0].AsString);
    Q.Next
  end;
  ConnecteTRUEFALSE(DBSOC);
  //-- Suppression des tablettes du domaine "W" --//
      ExecuteSQL('DELETE FROM COMMUN WHERE CO_TYPE LIKE "W%"');
      ExecuteSQL('DELETE FROM CHOIXCOD WHERE CC_TYPE LIKE "W%"');
      ExecuteSQL('DELETE FROM CHOIXEXT WHERE YX_TYPE LIKE "W%"');
      ExecuteSQL('DELETE FROM DECOMBOS WHERE DO_DOMAINE="X"');

  //-- Suppression des listes du domaine "W" --//
	ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE LIKE "W%" OR LI_LISTE LIKE "PRT_W%"');
  //-- Suppression des fiches du domaine "W" --//
  ExecuteSQL('DELETE FROM FORMES WHERE DFM_TYPEFORME LIKE "W%"');
	ExecuteSQL('DELETE FROM SCRIPTS WHERE SRC_TYPESCRIPT="W%"');
  //-- Suppression des états du domaine "W" --//
  ExecuteSQL('DELETE FROM MODELES WHERE MO_NATURE LIKE "W"');
  ExecuteSQL('DELETE FROM COMMUN WHERE CO_TYPE="NTD" AND CO_CODE LIKE "W%"');
  //-- Suppression des graphs du domaine "W" --//
	ExecuteSQL('DELETE FROM GRAPHS WHERE GR_GRAPHE LIKE "TW%"');
  CommitTrans;
end;

//GPAO
procedure MajGPAOApres;
var
  sSQL     : string ;
  iNum     : integer ;
  iInc     : integer ;
  sType    : string ;
  sCode    : string ;
  sLibelle : string ;
  sPrefixe : string ;

      { Insertion d'enregistrement dans la table MEA }
      procedure pInsertMEA(sQualif, sMesure, sLibelle, dQuotite : string);
      var
         sSQL : string;
      begin
        { sSQL := 'SELECT * FROM MEA ' +
                 'WHERE GME_QUALIFMESURE = "' + sQualif + '" ' +
                 'AND GME_MESURE = "' + sMesure + '" '; }
         // modif PCS GME_MESURE est une clé unique        
         sSQL := 'SELECT * FROM MEA ' +
                 'WHERE GME_MESURE = "' + sMesure + '" ';
         if not ExisteSQL(sSQL) then
            ExecuteSQL('INSERT INTO MEA (GME_QUALIFMESURE, GME_MESURE, GME_LIBELLE, GME_QUOTITE)'
                      +'VALUES("' + sQualif + '","' + sMesure + '","' + sLibelle + '",' + dQuotite + ')');
      end;
      { Insertion d'enregistrement dans la table "WCBNTYPEMOUV" }
      procedure pInsertWCBNTYPEMOUV(sType, sLibelle : string);
      var
         sSQL : string;
      begin
         sSQL := 'SELECT * FROM WCBNTYPEMOUV ' +
                 'WHERE WTM_TYPEMOUVEMENT = "' + sType + '" ';

         if not ExisteSQL(sSQL) then
            ExecuteSQL('INSERT INTO WCBNTYPEMOUV (WTM_TYPEMOUVEMENT,WTM_LIBELLE,WTM_DATECREATION,'+
                       'WTM_DATEMODIF,WTM_CREATEUR,WTM_UTILISATEUR,WTM_SOCIETE) ' +
                       'VALUES("' + sType + '","' + sLibelle + '",1900-01-01,2002-10-22,"001","001","001")');
      end;
begin
   BeginTrans ;

      {***************************************************}
      {******* Insertion destinée à des tablettes   ******}
      {***************************************************}

      //-- Insertion des natures de travail --//

      if not ExisteSQL('SELECT * FROM WNATURETRAVAIL') then
      begin
         sSQL :='INSERT INTO WNATURETRAVAIL (WNA_NATURETRAVAIL,WNA_DATECREATION,' +
                'WNA_DATEMODIF,WNA_CREATEUR,WNA_UTILISATEUR,WNA_SOCIETE,WNA_LIBELLE,'+
                'WNA_NATUREPIECEA,WNA_NATUREPIECEB,WNA_WITHMOD,WNA_WITHVAL,WNA_WITHPER,'+
                'WNA_ACTIF)VALUES("FAB",1900-01-01,2002-10-09,"CEG","CEG","001",'+
                '"Fabrication","WAP","WBP","X","X","-","X")';
         ExecuteSQL(sSQL);

         sSQL :='INSERT INTO WNATURETRAVAIL (WNA_NATURETRAVAIL,WNA_DATECREATION,' +
                'WNA_DATEMODIF,WNA_CREATEUR,WNA_UTILISATEUR,WNA_SOCIETE,WNA_LIBELLE,'+
                'WNA_NATUREPIECEA,WNA_NATUREPIECEB,WNA_WITHMOD,WNA_WITHVAL,WNA_WITHPER,'+
                'WNA_ACTIF)VALUES("RET",1900-01-01,2002-10-09,"CEG","CEG","001","Retouche","","WBP","X","X","-","-")';
         ExecuteSQL(sSQL);
      end;


      {***************************************}
      {******* Insertion dans ChoixExt  ******}
      {***************************************}

      //-- Tablette "WORIGINEPDR" --//
      InsertChoixExt('WRO','GAMME','Gamme','Gamme','');
      InsertChoixExt('WRO','NOMENCLATURE','Nomenclature','Nomenclature','');

      //-- Tablette "WTYPECOUT" --//
      InsertChoixExt('WPV','COUT DIRECT','Coût direct','Coût direct','');
      InsertChoixExt('WPV','COUT INDIRECT','Coût indirect','Coût indirect','');
      InsertChoixExt('WPV','VALEUR MARGE','Montant de marge','Montant de marge','');

      //-- Tablette "WTYPEART" --//
      InsertChoixExt('WRP','PERTE FIXE','Perte fixe','Fixe','Nomenclature');
      InsertChoixExt('WRP','PERTE FREINTE','Perte frainte','Frainte','Nomenclature');
      InsertChoixExt('WRP','PERTE PERIODIQUE','Perte périodique','Périodique','Nomenclature');
      InsertChoixExt('WRP','PERTE PROPORTION','Perte proportionnelle','Proportionnelle','Nomenclature');
      InsertChoixExt('WRP','TEMPS EXECUTION','Temps exécution','Exécution','Gamme');
      InsertChoixExt('WRP','TEMPS PREPARATION','Temps préparation','Préparation','Gamme');
      InsertChoixExt('WRP','VARIABLE','Variable','Variable','Nomenclature');

      //-- Champs libres //
      for iNum:=0 to 7 do
      begin
         case iNum of
            0 : sPrefixe := 'WNT';
            1 : sPrefixe := 'WNL';
            2 : sPrefixe := 'WGT';
            3 : sPrefixe := 'WGL';
            4 : sPrefixe := 'WOT';
            5 : sPrefixe := 'WOP';
            6 : sPrefixe := 'WPC';
            7 : sPrefixe := 'WVS';
         end ;
         for iInc := 1 to 3 do
         begin
            //-- Booleens libres //
            sType    := 'W' + IntToStr(iNum) + 'Y' ;
            sCode    := 'T' + sPrefixe + '_BOOLLIBRE' + IntToStr(iInc) ;
            sLibelle := 'Décision libre ' + IntToStr(iInc);
            InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
            //-- Char libres --//
            sType    := 'W' + IntToStr(iNum) + 'Z' ;
            sCode    := 'T' + sPrefixe + '_CHARLIBRE' + IntToStr(iInc) ;
            sLibelle := 'Texte libre ' + IntToStr(iInc);
            InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
            //-- Dates libres --//
            sType    := 'W' + IntToStr(iNum) + 'X' ;
            sCode    := 'T' + sPrefixe + '_DATELIBRE' + IntToStr(iInc) ;
            sLibelle := 'Date libre ' + IntToStr(iInc);
            InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
            //-- Valeur libres --//
            sType    := 'W' + IntToStr(iNum) + 'W' ;
            sCode    := 'T' + sPrefixe + '_VALLIBRE' + IntToStr(iInc) ;
            sLibelle := 'Valeur libre ' + IntToStr(iInc);
            InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
         end;
         //-- Libellés libres --//
         for iInc := 1 to 10 do
         begin
            case iInc of
            1..9 : begin
                     sType    := 'W' + IntToStr(iNum) + 'V' ;
                     sCode    := 'T' + sPrefixe + '_LIBRE'  + sPrefixe + IntToStr(iInc) ;
                     sLibelle := 'Table libre ' + IntToStr(iInc);
                     InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
                  end;
            10  : begin
                     sType    := 'W' + IntToStr(iNum) + 'V' ;
                     sCode    := 'T' + sPrefixe + '_LIBRE' + sPrefixe +  'A' ;
                     sLibelle := 'Table libre 10';
                     InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
                  end;
            end;
         end;
      end;

      {***************************************}
      {******* Insertion dans ChoixCode ******}
      {***************************************}

      //-- Tablette "WCODEFORME" --//
      InsertChoixCode ('WWF','CUB','Cube','Cube','');
      InsertChoixCode ('WWF','CYL','Cylindre','Cylindre','');

      //-- Tablette "WINITARTICLE" --//
      InsertChoixCode ('WIA','010','GA_QUALIFUNITESTO','Stock','Unité de stock;UNITEDEF=[DEF]');
      InsertChoixCode ('WIA','020','GA_UNITEQTEVTE','Vente','Unité de vente (Quantité);UNITEDEF=[DEF]');
      InsertChoixCode ('WIA','030','GA_UNITEQTEACH','Achat','Unité d''achat (Quantité);UNITEDEF=[DEF]');
      InsertChoixCode ('WIA','050','GA_QUALIFUNITEVTE','Vente','Unité de vente;DEFAUT;UNITEDEF=[DEF]');
      InsertChoixCode ('WIA','060','GA_UNITEPROD','Production','Unité de production;UNITEDEF=[DEF]');
      InsertChoixCode ('WIA','070','GA_UNITECONSO','Consommation','Unité de consommation;UNITEDEF=[DEF]');
      InsertChoixCode ('WIA','080','GA_QUALIFPOIDS','Poids','Unité des poids;UNITEDEF=K');
      InsertChoixCode ('WIA','110','GA_QUALIFLINEAIRE','Linéaire','Unité de longueur;UNITEDEF=M');
      InsertChoixCode ('WIA','120','GA_QUALIFSURFACE','Surface','Unité de surface;UNITEDEF=M2');
      InsertChoixCode ('WIA','130','GA_QUALIFVOLUME','Volume','Unité de volume;UNITEDEF=M3');
      InsertChoixCode ('WIA','140','GA_QUALIFHEURE','Temps','Unité de temps;UNITEDEF=H');

      //-- Tablette "WMETHODEAPPRO" --//
      InsertChoixCode ('WMA','CMD','A la commande','A la commande','');
      InsertChoixCode ('WMA','MRP','M.R.P','M.R.P','');
      InsertChoixCode ('WMA','NG' ,'Non géré','Non géré','');


      //-- Tablette "WMOTIFREBUT" --//
      InsertChoixCode ('WMR','001','Mon premier motif','Mon premier motif','');
      InsertChoixCode ('WMR','002','Mon second motif','Mon second motif','');

      //-- Tablette "WMOTIFSUSPECT" --//
      InsertChoixCode ('WMS','001','Mon suspect 1','Mon suspect 1','');
      InsertChoixCode ('WMS','002','Mon suspect 2','Mon suspect 2','');

      //-- Tablette "WTARIFSOU" --//
      InsertChoixCode ('WTU','C','à la commande','à la commande','');
      InsertChoixCode ('WTU','L','à la ligne','à la ligne','');

      //-- Tablette "WTARIFSQUOI" --//
      InsertChoixCode ('WTN','L','Des linéaires','Linéaires','');
      InsertChoixCode ('WTN','P','Des poids','Poids','');
      InsertChoixCode ('WTN','Q','Des quantités','Quantités','');
      InsertChoixCode ('WTN','S','Des surfaces','Surfaces','');
      InsertChoixCode ('WTN','V','Des volumes','Volumes','');
      //-- Tablettes YYCONTEXTES --//
      InsertChoixCode ('YCT','GP','GPAO','ctxGPAO','');
      {****************************************}
      {****** Traitements sur les tables ******}
      {****************************************}
      //-- Table "WCBNTYPEMOUV" --//
      pInsertWCBNTYPEMOUV('ACC', 'Achat confirmé');
      pInsertWCBNTYPEMOUV('ACN', 'Achat non confirmé');
      pInsertWCBNTYPEMOUV('BOF', 'Besoin / Ordre de fabrication');
      pInsertWCBNTYPEMOUV('BPF', 'Besoin / Proposition de fabrication');
      pInsertWCBNTYPEMOUV('ORF', 'Ordre de fabrication');
      pInsertWCBNTYPEMOUV('PRA', 'Proposition d''achat');
      pInsertWCBNTYPEMOUV('PRF', 'Proposition d''ordre de fabrication');
      pInsertWCBNTYPEMOUV('PRS', 'Proposition de Ss traitance d''achat');
      pInsertWCBNTYPEMOUV('PRT', 'Proposition transfert inter-dépôt');
      pInsertWCBNTYPEMOUV('STC', 'Stock en contrôle');
      pInsertWCBNTYPEMOUV('STO', 'Stock disponible');
      pInsertWCBNTYPEMOUV('VEN', 'Commande de vente');
      //-- Table "WNATURETRAVAIL" --//
      sSQL := 'UPDATE WNATURETRAVAIL ' +
              'SET WNA_ACTIF="-" ' +
              'WHERE WNA_NATURETRAVAIL <> "FAB"';
      ExecuteSQL(sSQL);
      sSQL := 'UPDATE WNATURETRAVAIL SET ' +
              'WNA_WITHMOD="X",' +
              'WNA_WITHVAL="X",' +
              'WNA_WITHPER="-" ' +
              'WHERE (WNA_WITHMOD IS NULL OR WNA_WITHMOD="")';
      ExecuteSQL(sSQL);
      //-- Table "MEA" --//
      // pInsertMEA('PIE','U','Unité','1.0000'); // suprimée le 17082004 demande de Karine

      // suprimée le 14092005 demande de Karine
      //pInsertMEA('TEM','H','Heure','1.0000');
      //pInsertMEA('TEM','J','Jour','8.0000');
      //pInsertMEA('TEM','MIN','Minute','0.0100');
      //pInsertMEA('TEM','SEC','Seconde','0.0001');
      //pInsertMEA('THC','HC','Heure centième','1.0000');
   CommitTrans;
end;

procedure CreationEnrTYPEREMISE;
var
  TOBCC, TOBDem, TOBL, TOBD: TOB;
  Ind, Pourc: integer;
  sLibre, Stg: string;
begin
  TOBCC := TOB.Create('', nil, -1);
  TOBCC.LoadDetailDB('CHOIXCOD', '"GTM"', '', nil, False);
  TOBDem := TOB.Create('', nil, -1);
  TOBDem.LoadDetailDB('TYPEREMISE', '', '', nil, False);
  for Ind := 0 to TOBCC.Detail.Count - 1 do
  begin
    TOBL := TOBCC.Detail[Ind];
    if TOBDem.FindFirst(['GTR_TYPEREMISE'], [TOBL.GetValue('CC_CODE')], False) = nil then
    begin
      TOBD := TOB.Create('TYPEREMISE', TOBDem, -1);
      TOBD.InitValeurs;
      TOBD.PutValue('GTR_TYPEREMISE', TOBL.GetValue('CC_CODE'));
      TOBD.PutValue('GTR_LIBELLE', TOBL.GetValue('CC_LIBELLE'));
      TOBD.PutValue('GTR_ABREGE', TOBL.GetValue('CC_ABREGE'));
      sLibre := TOBL.GetValue('CC_LIBRE');
      Stg := Copy(sLibre, 1, 1);
      if Stg <> 'X' then Stg := '-';
      TOBD.PutValue('GTR_CLIOBLIGFO', Stg);
      Stg := Copy(sLibre, 2, 1);
      if Stg <> '-' then Stg := 'X';
      TOBD.PutValue('GTR_IMPRIMABLE', Stg);
      Stg := Trim(Copy(sLibre, 3, 3));
      if IsNumeric(Stg) then Pourc := StrToInt(Stg) else Pourc := 0;
      if (Pourc < 0) or (Pourc > 100) then Pourc := 0;
      TOBD.PutValue('GTR_REMPOURMAX', Pourc);
      TOBD.PutValue('GTR_FIDELITE', '-');
      TOBD.PutValue('GTR_EXCLUFIDELITE', '-');
      TOBD.PutValue('GTR_FERME', '-');
      TOBD.PutValue('GTR_DATEDEBUT', iDate1900);
      TOBD.PutValue('GTR_DATEFIN', iDate2099);
      TOBD.UpdateDateModif;
    end;
  end;
  TOBDem.InsertOrUpdateDB;
  TOBCC.DeleteDB;
  TOBCC.Free;
  TOBDem.Free;
end;

Procedure PGMinConvPaie;
var
  ii    : Integer ;
  Q    : TQuery ;
  LeNbre,NbreP : integer;

    function ZZColleZeroDevant(Nombre , LongChaine : integer) : string ;
    var
    tabResult : string ;
    TabInt : string;
    i,j : integer;
    begin
    tabResult := '';
       for i := 1 to LongChaine do begin
          if Nombre < power(10,i) then
          begin
             TabInt := inttostr(Nombre);
            // colle (LongChaine-i zéro devant]
             for j := 0 to  (LongChaine-i-1)
                          do insert('0',TabResult,j);
             result := concat(TabResult,Tabint);
             exit;
          end;
        if i > LongChaine then result := inttostr(Nombre);
       end;
    end;
begin
if TableToVersion('MINCONVPAIE') >= 104 then Exit;

Q := opensql('SELECT max (PCP_NOMBRE) LENBRE FROM MINCONVPAIE',True) ;
IF NOT Q.eof THEN LeNbre := Q.FindField ('LENBRE').AsInteger ;
if LeNbre >= 999 then LeNbre := 997
  else LeNbre := 999 ;
Ferme (Q) ;

// Rajout champ PCP_NBRE
ExecuteSql('DELETE FROM MINCONVPAIE WHERE (PCP_PREDEFINI="CEG") OR (PCP_PREDEFINI = "") OR (PCP_PREDEFINI IS NULL)') ;
AddChamp('MINCONVPAIE', 'PCP_NBRE', 'VARCHAR(17)', 'PCP', False, '', 'Standard', '', False);

   // if not tableSurAutreBase('MINCONVPAIE') then  by PCS and MD 09112004 DESHARE désactivé dans MajStruct
   if IsMonoOuCommune then
  begin
    q := opensql('SELECT * FROM MINCONVPAIE',False) ;
    ii := 0 ;
    While not q.eof do
    begin
      Q.Edit;
      NbreP:= Q.FindField('PCP_NOMBRE').AsInteger ;
      if NbreP > 999 then
      begin
        NbreP := LeNbre  - ii ;
        ii := ii + 2 ;
      end ;
      Q.FindField('PCP_NBRE').AsString :=  ZZColleZeroDevant(NbreP, 3) ;
      Q.Post;
      Q.Next;
    end ;
    Ferme(q) ;
  end ;
end ;


Procedure InitTablePaie605;
begin
ExecuteSQL('UPDATE ETABCOMPL SET ETB_ACTIVITE="",ETB_DATECREATION="' + UsTime(NowH) + '",ETB_DATEMODIF  ="'+UsTime(NowH) +'",ETB_TYPTICKET ="", ETB_DADSSECTION="00", ETB_TYPDADSSECT="1"');
ExecuteSQL('UPDATE DEPORTSAL SET PSE_TYPETYPTK ="",PSE_TYPTICKET ="",PSE_DISTRIBUTION ="",PSE_PERSONNAL="",PSE_INFOCOMPL=""');
ExecuteSQL('UPDATE CUMULRUBRIQUE SET PCR_DATECREATION="' + UsTime(NowH) + '",PCR_DATEMODIF  ="'+UsTime(NowH) +'"');
ExecuteSQL('UPDATE PROFILRUB SET PPM_DATECREATION="' + UsTime(NowH) + '",PPM_DATEMODIF  ="'+UsTime(NowH) +'"');
ExecuteSQL('UPDATE PAIEENCOURS SET PPU_DATECREATION="' + UsTime(NowH) + '",PPU_DATEMODIF  ="'+UsTime(NowH) +'",'+
           'PPU_CREATEUR   ="CEG", PPU_UTILISATEUR="CEG"');
ExecuteSQL('UPDATE VENTIREMPAIE SET PVS_DATECREATION="' + UsTime(NowH) + '",PVS_DATEMODIF  ="'+UsTime(NowH) +'"');
ExecuteSQL('UPDATE VENTICOTPAIE SET PVT_DATECREATION="' + UsTime(NowH) + '",PVT_DATEMODIF  ="'+UsTime(NowH) +'"');
ExecuteSQL('UPDATE VENTIORGPAIE SET PVO_DATECREATION="' + UsTime(NowH) + '",PVO_DATEMODIF  ="'+UsTime(NowH) +'"');
ExecuteSQL('UPDATE MINCONVPAIE SET PCP_DATEMODIF  ="'+UsTime(NowH) +'"');
ExecuteSQL('UPDATE HISTOBULLETIN SET PHB_DATECREATION="' + UsTime(NowH) + '",PHB_DATEMODIF  ="'+UsTime(NowH) +'",'+
           'PHB_UTILISATEUR="CEG"');
ExecuteSQL('UPDATE HISTOCUMSAL SET PHC_DATECREATION="' + UsTime(NowH) + '",PHC_DATEMODIF  ="'+UsTime(NowH) +'",'+
           'PHC_UTILISATEUR="CEG"');
ExecuteSQL('UPDATE COTISATION SET PCT_PRESFINMOIS = "-"');
ExecuteSQL('UPDATE ATTESTATIONS SET PAS_ASSEDICCAISSE = "", PAS_ASSEDICNUM="", PAS_DERNIEREMPLOI="", PAS_PREAVISMOTIF="", PAS_HORHEBENT=0, PAS_HORANNENT=0, PAS_HORHEBSAL=0, PAS_HORANNSAL=0, PAS_MOTRUPCONT=""');
ExecuteSQL('UPDATE ORGANISMEPAIE SET POG_REGROUPDADSU="-"');
ExecuteSQL('UPDATE SALARIES SET PSA_TYPACTIVITE="ETB", PSA_ACTIVITE="", PSA_SSDECOMPTE="-"');
ExecuteSQL('UPDATE PROFILPAIE SET PPI_ACTIVITE=""');
ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA = "" WHERE SOC_NOM = "SO_PGCODECLIENT"');
ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA = "" WHERE SOC_NOM = "SO_PGCODERATTACH"');
ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA = "" WHERE SOC_NOM = "SO_PGFRAISGESTION"');
ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA = "-" WHERE SOC_NOM = "SO_PGPERSOTICKET"');
ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA = "" WHERE SOC_NOM = "SO_PGTYPEPREPACDE"');
ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA = "" WHERE SOC_NOM = "SO_PGREPERTTICKET"');
ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA = "-" WHERE SOC_NOM = "SO_PGTICKETRESTAU"');
ExecuteSQL('UPDATE INSCFORMATION SET PFI_DADSCAT="",PFI_REALISE="",PFI_NIVPRIORITE="",PFI_MOTIFINSCFOR="",PFI_ETATINSCFOR="",PFI_MOTIFETATINSC=""');
ExecuteSQL('UPDATE FORMATIONS SET PFO_NIVPRIORITE="",PFO_MOTIFINSCFOR="",PFO_MOTIFETATINSC="",PFO_INCLUSDECL="-"');
ExecuteSQL('UPDATE FORMATIONS SET PFO_ETATINSCFOR="ATT" WHERE PFO_DATEACCEPT="'+UsDateTime(IDate1900)+'" AND PFO_REFUSELE="'+UsDateTime(IDate1900)+'" AND PFO_REPORTELE="'+UsdateTime(IDate1900)+'"');
ExecuteSQL('UPDATE FORMATIONS SET PFO_ETATINSCFOR="REP" WHERE PFO_REPORTELE<>"'+UsDateTime(IDate1900)+'"');
ExecuteSQL('UPDATE FORMATIONS SET PFO_ETATINSCFOR="REF" WHERE PFO_REFUSELE<>"'+UsDateTime(IDate1900)+'"');
ExecuteSQL('UPDATE FORMATIONS SET PFO_ETATINSCFOR="VAL" WHERE PFO_DATEACCEPT<>"'+UsDateTime(IDate1900)+'"');
ExecuteSQL('UPDATE SALAIREFORM SET PSF_CONFIDENTIEL="-",PSF_TRANCHE1=0,PSF_TRANCHE2=0,PSF_TRANCHE3=0');
ExecuteSQL('UPDATE STAGE SET PST_COUTEVAL=0,PST_INCLUSDECL="X",PST_ACTIONFORM="",PST_REGLEMENTAIRE="-"');
ExecuteSQL('UPDATE SESSIONSTAGE SET PSS_ACTIONFORM="",PSS_INCLUSDECL="X",PSS_COUTEVAL=0');
ExecuteSQL('UPDATE SESSIONSTAGE SET PSS_NUMENVOI="-1" WHERE (PSS_NUMENVOI IS NULL OR PSS_NUMENVOI<=0)');
ExecuteSQL('UPDATE SESSIONSTAGE SET PSS_TOTALHT=0 WHERE (PSS_TOTALHT<=0 OR PSS_TOTALHT IS NULL)');
ExecuteSQL('UPDATE SESSIONANIMAT SET PAN_FRASPEDAG="-",PAN_FRAISAVECSAL="-" ');
ExecuteSQL('UPDATE TYPERETENUE SET PTR_NIVEAURET=0,PTR_TYPALIMPAIE="",PTR_CALCFRACTION=""');
ExecuteSQL('UPDATE RETENUESALAIRE SET PRE_LIBELLE="",PRE_NBJOURS=0,PRE_RETENUESAL="",'+
 'PRE_DATECREATION="'+UsDateTime(IDate1900)+'",PRE_DATEMODIF="'+UsDateTime(IDate1900)+'",'+
 'PRE_CREATEUR="",PRE_CLOTURE="-",PRE_REMBMAX="-",PRE_CALCFRACTION="",PRE_REMBEFFECTUE="-",PRE_ECHEANCIER="-"');
ExecuteSQL('UPDATE CURSUS SET PCU_RANGCURSUS=0,PCU_DATEDEBUT="'+UsDateTime(IDate1900)+'",PCU_DATEFIN="'+UsDateTime(IDate1900)+'"');
ExecuteSQL('UPDATE CURSUSSTAGE SET PCC_ORDRE=0,PCC_RANGCURSUS=0');
ExecuteSQL('UPDATE LIEUFORMATION SET PLF_ETABLISSEMENT="",PLF_CODEPOSTAL=""');
ExecuteSQL('UPDATE EXERFORMATION SET PFE_TAUXCHARGEC=0,PFE_TAUXCHARGENC=0,PFE_COUTFIXEFOR1=0,PFE_COUTFIXEFOR2=0,'+
  'PFE_COUTFIXEFOR3=0,PFE_COUTFIXEFOR4=0,PFE_COUTFIXEFOR5=0,PFE_COUTFIXEFOR6=0,PFE_COUTFIXEFOR7=0,PFE_COUTFIXEFOR8=0,PFE_MASSESAL=0,'+
  'PFE_MASSESALCDD=0,PFE_OPCACTFIXE1=-1,PFE_OPCACTFIXE2=-1,PFE_OPCACTFIXE3=-1,PFE_OPCACTFIXE4=-1,PFE_OPCACTFIXE5=-1,'+
  'PFE_OPCACTFIXE6=-1,PFE_OPCACTFIXE7=-1,PFE_OPCACTFIXE8=-1');
ExecuteSQL('UPDATE FRAISSALFORM SET PFS_CTPEDAG="-",PFS_CTSALARIE="-"');
ExecuteSQL('UPDATE EMPLOIINTERIM SET PEI_CONVENTION=""');
end;

procedure MajGPAOAvant612;
begin
  //-- Suppression des listes --//
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN("GCADRESSES","GCSAISIECSA","GCPIEDPORT","GCARTTIERS"'
    +',"GCPAYS","GCREGION","GCSAISIEGP","UORDRELIG","UORDREBES"'
    +',"WCBNPROPOSITION","WARTICLE","WNOMEDEC","WGAMMETET","WGAMMELIG","WNOMETET","WNOMELIG"'
    +',"WORDREBES","WORDREBES2","WORDRELIG2","WORDREPHASE","WORDREPHASE2","WORDRETET", "WORDRECMP_VUE")');
end;

procedure MajGPAOApres612;
var
  sSQL     : string ;
  TheTob   : tob;

  function wLoadTobFromSql(TableName, Sql: string; t:Tob; lAppend: boolean = false; NbLigne: integer = -1): boolean;
  var
    Q: tQuery;
  begin
    Result := false;

    Q := OpenSql(Sql, True, NbLigne);
    try
      if not Q.Eof then
      begin
        t.loadDetailDB(TableName, '', '', Q, lAppend);
        Result := true;
      end;
    finally
      Ferme(Q);
    end;
  end;

  procedure MaJ_ADR_NADRESSE;
  var
    sSQL, sRefCode  : string;
    iNumeroAdresse  : integer;
    TheTob          : tob;
    iCpt            : integer;
  begin

    TheTob := Tob.Create('ADRESSES', nil, -1);
    Try
      ExecuteSQL('UPDATE ADRESSES SET ADR_TYPEADRESSE="TIE" WHERE ADR_TYPEADRESSE="T"');
      sSQL:='SELECT ADR_TYPEADRESSE, ADR_NUMEROADRESSE, ADR_NATUREAUXI, ADR_REFCODE, ADR_NADRESSE FROM ADRESSES '
           + ' WHERE (ADR_TYPEADRESSE IN ("TIE","INT")) and ((ADR_NADRESSE=0) or (ADR_NADRESSE Is Null)) '
           + ' ORDER BY ADR_TYPEADRESSE, ADR_REFCODE';
      wLoadTobFromSql('ADRESSES', sSql, TheTob);

      sRefCode:=''; iNumeroAdresse:=1;
      if assigned(TheTob) then
      begin
        for iCpt:=0 to TheTob.Detail.count-1 do
        begin
          if sRefCode<>TheTob.Detail[iCpt].GetValue('ADR_REFCODE') then
          begin
            sRefCode := TheTob.Detail[iCpt].GetValue('ADR_REFCODE');
            iNumeroadresse:=1;
          end;
          if TheTob.Detail[iCpt].GetValue('ADR_NADRESSE')=0 then
          begin
            TheTob.Detail[iCpt].PutValue('ADR_NADRESSE',iNumeroAdresse);
            iNumeroadresse := iNumeroAdresse+1;
          end;
        end;
        TheTob.UpdateDB(True, False);
      end;
    finally
      TheTob.Free;
    end;
  end;

  Procedure MaJWIL(sString : string);
  var
    TheTob : tob;
  begin
    TheTob := Tob.Create('WINITCHAMPLIG', Nil, -1);
    try
      TheTob.PutValue('WIL_BYPARAMSOC'    , ReadTokenSt(sString));
      TheTob.PutValue('WIL_C1OPERATEUR'   , ReadTokenSt(sString));
      TheTob.PutValue('WIL_C1VALUE'       , ReadTokenSt(sString));
      TheTob.PutValue('WIL_C2BRANCHEMENT' , ReadTokenSt(sString));
      TheTob.PutValue('WIL_C2OPERATEUR'   , ReadTokenSt(sString));
      TheTob.PutValue('WIL_C2VALUE'       , ReadTokenSt(sString));
      TheTob.PutValue('WIL_C3BRANCHEMENT' , ReadTokenSt(sString));
      TheTob.PutValue('WIL_C3OPERATEUR'   , ReadTokenSt(sString));
      TheTob.PutValue('WIL_C3VALUE'       , ReadTokenSt(sString));
      TheTob.PutValue('WIL_CUSTOMIZABLE'  , ReadTokenSt(sString));
      TheTob.PutValue('WIL_DEFINITVALEUR' , ReadTokenSt(sString));
      TheTob.PutValue('WIL_FAMCHAMP'      , ReadTokenSt(sString));
      TheTob.PutValue('WIL_IDENTIFIANT'   , ValeurI(ReadTokenSt(sString)));
      TheTob.PutValue('WIL_INITCTXCODE'   , ReadTokenSt(sString));
      TheTob.PutValue('WIL_INITCTXTYPE'   , ReadTokenSt(sString));
      TheTob.PutValue('WIL_INITVALEUR'    , ReadTokenSt(sString));
      TheTob.PutValue('WIL_NOMCHAMP'      , ReadTokenSt(sString));
      TheTob.PutValue('WIL_NOMTABLE'      , ReadTokenSt(sString));
      TheTob.PutValue('WIL_NUMORDRE'      , ValeurI(ReadTokenSt(sString)));
      TheTob.PutValue('WIL_NUMVERSIONTET' , ValeurI(ReadTokenSt(sString)));
      TheTob.PutValue('WIL_OBLIGATOIRE'   , ReadTokenSt(sString));
      TheTob.PutValue('WIL_PRIORITE'      , ReadTokenSt(sString));
      TheTob.InsertOrUpdateDB;
    finally
      TheTob.Free;
    end;
  end;

begin
  BeginTrans ;

  ExecuteSql('UPDATE ADRESSES SET ADR_DELAIMOYEN=0, ADR_EAN="", ADR_INCOTERM="", ADR_MODEEXP="", ADR_LIEUDISPO="", ADR_NIF="", ADR_EMAIL="", ADR_NUMEROCONTACT=0, ADR_REGION="", ADR_LIVR="-", ADR_FACT="-", ADR_REGL="-", ADR_NADRESSE=0');

  //-- Table ADRESSES --//
  MaJ_ADR_NADRESSE;
  // Adresses : Mise à jour de la nature d'auxiliaire

  ExecuteSQL('UPDATE ADRESSES'
             + ' SET ADR_NATUREAUXI= (SELECT MIN(T_NATUREAUXI) FROM TIERS WHERE (T_TIERS=ADRESSES.ADR_REFCODE))'
             + ' WHERE (ADR_TYPEADRESSE="TIE")'
              );
  ExecuteSql('UPDATE ADRESSES SET ADR_LIVR="X" WHERE (ADR_TYPEADRESSE="TIE")');
  ExecuteSQL('UPDATE ADRESSES SET ADR_NATUREAUXI= "" WHERE ADR_NATUREAUXI IS NULL');

  //-- Table ARTICLE --//
  ExecuteSQL('UPDATE ARTICLE SET GA_ARTICLEPROFIL= "",GA_EAN="" WHERE GA_ARTICLEPROFIL IS NULL');

  //-- Table CONTACT --//
  ExecuteSQL('UPDATE CONTACT SET C_TIERS=(SELECT T_TIERS FROM TIERS WHERE (T_NATUREAUXI=CONTACT.C_NATUREAUXI) and (T_AUXILIAIRE=CONTACT.C_AUXILIAIRE)) WHERE C_TYPECONTACT="T"');
  ExecuteSQL('UPDATE CONTACT SET C_TIERS = "" WHERE (C_TIERS is NULL)');

  //-- Table DEPOT --//
  ExecuteSQL('UPDATE DEPOTS SET GDE_TIERS="" WHERE GDE_TIERS IS NULL');

  //-- Table PAYS --//
  ExecuteSQL('UPDATE PAYS SET PY_LIEUDISPO="002" WHERE (PY_LIEUDISPO IS NULL)AND(PY_MEMBRECEE="X")');
  ExecuteSQL('UPDATE PAYS SET PY_LIEUDISPO="003" WHERE (PY_LIEUDISPO IS NULL)AND(PY_MEMBRECEE="-")');
  ExecuteSQL('UPDATE PAYS SET PY_LIEUDISPO="001" WHERE PY_PAYS="FRA"');

  //-- Table PIEDPORT --//
  ExecuteSQL('UPDATE PIEDPORT SET GPT_BLOQUEFRAIS= "VER" WHERE GPT_BLOQUEFRAIS IS NULL');

  //-- Table PORT --//
  ExecuteSQL('UPDATE PORT SET GPO_VERROU="VER" WHERE GPO_VERROU IS NULL');

  //-- Table PROFILART --//
  ExecuteSQL('UPDATE PROFILART SET GPF_EAN="", GPF_TYPEPROFILART="" WHERE GPF_EAN IS NULL');

  //-- Table TIERS --//
  ExecuteSQL('UPDATE TIERS SET T_REGION="" WHERE T_REGION IS NULL');

  //-- Table TIERSCOMPL --//
  sSQL :='UPDATE TIERSCOMPL SET	'+
         'YTC_REPRESENTANT2="",YTC_REPRESENTANT3="", '+           //YTC_NUMEROCONTACT=0,
         'YTC_TAUXREPR1=0,YTC_TAUXREPR2=0,YTC_TAUXREPR3=0,YTC_TARIFSPECIAL="", '+
         'YTC_INCOTERM="",YTC_MODEEXP="",YTC_LIEUDISPO="",YTC_TIERSLIVRE="", '+
         'YTC_NADRESSELIV=0,YTC_NADRESSEFAC=0,YTC_STATIONEDI="",YTC_NOTRECODETIERS="", '+
         'YTC_NOTRECODCOMPTA="",YTC_COMMSPECIAL="" WHERE YTC_REPRESENTANT2 IS NULL';
  ExecuteSQL(sSQL);

  //-- Table PIECEADRESSES --//
  sSQL := 'UPDATE PIECEADRESSE SET GPA_EAN="",'+
          'GPA_INCOTERM="",GPA_LIEUDISPO="",GPA_MODEEXP="",' +
          'GPA_NIF="",GPA_NUMEROCONTACT=0,GPA_REGION="" ' +
          'WHERE GPA_REGION IS NULL';
  ExecuteSQL(sSQL);

  //-- Table des codes postaux --//
  sSQL := 'UPDATE CODEPOST SET O_PAYS="" '+
          'WHERE O_PAYS IS NULL';
  ExecuteSQL(sSQL);

  //-- Création de l'élément ST dans la table QSTRATCA si celle-ci est vide --//
  if not ExisteSQL('SELECT QST_CTX FROM QSTRATCA WHERE QST_CTX="0"')then
  begin
    ExecuteSQL('INSERT INTO QSTRATCA (QST_CTX,QST_STRATECA,QST_LIBSTRATECA,' +
      'QST_CRITCALREG,QST_CRITCALPTREG,QST_CRITCALPOLY,QST_CRITCALPPOLY,' +
      'QST_CRITCALCOMP,QST_CRITCALPCOMP,QST_ORDRECRIT,QST_CRITCALLMIN,' +
      'QST_CRITCPTREG) VALUES ("0","ST","Stratégie standard","5","10","' +
      '5","10","5","10","132","-","7")');
  end;

  //-- Table des paramètre GPAO --//
  sSQL := 'UPDATE WPARAM SET WPA_VARCHAR11="",WPA_VARCHAR12="",WPA_VARCHAR13="",WPA_VARCHAR14="",WPA_VARCHAR15=""'+
          ' WHERE WPA_VARCHAR11 IS NULL';
  ExecuteSQL(sSQL);
  
  //-------------- Tablettes ----------------------//

  //-- AFTTYPERESSOURCE --//
  InsertChoixCode ('TRE','OUT','Outils','Outils','TYPEFICHE=M;FICHE=WRESSOURCE_FIC;NATUREFICHE=W;IMGNUM=40');
  InsertChoixCode ('TRE','MAT','Matériels','Matériel','TYPEFICHE=M;FICHE=WRESSOURCE_FIC;NATUREFICHE=W;IMGNUM=39');

  UpdateChoixCodeLibre('TRE','AUT','FICHE=WRESSOURCE_FIC;NATUREFICHE=W');
  UpdateChoixCodeLibre('TRE','INT','TYPEFICHE=P;FICHE=WRESSOURCEPER_FIC;NATUREFICHE=W;IMGNUM=80');
  UpdateChoixCodeLibre('TRE','LOC','TYPEFICHE=M;FICHE=WRESSOURCE_FIC;NATUREFICHE=W;IMGNUM=14');
  UpdateChoixCodeLibre('TRE','MAC','TYPEFICHE=M;FICHE=WRESSOURCE_FIC;NATUREFICHE=W;IMGNUM=38');
  UpdateChoixCodeLibre('TRE','SAL','TYPEFICHE=P;FICHE=WRESSOURCEPER_FIC;NATUREFICHE=W;IMGNUM=81');
  UpdateChoixCodeLibre('TRE','ST','TYPEFICHE=P;FICHE=WRESSOURCEPER_FIC;NATUREFICHE=W;IMGNUM=74');

  //-- GCINCOTERM --//
  InsertChoixCode ('INC','CFR','Coût de frêt (C&F)','','');
  InsertChoixCode ('INC','CIF','Coût, assurance et frêt (CAF)','','');
  InsertChoixCode ('INC','CIP','Port payé, assurance comprise jusqu''à ...','','');
  InsertChoixCode ('INC','CPT','Port payé jusqu''à ...','','');
  InsertChoixCode ('INC','DAF','Rendu frontière','','');
  InsertChoixCode ('INC','DDP','Rendus droits acquittés','','');
  InsertChoixCode ('INC','DDU','Rendu droits non acquittés','','');
  InsertChoixCode ('INC','DEQ','Rendu à quai','','');
  InsertChoixCode ('INC','DES','Rendu "ex ship"','','');
  InsertChoixCode ('INC','EXW','A l''usine','','');
  InsertChoixCode ('INC','FAS','Franco le long du navire','','');
  InsertChoixCode ('INC','FCA','Franco transport','','');
  InsertChoixCode ('INC','FOB','Franco de bord','','');

  //-- GCMODEEXP --//
  InsertChoixCode ('MEX','1','MARITIME','','');
  InsertChoixCode ('MEX','2','CHEMIN de FER','','');
  InsertChoixCode ('MEX','3','ROUTE','','');
  InsertChoixCode ('MEX','4','AIR','','');
  InsertChoixCode ('MEX','5','ENVOIS POSTAUX','','');
  InsertChoixCode ('MEX','7','INSTALLATIONS DE TRANSPORTS FIXES','','');
  InsertChoixCode ('MEX','8','NAVIGATION INTERIEURE','','');
  InsertChoixCode ('MEX','9','PROPULSION PROPRE','','');

  //-- YTARIFSOU --//
  InsertChoixCode ('YTU','C','à la commande','','');
  InsertChoixCode ('YTU','L','à la ligne','','');

  //-- YTARIFSQUOI --//
  InsertChoixCode ('YTN','L','Des linéaires','Linéaires','');
  InsertChoixCode ('YTN','P','Des poids','Poids','');
  InsertChoixCode ('YTN','Q','Des quantités','Quantités','');
  InsertChoixCode ('YTN','S','Des surfaces','Surfaces','');
  InsertChoixCode ('YTN','V','Des volumes','Volumes','');

  //-- YTARIFSRECHERCHE --//
  InsertChoixCode ('YTR','C','à la commande','','');
  InsertChoixCode ('YTR','L','à la ligne','','');

  //-------------- Ajout dans les tables ----------//

  //-- WINITCHAMPTET --//
  TheTob := Tob.Create('WINITCHAMPTET', Nil, -1);
  try
    TheTob.PutValue('WIT_AVECCONDITION', '-');
    TheTob.PutValue('WIT_BLOCNOTE',''       );
    TheTob.PutValue('WIT_C1LIBELLE',''      );
    TheTob.PutValue('WIT_C1NOMCHAMP',''     );
    TheTob.PutValue('WIT_C2LIBELLE',''      );
    TheTob.PutValue('WIT_C2NOMCHAMP',''     );
    TheTob.PutValue('WIT_C3LIBELLE',''      );
    TheTob.PutValue('WIT_C3NOMCHAMP',''     );
    TheTob.PutValue('WIT_CONTEXTEPERE',''   );
    TheTob.PutValue('WIT_INITCTXCODE','DEFAUT'  );
    TheTob.PutValue('WIT_INITCTXTYPE','DEF'     );
    TheTob.PutValue('WIT_LIBELLE','Initialisations et champs obligatoire de la table ARTICLE' );
    TheTob.PutValue('WIT_NBRCONDITIONS',0   );
    TheTob.PutValue('WIT_NOMTABLE','ARTICLE');
    TheTob.PutValue('WIT_NUMVERSION',1      );
    TheTob.PutValue('WIT_STANDARD','X'      );
    TheTob.InsertOrUpdateDB;
  finally
    TheTob.Free;
  end;

  //-- WINITCHAMPLIG --//
  MaJWIL('-;;;AND;;;AND;;;X;MRP   ;DIV;41 ;DEFAUT;DEF;MRP   ;GA_METHACH         ;ARTICLE;14 ;1;X');
  MaJWIL('-;;;AND;;;AND;;;X;MRP   ;DIV;42 ;DEFAUT;DEF;MRP   ;GA_METHPROD        ;ARTICLE;15 ;1;X');
  MaJWIL('-;;;AND;;;AND;;;X;1,000 ;DIV;1  ;DEFAUT;DEF;1,000 ;GA_PCB             ;ARTICLE;1  ;1;-');
  MaJWIL('-;;;AND;;;AND;;;X;1,000 ;DIV;2  ;DEFAUT;DEF;1,000 ;GA_QPCBACH         ;ARTICLE;2  ;1;-');
  MaJWIL('-;;;AND;;;AND;;;X;H     ;UNI;3  ;DEFAUT;DEF;H     ;GA_QUALIFHEURE     ;ARTICLE;3  ;1;-');
  MaJWIL('-;;;AND;;;AND;;;-;M     ;UNI;4  ;DEFAUT;DEF;M     ;GA_QUALIFLINEAIRE  ;ARTICLE;4  ;1;-');
  MaJWIL('-;;;AND;;;AND;;;-;K     ;UNI;5  ;DEFAUT;DEF;K     ;GA_QUALIFPOIDS     ;ARTICLE;5  ;1;-');
  MaJWIL('-;;;AND;;;AND;;;-;M2    ;UNI;6  ;DEFAUT;DEF;M2    ;GA_QUALIFSURFACE   ;ARTICLE;6  ;1;-');
  MaJWIL('-;;;AND;;;AND;;;X;UNI   ;UNI;7  ;DEFAUT;DEF;UNI   ;GA_QUALIFUNITESTO  ;ARTICLE;7  ;1;X');
  MaJWIL('-;;;AND;;;AND;;;X;UNI   ;UNI;8  ;DEFAUT;DEF;UNI   ;GA_QUALIFUNITEVTE  ;ARTICLE;8  ;1;X');
  MaJWIL('-;;;AND;;;AND;;;X;M3    ;UNI;9  ;DEFAUT;DEF;M3    ;GA_QUALIFVOLUME    ;ARTICLE;9  ;1;-');
  MaJWIL('-;;;AND;;;AND;;;X;UNI   ;UNI;10 ;DEFAUT;DEF;UNI   ;GA_UNITECONSO      ;ARTICLE;10 ;1;X');
  MaJWIL('-;;;AND;;;AND;;;X;UNI   ;UNI;11 ;DEFAUT;DEF;UNI   ;GA_UNITEPROD       ;ARTICLE;11 ;1;X');
  MaJWIL('-;;;AND;;;AND;;;X;UNI   ;UNI;12 ;DEFAUT;DEF;UNI   ;GA_UNITEQTEACH     ;ARTICLE;12 ;1;X');
  MaJWIL('-;;;AND;;;AND;;;X;UNI   ;UNI;13 ;DEFAUT;DEF;UNI   ;GA_UNITEQTEVTE     ;ARTICLE;13 ;1;X');

  CommitTrans;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 27/06/2003
Modifié le ... : 27/06/2003
Description .. : Moulinette pour la fusion des champs date et des champs
Suite ........ : heure de la table juevenement :
Suite ........ : JEV_DATE et JEV_HEUREDEB
Suite ........ : JEV_DATEFIN et JEV_HEUREFIN
Suite ........ : JEV_ALERTEDATE et JEV_ALERTEHEURE
Suite ........ : avant suppresion des champs JEV_HEUREDEB,
Suite ........ : JEV_HEUREFIN, JEV_ALERTEHEURE
Mots clefs ... :
*****************************************************************}
procedure ConvertirUneDateHeure( dDate_p : TDate; tHeure_p : TTime; var dtDateHeure_p : TDateTime );
begin
   if ( dDate_p = iDate1900 ) then
      dtDateHeure_p := dDate_p
   else if ( dDate_p <> iDate1900 ) and ( Length( DateTimeToStr( dDate_p ) ) = 19 ) then
      dtDateHeure_p := dDate_p         // Déjà mis à jour
   else if ( dDate_p <> iDate1900 ) and ( tHeure_p = iDate1900) then
      dtDateHeure_p := dDate_p
   else if ( dDate_p <> iDate1900 ) and ( tHeure_p <> iDate1900) then
      dtDateHeure_p := dDate_p + tHeure_p;
end;

procedure ConvertirLesDatesHeures( OBEvenement_p : TOB );
var
   dtDateHeureDeb_l, dtDateHeureFin_l, dtDateHeureAlerte_l : TDateTime;
   dDateDeb_l, dDateFin_l, dDateAlerte_l : TDate;
   tHeureDeb_l, tHeureFin_l, tHeureAlerte_l : TTime;
begin
   dDateDeb_l := OBEvenement_p.GetValue('JEV_DATE');
   tHeureDeb_l := OBEvenement_p.GetValue('JEV_HEUREDEB');
   ConvertirUneDateHeure( dDateDeb_l, tHeureDeb_l, dtDateHeureDeb_l );

   dDateFin_l := OBEvenement_p.GetValue('JEV_DATEFIN');
   tHeureFin_l := OBEvenement_p.GetValue('JEV_HEUREFIN');
   ConvertirUneDateHeure( dDateFin_l, tHeureFin_l, dtDateHeureFin_l );

   dDateAlerte_l := OBEvenement_p.GetValue('JEV_ALERTEDATE');
   tHeureAlerte_l := OBEvenement_p.GetValue('JEV_ALERTEHEURE');
   ConvertirUneDateHeure( dDateAlerte_l, tHeureAlerte_l, dtDateHeureAlerte_l );

   OBEvenement_p.PutValue('JEV_DATE', dtDateHeureDeb_l );
   OBEvenement_p.PutValue('JEV_DATEFIN', dtDateHeureFin_l );
   OBEvenement_p.PutValue('JEV_ALERTEDATE', dtDateHeureAlerte_l );
   OBEvenement_p.PutValue('JEV_DUREE', dtDateHeureFin_l - dtDateHeureDeb_l );
end;

Procedure ConvertirDateHeureEvenement;
var
   OBEvenement_l : TOB;
   QRYEvenement_l : TQuery;
   nEvtInd_l, nEvtNb_l : integer;
begin
   OBEvenement_l := TOB.Create( 'JUEVENEMENT', nil, -1 );
   QRYEvenement_l := OpenSQL( 'select JEV_NOEVT, JEV_DATE, JEV_HEUREDEB, ' +
                              '       JEV_DATEFIN, JEV_HEUREFIN, ' +
                              '       JEV_ALERTEDATE, JEV_ALERTEHEURE, JEV_DUREE ' +
                              'from JUEVENEMENT ',
                              true );

   OBEvenement_l.LoadDetailDB('JUEVENEMENT', '', '', QRYEvenement_l, false );
   nEvtNb_l := OBEvenement_l.Detail.Count;
   For nEvtInd_l := 0 to nEvtNb_l - 1 do
   begin
      ConvertirLesDatesHeures( OBEvenement_l.Detail[nEvtInd_l] );
      OBEvenement_l.Detail[nEvtInd_l].UpdateDB;
   end;
   Ferme(QRYEvenement_l);
   OBEvenement_l.Free;
end;

/////////////////////////////////////////////////////////////////
procedure MAJF2725Pendant(NomTable:string);
var i : integer;
  Q : TQuery;
begin
if NomTable='F2725_SOC' then
begin
  // Table des societes (F2725_SOC)
  i := 1;
  Q := OpenSQL('select fss_n04558 from f2725_soc order by fss_n04550', false);
  while not Q.Eof do
  begin
    Q.Edit;
    Q.FindField('FSS_N04558').AsInteger := i;
    Q.Post;
    Q.Next;
    inc(i);
  end;
  Ferme(Q);
end;
  // Table des enfants (F2725_ENF)
if (NomTable = 'F2725_ENF') then
begin
  i := 1;
  Q := OpenSQL('select fse_n04512 from f2725_enf order by fse_n04500', false);
  while not Q.Eof do
  begin
    Q.Edit;
    Q.FindField('FSE_N04512').AsInteger := i;
    Q.Post;
    Q.Next;
    inc(i);
  end;
  Ferme(Q);
end;
end;
/////////////////////////////////////////////////////////////////

Procedure MajImoAvantPourSynchro ;
Var Q : tQuery ;
    GoSynchro : Boolean ;
BEGIN
GoSynchro:=TRUE ;
Q:=OpenSQL('SELECT DT_NUMVERSION from DETABLES WHERE DT_NOMTABLE="IMOREF" ',TRUE) ;
If Not Q.EOF Then GoSynchro:=Q.FindField('DT_NUMVERSION').AsInteger<=12 ;
Ferme(Q) ;
If Not GoSynchro Then Exit ;
TestDropTable('I2054') ;
TestDropTable('IAFFCOMPTA') ;
TestDropTable('IAIDSAI') ;
TestDropTable('ICODGEO') ;
TestDropTable('ICODP00') ;
TestDropTable('ICODSPC') ;
TestDropTable('ICPTECR') ;
TestDropTable('ICPTIMG') ;
TestDropTable('ICPTMET') ;
TestDropTable('IDATCLO') ;
TestDropTable('IDUREETAUX') ;
TestDropTable('IECRANA') ;
TestDropTable('IECRLGN') ;
TestDropTable('IMOASS') ;
TestDropTable('IMOAVANTVIR') ;
TestDropTable('IMOCODASS') ;
TestDropTable('IMOCODRVL') ;
TestDropTable('IMOCODSUB') ;
TestDropTable('IMOD205') ;
TestDropTable('IMODOT') ;
TestDropTable('IMODOT2') ;
TestDropTable('IMOECA') ;
TestDropTable('IMOFAC') ;
TestDropTable('IMOFIS') ;
TestDropTable('IMOHISTOPART') ;
TestDropTable('IMOLS0') ;
TestDropTable('IMOLS1') ;
TestDropTable('IMOLS2') ;
TestDropTable('IMOMET') ;
TestDropTable('IMOPART') ;
TestDropTable('IMOPER') ;
TestDropTable('IMOREF') ;
TestDropTable('IMOREFEDI') ;
TestDropTable('IMOREP') ;
TestDropTable('IMORVL') ;
TestDropTable('IMOSO1') ;
TestDropTable('IMOSO2') ;
TestDropTable('IMOSUB') ;
TestDropTable('IMOVEN') ;
TestDropTable('IMOVI1') ;
TestDropTable('IMOVI2') ;
TestDropTable('IPERCAL2') ;
TestDropTable('IUTPARAM') ;
END ;

procedure RTMajDureeAction;
var Q    : TQuery ;
    Hour,Min,Sec,Msec : Word;
begin
    Q := opensql('SELECT * FROM ACTIONS where RAC_DUREEACT<>"01/01/1900" and RAC_DUREEACT<>"12/30/1899"',False) ;
    While not Q.eof do
    begin
      Q.Edit;
      DecodeTime(Q.FindField('RAC_DUREEACT').AsDateTime, Hour, Min, Sec, MSec);
      Q.FindField('RAC_DUREEACTION').AsFloat := (Hour*60)+Min;
      Q.Post;
      Q.Next;
    end ;
    Ferme(Q) ;
end;


Procedure Init_WInitChampLig;
var
  TobWIL: Tob;
  QWIL: TQuery;
begin
ExecuteSQL('UPDATE WINITCHAMPLIG SET WIL_IDENTIFIANT=(SELECT MAX(WIL_IDENTIFIANT) + 1 FROM WINITCHAMPLIG) WHERE WIL_IDENTIFIANT=41');
ExecuteSQL('UPDATE WINITCHAMPLIG SET WIL_IDENTIFIANT=(SELECT MAX(WIL_IDENTIFIANT) + 1 FROM WINITCHAMPLIG) WHERE WIL_IDENTIFIANT=42');
TobWIL := Tob.Create('WJETONS', nil, -1);
try
  TobWIL.PutValue('WJT_PREFIXE', 'WJT');
  TobWIL.PutValue('WJT_CONTEXTE', '');
  QWIL := OpenSQL('SELECT MAX(WIL_IDENTIFIANT) + 1 FROM WINITCHAMPLIG', true, 1);
  try
    if not QWIL.Eof then
      TobWIL.PutValue('WJT_JETON', QWIL.Fields[0].AsInteger)
    else
      TobWIL.PutValue('WJT_JETON', 1);
  finally
    Ferme(QWIL);
  end;
  TobWIL.InsertOrUpdateDB;
finally
  TobWIL.Free;
end;
end;

Procedure GPAOMajApres_617;
      { Mise à jour d'enregistrement dans la table MEA }
      procedure pUpdateMEA(sQualif, sMesure, sLibelle, dQuotite : string);
      var
         sSQL : string;
      begin
         sSQL := 'SELECT * FROM MEA ' +
                 'WHERE GME_MESURE = "' + sMesure + '" ';
         if not ExisteSQL(sSQL) then
            ExecuteSQL('UPDATE MEA SET GME_QUOTITE ="'+ dQuotite +'" WHERE GME_QUALIFMESURE="'+sQualif+'" AND GME_MESURE="'+sMesure+'"');
      end;

Begin
//-- MAJ TABLE QPHASE --//
if not ExisteSQL('SELECT QPH_PHASE FROM QPHASE WHERE QPH_PHASE="NUL"') then
  ExecuteSQL('INSERT INTO QPHASE (QPH_CTX,QPH_PHASE,QPH_PHASELIB,QPH_PHASECRI,QPH_COCHENEGOCE, QPH_DATEMODIF)'
           + ' VALUES ("0","NUL","Phase nulle","-","-","'+ UsDateTime(iDate1900)  +'")');

//-- MAJ TABLE QITI --//
if not ExisteSQL('SELECT QIT_CODITI FROM QITI WHERE QIT_CODITI="NUL"') then
  ExecuteSQL('INSERT INTO QITI (QIT_CTX,QIT_CODITI,QIT_CODITILIB,QIT_ITISOUM,QIT_COSOUM,QIT_DATEMODIF)'
           + ' VALUES ("0","NUL","Itinéraire nul","","-","'+ UsDateTime(iDate1900)  +'")');

//-- MAJ TABLE QPHASEITI --//
if not ExisteSQL('SELECT QP_CODITI FROM QPHASEITI WHERE QP_CODITI="NUL"') then
  ExecuteSQL('INSERT INTO QPHASEITI (QP_CTX,QP_CODITI,QP_OPEITI,QP_PHASE,QP_OPEITILIB,QP_DATEMODIF)'
           + ' VALUES ("0","NUL","110","NUL","Phase nulle","'+ UsDateTime(iDate1900)  +'")');

//-- MAJ TABLE QCIRCUIT --//
if not ExisteSQL('SELECT QCI_CIRCUIT FROM QCIRCUIT WHERE QCI_CIRCUIT="NUL"') then
  ExecuteSQL('INSERT INTO QCIRCUIT (QCI_CTX,QCI_CIRCUIT,QCI_CIRCUITLIB,QCI_CODITI,QCI_PRIVILEGIE,QCI_DATEMODIF)'
           + ' VALUES ("0","NUL","Circuit nul","NUL","-","'+ UsDateTime(iDate1900)  +'")');

//-- MAJ TABLE QDETCIRC --//
if not ExisteSQL('SELECT QDE_CIRCUIT FROM QDETCIRC WHERE QDE_CIRCUIT="NUL"') then
  ExecuteSQL('INSERT INTO QDETCIRC (QDE_CTX,QDE_CIRCUIT,QDE_OPECIRC,QDE_JALON,QDE_POLE,QDE_SITE,QDE_GRP,QDE_TYPTRANS,QDE_JALONMAX,QDE_JALONCOURT,QDE_TYPTRANSCOURT,QDE_DATEMODIF)'
           + ' VALUES ("0","NUL","110","0","","","","","0","0","","'+ UsDateTime(iDate1900)  +'")');

//-- Tablette AFTTYPERESSOURCE --//
UpdateChoixCodeLibre('TRE','OUT','TYPEFICHE=M;FICHE=WRESSOURCE_FIC;NATUREFICHE=W;IMGNUM=40');
UpdateChoixCodeLibre('TRE','MAT','TYPEFICHE=M;FICHE=WRESSOURCE_FIC;NATUREFICHE=W;IMGNUM=39');

//-- Mise à jour pour les tarifs --//
// Table supprimée en 644 ExecuteSQL('UPDATE YTARIFSFONCTIONNALITE SET YFO_GRAPHIQUE="75" WHERE (YFO_FONCTIONNALITE="201") or (YFO_FONCTIONNALITE="202")');
// ExecuteSQL('UPDATE YTARIFSFONCTIONNALITE SET YFO_GRAPHIQUE="76" WHERE (YFO_FONCTIONNALITE="101") or (YFO_FONCTIONNALITE="301") or (YFO_FONCTIONNALITE="401")');
ExecuteSQL('UPDATE PARPIECE SET GPP_TARIFMODULE="201" WHERE (GPP_VENTEACHAT="VEN")');
ExecuteSQL('UPDATE PARPIECE SET GPP_TARIFMODULE="101" WHERE (GPP_VENTEACHAT="ACH") and (GPP_PILOTEORDRE="-")');
ExecuteSQL('UPDATE PARPIECE SET GPP_TARIFMODULE="301" WHERE (GPP_VENTEACHAT="ACH") and (GPP_PILOTEORDRE="X")');

  //-- Table "MEA" --//
  pUpdateMEA('TEM','H','Heure','3600');
  pUpdateMEA('TEM','J','Jour','86400');
  pUpdateMEA('TEM','MIN','Minute','60');
  pUpdateMEA('TEM','SEC','Seconde','1');
  pUpdateMEA('THC','HC','Heure centième','3600');

//-- AJOUT DE DONNEES DANS LA TABLE EDIFAMILLEEMG --//
if not ExisteSQL('SELECT EFM_CODEFAMILLE FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="ARC"') then
  ExecuteSQL('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE, EFM_NATURESPIECE, EFM_SERIALISE) '
            +'VALUES("ARC", "Accusé de réception de commande EDI_ARC", "ACV,CC,CCE,CCR", "-")');

if not ExisteSQL('SELECT EFM_CODEFAMILLE FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="BL"') then
  ExecuteSQL('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE, EFM_NATURESPIECE, EFM_SERIALISE) '
            +'VALUES("BL", "Bon de livraison standard EDI_BL", "ALF,ALV,APV,BLC,LCE,LCR,LFR,PRE", "-")');

if not ExisteSQL('SELECT EFM_CODEFAMILLE FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="CDE"') then
  ExecuteSQL('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE, EFM_NATURESPIECE, EFM_SERIALISE) '
            +'VALUES("CDE", "Commande standard EDI_CDE", "ACV,CC,CCE,CCR", "-")');

if not ExisteSQL('SELECT EFM_CODEFAMILLE FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="FAC"') then
  ExecuteSQL('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE, EFM_NATURESPIECE, EFM_SERIALISE) '
            +'VALUES("FAC", "Facture EDI_FAC", "ABT,AF,AFP,AFS,APR,AVC,AVS,FAC,FBT", "-")');

//-- MAJ TABLE WCUQUALIFIANT --//
ExecuteSQL('UPDATE WCUQUALIFIANT SET WDB_OPERATEUR="*" WHERE WDB_QUALIFIANT IN ("QTS","QLS")');

//-- MAJ TABLE WFORMULECHAMP --//
ExecuteSQL('UPDATE WFORMULECHAMP SET WDX_CHAMPSAIS="WOB_QTSTSAIS", WDX_UNITECHAMPSAIS="WOB_UNITELIEN" '
          +'                        ,WDX_CHAMPCOEF="WOB_COEFLIEN", WDX_CHAMPCALC="WOB_QTSTSTOC" '
          +'                        ,WDX_UNITECHAMPCALC="WOB_QUALIFUNITESTO" '
          +' WHERE WDX_CONTEXTE="WOB" AND WDX_QUALIFIANT = "QTS" ');

End;


procedure CorrigeRTTiers ;
var
  st,NomVue : string;
  QSoc: TQuery;
begin
  Nomvue:='RTTIERS';
  if not TableExiste('DEVUES') then exit;
  QSoc := OpenSQL('SELECT DV_NOMVUE, DV_NUMVERSION, DV_SQL FROM DEVUES where DV_NOMVUE="'+NomVue+'"', False);
  if not QSoc.EOF then
  begin
    QSoc.edit;
    st := QSoc.Fields[2].AsString;
    st:=FindEtReplace(st,'T_ZONECOM'+#13#10+'T_FACTURE','T_ZONECOM,'+#13#10+' T_FACTURE',True );
    QSoc.Fields[2].AsString:=st;
    QSoc.post;
    DropVuePourrie(NomVue);
    DBCreateView(DBSOC, NomVue, St, V_PGI.Driver);
  end;
  Ferme(QSoc);
end;


procedure MajFluxTreso;
var
  TobTreso, TT: Tob;
begin

  TobTreso := Tob.Create('DetailFluxTreso', nil, -1);
  try
    TT := Tob.Create('FLUXTRESO', TobTreso, -1);
    TT.PutValue('TFT_FLUX', 'EQD');
    TT.PutValue('TFT_TYPEFLUX', 'ED');
    TT.PutValue('TFT_LIBELLE', 'Equilibrage dépense');
    TT.PutValue('TFT_ABREGE', 'Equilibrage dép');

    TT := Tob.Create('FLUXTRESO', TobTreso, -1);
    TT.PutValue('TFT_FLUX', 'EQR');
    TT.PutValue('TFT_TYPEFLUX', 'ER');
    TT.PutValue('TFT_LIBELLE', 'Equilibrage recette');
    TT.PutValue('TFT_ABREGE', 'Equilibrage rec');

    TT := Tob.Create('FLUXTRESO', TobTreso, -1);
    TT.PutValue('TFT_FLUX', 'REI');
    TT.PutValue('TFT_TYPEFLUX', 'REI');
    TT.PutValue('TFT_LIBELLE', 'Réinitialisation');
    TT.PutValue('TFT_ABREGE', 'Réinitialisation');

    TobTreso.InsertOrUpdateDB(False);
   finally
     TobTreso.ClearDetail;
     TobTreso.free;
   end;

  TobTreso := Tob.Create('DetailTYPEFLUX', nil, -1);
  try
    TT := Tob.Create('TYPEFLUX', TobTreso, -1);
    TT.PutValue('TTL_TYPEFLUX', 'ED');
    TT.PutValue('TTL_LIBELLE', 'Equilibrage dépense');
    TT.PutValue('TTL_ABREGE', 'Equilibrage dép');
    TT.PutValue('TTL_SENS', 'D');

    TT := Tob.Create('TYPEFLUX', TobTreso, -1);
    TT.PutValue('TTL_TYPEFLUX', 'ER');
    TT.PutValue('TTL_LIBELLE', 'Equilibrage recette');
    TT.PutValue('TTL_ABREGE', 'Equilibrage rec');
    TT.PutValue('TTL_SENS', 'C');

    TT := Tob.Create('TYPEFLUX', TobTreso, -1);
    TT.PutValue('TTL_TYPEFLUX', 'REI');
    TT.PutValue('TTL_LIBELLE', 'Réinitialisation');
    TT.PutValue('TTL_ABREGE', 'Réinitialisation');
    TT.PutValue('TTL_SENS', 'D');

    TobTreso.InsertOrUpdateDB(False);
   finally
     TobTreso.ClearDetail;
     TobTreso.free;
   end;

end;

procedure RT_InsertLibelleInfoComplContact;
var i: integer;
begin
  for i := 0 to 4 do InsertChoixCode('R6Z', 'BL' + intTostr(i), 'Décision libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R6Z', 'DL' + intTostr(i), 'Date libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R6Z', 'CL' + intTostr(i), 'Table libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R6Z', 'TL' + intTostr(i), 'Texte libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R6Z', 'VL' + intTostr(i), 'Valeur libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R6Z', 'ML' + intTostr(i), 'Multi-choix libre ' + intTostr(i + 1), '', '');
  InsertChoixCode('R6Z', 'BLO', 'Bloc-notes ', '', '');
end;

procedure W_SAV_COMPL;
var i: integer;
begin
  for i := 0 to 4 do InsertChoixCode('R5Z', 'BL' + intTostr(i), 'Décision libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R5Z', 'DL' + intTostr(i), 'Date libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R5Z', 'CL' + intTostr(i), 'Table libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R5Z', 'TL' + intTostr(i), 'Texte libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R5Z', 'VL' + intTostr(i), 'Valeur libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R5Z', 'ML' + intTostr(i), 'Multi-choix libre ' + intTostr(i + 1), '', '');
end;

// Procédure existante à modifiée :
procedure W_SAV_COMPETENCE;
var i: integer;
begin
  for i := 1 to 3 do InsertChoixCode('ZLI', 'CP' + intTostr(i), 'Table libre ' + intTostr(i + 1), '', '');
  for i := 1 to 5 do InsertChoixCode('ZLI', 'RH' + intTostr(i), 'Table libre ' + intTostr(i + 1), '', '');
end;


// fct qui passe les champ Jour* de Tache et TAchemodele dans une nouvelle table
// en janvier 2004 (5.0) aucun client n'utilisait ces tables
//risque minime que des cleints GA démarre sur planning d'ici 10/2004, donc transfert de table fait
//mais ne devrait concerné que très peu de client
procedure AlimTob1 ( QQ : Tquery;NoJour : integer);
var tobdet : Tob;
begin
Tobdet := Tob.Create ('',TobTacheJour,-1);
TOBDet.AddChampSupValeur('Affaire',QQ.Findfield('ATA_AFFAIRE').AsString,False) ;
TOBdet.AddChampSupValeur('NumeroTache',QQ.Findfield('ATA_NumeroTache').AsInteger,False) ;
TOBDet.AddChampSupValeur('JourAplanif',noJour,False) ;
TOBDet.AddChampSupValeur('ModeleTache','',False) ;
TOBDet.AddChampSupValeur('TypeJour','AFF',False) ;
end;

procedure AlimTob2 ( QQ : Tquery;NoJour : integer);
Var tobdet : tob;
begin
Tobdet := Tob.Create ('',TobTacheJour,-1);
TOBDet.AddChampSupValeur('Affaire','',False) ;
TOBDet.AddChampSupValeur('NumeroTache',0,False) ;
TOBdet.AddChampSupValeur('JourAplanif',noJour,False) ;
TOBdet.AddChampSupValeur('ModeleTache',QQ.Findfield('AFM_ModeleTache').AsString,False) ;
TOBdet.AddChampSupValeur('TypeJour','MOD',False) ;
end;

procedure GIGATacheAvant;
var sql :string;
  QQ: TQuery;
begin
TobTacheJour:=Nil;
if (TableExiste('TACHE')) and (ChampPhysiqueExiste('TACHE', 'ATA_JOUR1')) then
begin
  TobTacheJour := Tob.create ('les tachesjour', nil, -1);
  Sql := 'SELECT ata_jour1,ata_jour2,ata_jour3,Ata_jour4,ata_jour5,ata_jour6,ata_jour7,ata_affaire,ata_numerotache '
           +' FROM TACHE WHERE ata_jour1="X" or ata_jour2="X" or ata_jour3="X" or '
           +' ata_jour4="X" or ata_jour5="X" or ata_jour6="X" or ata_jour7="X"';
  QQ := OpenSQL(Sql, True);
  while not QQ.EOF do
    begin
    If QQ.Findfield('ATA_JOUR1').AsString ='X' then AlimTob1( QQ,1);
    If QQ.Findfield('ATA_JOUR2').AsString ='X' then AlimTob1( QQ,2);
    If QQ.Findfield('ATA_JOUR3').AsString ='X' then AlimTob1( QQ,3);
    If QQ.Findfield('ATA_JOUR4').AsString ='X' then AlimTob1( QQ,4);
    If QQ.Findfield('ATA_JOUR5').AsString ='X' then AlimTob1( QQ,5);
    If QQ.Findfield('ATA_JOUR6').AsString ='X' then AlimTob1( QQ,6);
    If QQ.Findfield('ATA_JOUR7').AsString ='X' then AlimTob1( QQ,7);
    QQ.next;   //peu d'enrgt et de bases concernées... onprend le risque
    end;
  ferme(QQ);
End;
  // la table n'existait pas avant socref 615 ...
if (TableExiste('AFMODELETACHE')) and (ChampPhysiqueExiste('AFMODELETACHE', 'AFM_JOUR1')) then
//if Not ExisteSql ('select dt_nomtable from detables where dt_nomtable="afmodeletache"') then exit;
begin
  Sql := 'SELECT afm_jour1,afm_jour2,afm_jour3,Afm_jour4,afm_jour5,afm_jour6,'
      +'afm_jour7,afm_modeletache FROM AFMODELETACHE WHERE afm_jour1="X" '
      + 'or afm_jour2="X" or afm_jour3="X" or afm_jour4="X" or afm_jour5="X" '
      + 'or afm_jour6="X" or afm_jour7="X"';
  QQ := OpenSQL(Sql, True);
  while not QQ.EOF do
    begin
    If QQ.Findfield('AFM_JOUR1').AsString ='X' then AlimTob2( QQ,1);
    If QQ.Findfield('AFM_JOUR2').AsString ='X' then AlimTob2( QQ,2);
    If QQ.Findfield('AFM_JOUR3').AsString ='X' then AlimTob2( QQ,3);
    If QQ.Findfield('AFM_JOUR4').AsString ='X' then AlimTob2( QQ,4);
    If QQ.Findfield('AFM_JOUR5').AsString ='X' then AlimTob2( QQ,5);
    If QQ.Findfield('AFM_JOUR6').AsString ='X' then AlimTob2( QQ,6);
    If QQ.Findfield('AFM_JOUR7').AsString ='X' then AlimTob2( QQ,7);
    QQ.next;   //peu d'enrgt et de bases concernées... onprend le risque
    end;
  ferme(QQ);
end;
End;

procedure GIGATacheApres;
var ii : integer;
TobDet : tob;
Sql : string;
begin     //ecriture depuis la tob, dans la nouvelle table AfTacheJour
If TobTacheJour=Nil then Exit; // en cas de deuxième passge
for ii :=0 to (TobTacheJour.detail.count -1) do
 begin
 TobDet :=TobTacheJOur.detail[ii];
 Sql := 'INSERT INTO AFTACHEJOUR (ATJ_JOURNUMERO,ATJ_TYPEJOUR,ATJ_AFFAIRE,ATJ_NUMEROTACHE,';
 Sql := Sql + 'ATJ_JOURAPLANIF,ATJ_JOURHEUREDEB,ATJ_JOURNBHEURES,ATJ_MODELETACHE)';
 Sql :=Sql + ' VALUES (' + IntTostr (ii+1)+',"'+ TobDet.getvalue('TYPEJOUR')+'",';
 Sql :=Sql + '"'+ TobDet.getvalue('Affaire')+'",';
 Sql :=Sql +  IntTostr(TobDet.getvalue('NumeroTAche'));
 Sql := Sql+ ','+IntTostr(TobDEt.getvalue('JourAplanif'));
 Sql :=Sql +',0,0,';
 Sql :=Sql + '"'+ TobDet.getvalue('ModeleTache')+'")';
 ExecuteSQL(Sql);
 end;
TobTacheJour.free;
end;



Procedure GIGAAvant644;
var QQ : TQuery;
Begin
  //GI/GA
  TOBAfPlanningEtat:=Nil;
  if TableExiste('AFPLANNING') and ChampPhysiqueExiste('AFPLANNING', 'APL_PREDEFMAJ') then
    ExecuteSql ('UPDATE AFPLANNING SET APL_ETATLIGNE="TER" WHERE APL_TERMINE="X"');
  if TableExiste('AFPLANNINGETAT')  then
  begin
    TOBAfPlanningEtat:=TOB.Create('la tob afplanningetat',Nil,-1) ;
    // 3 ou 4 enrgt maxi avec 4 zones.. on peut faire un SELECT *
    QQ:=OpenSQL('SELECT * FROM AFPLANNINGETAT',True) ;
    if not QQ.Eof then TobAfPlanningEtat.LoadDetailDB('La tob af','','',QQ, False) ;
    ferme (QQ);
  End;
  TestDropTable('PLANCHARGE');
  GiGaTacheAvant;
  if TableExiste('TACHE') and (tableToVersion( 'TACHE')< 115) then
     begin
     AddChamp('TACHE','ATA_ECARTQTEINIT','DOUBLE','ATA',False,0,'','',False) ;
     AddChamp('TACHE','ATA_ECARTQTEREF','DOUBLE','ATA',False,0,'','',False) ;
     AddChamp('TACHE','ATA_ECARTPR','DOUBLE','ATA',False,0,'','',False) ;
     AddChamp('TACHE','ATA_ECARTPRDEV','DOUBLE','ATA',False,0,'','',False) ;
     AddChamp('TACHE','ATA_ECARTPVHT','DOUBLE','ATA',False,0,'','',False) ;
     AddChamp('TACHE','ATA_ECARTPVHTDEV','DOUBLE','ATA',False,0,'','',False) ;
     if ChampPhysiqueExiste ( 'TACHE','ATA_ECARTQTEINIT' )  then
        if (tableToVersion( 'TACHE')> 107) then
         ExecuteSQL('UPDATE TACHE SET ATA_ECARTQTEINIT=ATA_QTEAPLANIFIER'
         + ' ,ATA_ECARTQTEREF=ATA_QTEAPLANIFUREF'
         + ' ,ATA_ECARTPR=ATA_RAPPTPR'
         + ' ,ATA_ECARTPRDEV=ATA_RAPPTPRDEV'
         + ' ,ATA_ECARTPVHT=ATA_RAPPTVENTEHT'
         + ' ,ATA_ECARTPVHTDEV=ATA_RAPPTVTDEVHT')
        else         ExecuteSQL('UPDATE TACHE SET ATA_ECARTQTEINIT=ATA_QTEAPLANIFIER'
         + ' ,ATA_ECARTQTEREF=0'
         + ' ,ATA_ECARTPR=ATA_RAPPTPR'
         + ' ,ATA_ECARTPRDEV=ATA_RAPPTPRDEV'
         + ' ,ATA_ECARTPVHT=ATA_RAPPTVENTEHT'
         + ' ,ATA_ECARTPVHTDEV=ATA_RAPPTVTDEVHT');

     end;
   if TableExiste('TACHERESSOURCE') and (tableToVersion( 'TACHERESSOURCE')< 104) then
     begin
     AddChamp('TACHERESSOURCE','ATR_ECARTQTEINIT','DOUBLE','ATR',False,0,'','',False) ;
     AddChamp('TACHERESSOURCE','ATR_ECARTQTEREF','DOUBLE','ATR',False,0,'','',False) ;
     AddChamp('TACHERESSOURCE','ATR_ECARTPR','DOUBLE','ATR',False,0,'','',False) ;
     AddChamp('TACHERESSOURCE','ATR_ECARTPRDEV','DOUBLE','ATR',False,0,'','',False) ;
     AddChamp('TACHERESSOURCE','ATR_ECARTPVHT','DOUBLE','ATR',False,0,'','',False) ;
     AddChamp('TACHERESSOURCE','ATR_ECARTPVHTDEV','DOUBLE','ATR',False,0,'','',False) ;
     if ChampPhysiqueExiste ( 'TACHERESSOURCE','ATR_ECARTQTEINIT' )  then
      if tableToVersion( 'TACHERESSOURCE')> 102 then
        ExecuteSQL('UPDATE TACHERESSOURCE SET ATR_ECARTQTEINIT=ATR_QTEAPLANIFIER'
         + ' ,ATR_ECARTQTEREF=ATR_QTEAPLANIFUREF'
         + ' ,ATR_ECARTPR=ATR_RAPPTPR'
         + ' ,ATR_ECARTPRDEV=ATR_RAPPTPRDEV'
         + ' ,ATR_ECARTPVHT=ATR_RAPPTVENTEHT'
         + ' ,ATR_ECARTPVHTDEV=ATR_RAPPTVTDEVHT')
      else       ExecuteSQL('UPDATE TACHERESSOURCE SET ATR_ECARTQTEINIT=ATR_QTEAPLANIFIER'
         + ' ,ATR_ECARTQTEREF=0'
         + ' ,ATR_ECARTPR=ATR_RAPPTPR'
         + ' ,ATR_ECARTPRDEV=ATR_RAPPTPRDEV'
         + ' ,ATR_ECARTPVHT=ATR_RAPPTVENTEHT'
         + ' ,ATR_ECARTPVHTDEV=ATR_RAPPTVTDEVHT');

     end;
      // on détruit les liste du planning pour les ramener correcte
  ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE IN ("AFPLANETATMUL","AFPLANPARAMMUL","AFTACHE_MUL","AFPLANETATMUL"'
   +',"AFLIGPLANNING","AFMULPLANNING","AFTACHEGEN","AFMULPLANNINGGENE","AFPLANNINGREVAL","AFTACHEREVAL_MUL","AFMULREVALRESS"'
   +',"AFMULARTICLEREVAL","AFTACHELIG")');
  ///fin GIGA
End;

Procedure GIGAAPres644;
var ii:integer;
 TobDet ,TobAfPlanningParam,Tobdet2: Tob;
 XX : string;
Begin
  // GI/GA
  // MCD 01122004 if GetParamsoc('SO_AFREALPLAN') then XX:='X' else XX:='-';
  if GetParamsoc('SO_AFGENELIMIT') then XX:='X' else XX:='-';
  ExecuteSql ('UPDATE TACHE SET ATA_FAITPARQUI="",ATA_RAPPLANNING="'+XX+'",ATA_ANNEENB=1');
  ExecuteSql ('UPDATE AFMODELETACHE SET AFM_ANNEENB=1,AFM_METHODEDECAL="P",AFM_NBJOURSDECAL=1');
  ExecuteSql ('UPDATE AFPLANNINGPARAM SET APP_TYPEAFPARAM="PLA" where  APP_TYPEAFPARAM is null');
  if GetParamSoc ('So_AFPLANDECHARGE') then
     begin
     ExecuteSql ('UPDATE TACHE SET ATA_QTEINITPLA=0, ATA_QTEINITREFPLA=0');
     ExecuteSql ('UPDATE TACHERESSOURCE SET ATR_QTEINITPLA=0, ATR_QTEINITREFPLA=0');
     ExecuteSql ('UPDATE AFPLANNING SET APL_TYPEPLA="PC" WHERE APL_TYPEPLA is null');
     end
     else begin
     ExecuteSql ('UPDATE TACHE SET ATA_QTEINITPLA=ATA_QTEINITIALE, ATA_QTEINITREFPLA=ATA_QTEINITUREF, ATA_QTEINITUREF=0, ATA_QTEINITIALE=0');
     ExecuteSql ('UPDATE TACHERESSOURCE SET ATR_QTEINITPLA=ATR_QTEINITIALE, ATR_QTEINITREFPLA=ATR_QTEINITUREF, ATR_QTEINITUREF=0, ATR_QTEINITIALE=0');
     ExecuteSql ('UPDATE AFPLANNING SET APL_TYPEPLA="PLA" WHERE APL_TYPEPLA is null');
     end;
  InsertChoixCode('AFQ', 'CAB', 'Cabinet', 'Cabinet', '');
  InsertChoixCode('AFQ', 'CLT', 'Client', 'Client', '');
  InsertChoixCode('AFQ', 'MIX', 'Mixte', 'Mixte', '');
   // il faut passer les info de la table AfPlanningEtat qui est détruite dans AfPlanningParam
  if  TobAfPlanningEtat<>nil then
  begin
    TOBAfPlanningParam:=TOB.Create('AFPLANNINGPARAM',Nil,-1) ;
    For ii:=0   to TobAfPlanningEtat.detail.count -1 do
    begin
      TobDEt := TobAfPlanningEtat.detail[ii];
      TobDet2 := TOB.Create('AFPLANNINGPARAM',TobAfPlanningParam,-1);
      TobDet2.putvalue('APP_TYPEAFPARAM','STA');
      TobDEt2.putvalue('APP_CODEPARAM', TobDet.getvalue('APE_ETATLIGNE'));
      TobDEt2.putvalue('APP_LIBELLEPARAM', TobDet.getvalue('APE_LIBELLEETAT'));
      TobDEt2.putvalue('APP_PARAMS', TobDet.getvalue('APE_PARAMS'));
    end;
    TobAfPlanningParam.InsertOrUpdateDB(True);
    TobAfPlanningEtat.free;
    TobAfPlanningParam.free;
  end;
  TobAfPlanningEtat :=Nil;
  TobAfPlanningParam :=Nil;
  GiGaTacheApres;
  ExecuteSql ('DELETE  from choixcod where cc_type="atu" and cc_code in ("094","009","013","014","015")');
  InsertChoixCode('ATU', '904', 'scotraire', '', 'temporaire');
  InsertChoixCode('ATU', '009', 'code ressource', '', 'code assistant');
  InsertChoixCode('ATU', '013', 'cette ressource', '', 'cet assistant');
  InsertChoixCode('ATU', '014', 'ressources', '', 'assistants');
  InsertChoixCode('ATU', '015', 'ressource', '', 'assistant');
  //fin GIGA
End;



procedure GPMajTypeTarif;
var TSoc, TRef: THTable;
  REFfonctionnalite: string;
begin
  TSoc := THTable.Create(Application);
  TSoc.DatabaseName := DBSoc.DatabaseName;
  TSoc.Tablename := 'YTARIFSTYPE';
  TSoc.Indexname := 'YTV_CLE1';
  TSoc.Open;
  TRef := OpenTableREF('YTARIFSTYPE', 'YTV_CLE1');
  while not TRef.EOF do
  begin
    REFfonctionnalite := TRef.FindField('YTV_FONCTIONNALITE').AsString;
    If not TSoc.Findkey([REFfonctionnalite]) then
      AddParamGc(TSoc,TRef);
    TRef.Next;
  end;
  Ferme(TRef);
  TSoc.Close;
  TSoc.Free;
end;

procedure GPMajSTKNature;
var TSoc, TRef: THTable;
  REFQualifmvt: string;
begin
  TSoc := THTable.Create(Application);
  TSoc.DatabaseName := DBSoc.DatabaseName;
  TSoc.Tablename := 'STKNATURE';
  TSoc.Indexname := 'GSN_CLE1';
  TSoc.Open;
  TRef := OpenTableREF('STKNATURE', 'GSN_CLE1');
  while not TRef.EOF do
  begin
    REFQualifmvt := TRef.FindField('GSN_QUALIFMVT').AsString;
    If not TSoc.Findkey([REFQualifmvt]) then
      AddParamGc(TSoc,TRef);
    TRef.Next;
  end;
  Ferme(TRef);
  TSoc.Close;
  TSoc.Free;
end;

procedure GPMajSTKValoParam;
var
  TSoc, TRef: THTable;
begin
  If  not ExisteSql ('select 1 from STKVALOPARAM ') then
  begin
    TSoc := THTable.Create(Application);
    TSoc.DatabaseName := DBSoc.DatabaseName;
    TSoc.Tablename := 'STKVALOPARAM';
    TSoc.Indexname := 'GVP_CLE1';
    TSoc.Open;
    TRef := OpenTableREF('STKVALOPARAM', 'GVP_CLE1');
    while not TRef.EOF do
    begin
      AddParamGc(TSoc,TRef);
      TRef.Next;
    end;
    Ferme(TRef);
    TSoc.Close;
    TSoc.Free;
  end;

  If  not ExisteSql ('select 1 from STKVALOTYPE ') then
  begin
    TSoc := THTable.Create(Application);
    TSoc.DatabaseName := DBSoc.DatabaseName;
    TSoc.Tablename := 'STKVALOTYPE';
    TSoc.Indexname := 'GVT_CLE1';
    TSoc.Open;
    TRef := OpenTableREF('STKVALOTYPE', 'GVT_CLE1');
    while not TRef.EOF do
    begin
      AddParamGc(TSoc,TRef);
      TRef.Next;
    end;
    Ferme(TRef);
    TSoc.Close;
    TSoc.Free;
  end;
end;



Procedure MajImo646;
Var St : String ;
BEGIN
St:='UPDATE IAIDSAI SET IAI_BASE05=0, IAI_BASE05_MIN=0, IAI_BASE05_MIN_O="-", IAI_BASE05_MAX=0, IAI_BASE05_MAX_O="-", ' +
    'IAI_METH05="", IAI_METH05_O="-", IAI_SPEC05="", IAI_TAUX05=0, IAI_TAUX05_O="-", IAI_TAUX05_MIN=0, IAI_TAUX05_MIN_O="-", '+
    'IAI_TAUX05_MAX=0, IAI_TAUX05_MAX_O="-", IAI_DUREE05=0, IAI_DUREE05_O="-", IAI_DUREE05_MIN=0, IAI_DUREE05_MIN_O="-", '+
    'IAI_DUREE05_MAX=0, IAI_DUREE05_MAX_O="-", IAI_COEF05=0, '+
    'IAI_BASE06=0, IAI_BASE06_MIN=0, IAI_BASE06_MIN_O="-", IAI_BASE06_MAX=0, IAI_BASE06_MAX_O="-", ' +
    'IAI_METH06="", IAI_METH06_O="-", IAI_SPEC06="", IAI_TAUX06=0, IAI_TAUX06_O="-", IAI_TAUX06_MIN=0, IAI_TAUX06_MIN_O="-", '+
    'IAI_TAUX06_MAX=0, IAI_TAUX06_MAX_O="-", IAI_DUREE06=0, IAI_DUREE06_O="-", IAI_DUREE06_MIN=0, IAI_DUREE06_MIN_O="-", '+
    'IAI_DUREE06_MAX=0, IAI_DUREE06_MAX_O="-", IAI_COEF06=0 ' ;
ExecuteSQL(St) ;
St:='UPDATE IECRLGN SET ILG_SCENARIO="" WHERE ILG_SCENARIO IS NULL' ;
ExecuteSQL(St) ;
St:='UPDATE IMOMET SET IMT_VALRES=0' ;
ExecuteSQL(St) ;
St:='UPDATE IMOREF SET IRF_BASE05=0, IRF_TAUX05=0, IRF_DUREE05=0, IRF_METH05="", '+
    'IRF_BASE06=0, IRF_TAUX06=0, IRF_DUREE06=0, IRF_METH06="", '+
    'IRF_CBEDITE="-", IRF_OKINVENTAIRE="-", IRF_AIDESAISIEREF="", IRF_METHODEIAS="", IRF_PRIXCESSION=0, '+
    'IRF_UTILITE=0, IRF_NOREEVALUATION="-", IRF_UGT="" ' ;
ExecuteSQL(St) ;
St:='UPDATE IPARAMECR SET IPE_01_44551="", IPE_01_44566="", IPE_01_671="", IPE_01_771="" ' ;
ExecuteSQL(St) ;
St:='UPDATE IAIDSAI SET IAI_QUALIF="CPT" WHERE IAI_QUALIF IS NULL' ;
ExecuteSQL(St) ;
END ;

procedure RT_InsertLibelleTablettesFour;
var i: integer;
begin
  { infos compl four }
  for i := 0 to 4 do InsertChoixCode('R3Z', 'BL' + intTostr(i), 'Décision libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R3Z', 'DL' + intTostr(i), 'Date libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R3Z', 'CL' + intTostr(i), 'Table libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R3Z', 'TL' + intTostr(i), 'Texte libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R3Z', 'VL' + intTostr(i), 'Valeur libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R3Z', 'ML' + intTostr(i), 'Multi-choix libre ' + intTostr(i + 1), '', '');
  InsertChoixCode('R3Z', 'BLO', 'Bloc-notes ', '', '');
  { tables libres chainages }
  for i := 1 to 3 do InsertChoixCode('RFH', 'RF' + intTostr(i), 'Table libre chaînage ' + intTostr(i), '', '');
  { tables libres actions }
  for i := 1 to 3 do InsertChoixCode('RLZ', 'FA' + intTostr(i), 'Table libre action ' + intTostr(i), '', '');
End;

// JTR - Mise à jour des nouveaux évènements PCP pour SocRef 654
procedure MajEvtPCP_V654;
Var TSoc,TRef : THTable ;
    EvtCode, EvtTypeTrt : String ;
BEGIN
  TSoc := THTable.Create(Application) ;
  TSoc.DatabaseName := DBSoc.DatabaseName ;
  TSoc.Tablename:='STOXEVENTS' ;
  TSoc.Indexname:='SEV_CLE1' ;
  TSoc.Open ;
  TRef := OpenTableRef('STOXEVENTS','SEV_CLE1') ;
  While Not TRef.EOF do
  begin
    if (TRef.FindField('SEV_CODEEVENT').AsString = 'PCPU_DEMANDE')
       or (TRef.FindField('SEV_CODEEVENT').AsString = 'PCPU_GENERIQUE')
       or (TRef.FindField('SEV_CODEEVENT').AsString = 'PCPU_INIT_ACH')
       or (TRef.FindField('SEV_CODEEVENT').AsString = 'PCPU_INIT_VTE')
       or (TRef.FindField('SEV_CODEEVENT').AsString = 'PCPU_MOD_ACH')
       or (TRef.FindField('SEV_CODEEVENT').AsString = 'PCPU_MOD_VTE')
       or (TRef.FindField('SEV_CODEEVENT').AsString = 'PCPU_REMONTEE_ACH')
       or (TRef.FindField('SEV_CODEEVENT').AsString = 'PCPU_REMONTEE_VTE')then
    begin
      EvtCode := TRef.FindField('SEV_CODEEVENT').AsString ;
      EvtTypeTrt := TRef.FindField('SEV_TYPETRT').AsString ;
      if Not TSoc.Findkey([EvtTypeTrt,EvtCode]) then
      begin
        AddParamGC(TSoc,TRef);
      end;
    end;
    TRef.Next ;
  end;
  Ferme(TRef) ;
  TSoc.Close ;
  TSoc.Free ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 21/02/2002
Description .. : Renvoie un where pour le code générique
*****************************************************************}
function whereGA(Article: string): string;
begin
	Result := 'GA_ARTICLE = "' + Article + '"';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 10/06/2002
Modifié le ... :  01/12/2004
Description .. : Recherche d'un article
Suite ........ : sur le code article sur 18 caractères
Mots clefs ... :
*****************************************************************}
function wExistGA(Article: string; WithAlert: Boolean = false; Where : string = ''): Boolean;
var
	Sql  : string;
begin
	Sql := 'SELECT 1'
			 + ' FROM ARTICLE'
			 + ' WHERE ' + WhereGA(Article)
       + iif(Where <> '', ' AND ' + Where, '')
       ;

	Result := ExisteSQL(Sql);

 // Pas dans PGIMajVer
 // if WithAlert and (not Result) then
 // begin
 //   PgiError(TraduireMemoire('L''article ') + wGetCodeArticleFromArticle(Article) + TraduireMemoire(' n''existe pas.'), 'Article');
 // end;
end;

procedure GIGAAvant658;
  var
  TOB292, TOB293, TOB143, TOB152, TOB151, TOB144, TOB294, TOBDet : TOB;
  QQ : TQuery;
  ssql: string;
  ii, indice : integer;
begin
  // Si a déjà passé la maj de socref, on ne relance pas la procédure de transferts de confidentialité
  if ExisteSQL( 'SELECT MN_1 FROM MENU WHERE MN_1=294') then
    exit;
 
  // On récupère les TOB nécessaire au transfert de confidentialité des menus qui changent de module
  // Les autres menus transfèrent leur confidentialité automatiquement par la maj de socref s'ils restent dans le même module
  // 292
  TOB292 := TOB.Create ('tobmenu_292', Nil, -1);
  // très peu de champs, très peu de lignes, on peut faire un SELECT *
  ssql := 'SELECT * FROM MENU WHERE MN_1=141 AND MN_TAG IN'
          + ' (71101,71503,71501,-71210,71211,71212,71214,71213,71204,71430,71403,71404,71405,71406)';
  try
    QQ := OpenSQL (ssql, True);
    if not QQ.Eof then
      TOB292.LoadDetailDB ('MENU', '', '', QQ, False);
  finally
    ferme (QQ);
  end;
 
  // 293
  TOB293 := TOB.Create ('tobmenu_293', Nil, -1);
  // très peu de champs, très peu de lignes, on peut faire un SELECT *
  ssql := 'SELECT * FROM MENU WHERE MN_1=141 AND MN_TAG IN'
          + ' (71104,71502,-71230,71231,71232,71233,71234,71241,71242,71440,72502,71441,71442,71445)';
  try
    QQ := OpenSQL (ssql, True);
    if not QQ.Eof then
      TOB293.LoadDetailDB ('MENU', '', '', QQ, False);
  finally
    ferme (QQ);
  end;
 
  // 143
  TOB143 := TOB.Create ('tobmenu_143', Nil, -1);
  // très peu de champs, très peu de lignes, on peut faire un SELECT *
  ssql := 'SELECT * FROM MENU WHERE (MN_1=142 AND MN_TAG IN (72504,72505))'
          + ' OR (MN_1=151 AND MN_TAG IN (139100,73520,73521,73510,-73530,73531,73532,-73540,73541,73542))';
  try
    QQ := OpenSQL (ssql, True);
    if not QQ.Eof then
      TOB143.LoadDetailDB ('MENU', '', '', QQ, False);
  finally
    ferme (QQ);
  end;
 
  // 152
  TOB152 := TOB.Create ('tobmenu_152', Nil, -1);
  // très peu de champs, très peu de lignes, on peut faire un SELECT *
  ssql := 'SELECT * FROM MENU WHERE MN_1=151 AND MN_TAG IN'
          + ' (139200,-139210,139211,139212,-33220,139221,139222)';
  try
    QQ := OpenSQL (ssql, True);
    if not QQ.Eof then
      TOB152.LoadDetailDB ('MENU', '', '', QQ, False);
  finally
    ferme (QQ);
  end;
 
  // 151
  TOB151 := TOB.Create ('tobmenu_151', Nil, -1);
  // très peu de champs, très peu de lignes, on peut faire un SELECT *
  ssql := 'SELECT * FROM MENU WHERE (MN_1=141 AND MN_TAG=71250)'
          + ' OR (MN_1=142 AND MN_TAG IN (-72700,72702,72705,72704,72703,72706,72701,72708))';
  try
    QQ := OpenSQL (ssql, True);
    if not QQ.Eof then
      TOB151.LoadDetailDB ('MENU', '', '', QQ, False);
  finally
    ferme (QQ);
  end;
 
  // 144
  TOB144 := TOB.Create ('tobmenu_144', Nil, -1);
  // très peu de champs, très peu de lignes, on peut faire un SELECT *
  ssql := 'SELECT * FROM MENU WHERE (MN_1=141 AND MN_TAG IN'
          + ' (-71510,71511,71512,71513,71515,-71270,71271,71272,71420,71410,71407,72509,71414,71607,-71600,'
          + '71601,71604,71602,71603,71605,71606,71615,71616,-71610,71611,71612,71613,-71618,71617,71618))'
          + ' OR (MN_1=152 AND MN_TAG IN (-138400,138401,138402,138403))';
  try
    QQ := OpenSQL (ssql, True);
    if not QQ.Eof then
      TOB144.LoadDetailDB ('MENU', '', '', QQ, False);
  finally
    ferme (QQ);
  end;
 
  // 294
  TOB294 := TOB.Create ('tobmenu_294', Nil, -1);
  // très peu de champs, très peu de lignes, on peut faire un SELECT *
  ssql := 'SELECT * FROM MENU WHERE MN_1=144 AND MN_TAG IN'
          + ' (74800,74801,74802,74650,74654,74653,74656,74651,74600,-74615,74752,74755,74612,74608,74617,'
          + '-74616,74605,74621,74604,74606,74607,74603,-74623,74624,74623,74602,74614,74613,74609,74601,'
          + '74700,74702,74701,-74706,74707,74706,74710,74708,74709,-74712,74712,74714,74713,74711,74703,'
          + '74704,74705,0,31,32,-74610,74615,74616,74622,74619,74618)';
  try
    QQ := OpenSQL (ssql, True);
    if not QQ.Eof then
      TOB294.LoadDetailDB ('MENU', '', '', QQ, False);
  finally
    ferme (QQ);
  end;
 
  // Boucles de modification des clés des lignes de menu
  indice := 1000;
  for ii := 0 to TOB292.detail.count-1 do
    begin
      TOBDet := TOB292.detail[ii];
      TOBDet.PutValue('MN_1', 292);
      TOBDet.PutValue('MN_2', indice);
      Inc(indice);
    end;
  for ii := 0 to TOB293.detail.count-1 do
    begin
      TOBDet := TOB293.detail[ii];
      TOBDet.PutValue('MN_1', 293);
      TOBDet.PutValue('MN_2', indice);
      Inc(indice);
    end;
  for ii := 0 to TOB143.detail.count-1 do
    begin
      TOBDet := TOB143.detail[ii];
      TOBDet.PutValue('MN_1', 143);
      TOBDet.PutValue('MN_2', indice);
      Inc(indice);
    end;
  for ii := 0 to TOB152.detail.count-1 do
    begin
      TOBDet := TOB152.detail[ii];
      TOBDet.PutValue('MN_1', 152);
      TOBDet.PutValue('MN_2', indice);
      Inc(indice);
    end;
  for ii := 0 to TOB151.detail.count-1 do
    begin
      TOBDet := TOB151.detail[ii];
      TOBDet.PutValue('MN_1', 151);
      TOBDet.PutValue('MN_2', indice);
      Inc(indice);
    end;
  for ii := 0 to TOB144.detail.count-1 do
    begin
      TOBDet := TOB144.detail[ii];
      TOBDet.PutValue('MN_1', 144);
      TOBDet.PutValue('MN_2', indice);
      Inc(indice);
    end;
  for ii := 0 to TOB294.detail.count-1 do
    begin
      TOBDet := TOB294.detail[ii];
      TOBDet.PutValue('MN_1', 294);
      TOBDet.PutValue('MN_2', indice);
      Inc(indice);
    end;
 

  // Insertion en base des lignes de sauvegarde des confidentialités
  try
    if (TOB292.detail.count <> 0) then
      TOB292.InsertDB (nil);
    if (TOB293.detail.count <> 0) then
      TOB293.InsertDB (nil);
    if (TOB143.detail.count <> 0) then
      TOB143.InsertDB (nil);
    if (TOB152.detail.count <> 0) then
      TOB152.InsertDB (nil);
    if (TOB151.detail.count <> 0) then
      TOB151.InsertDB (nil);
    if (TOB144.detail.count <> 0) then
      TOB144.InsertDB (nil);
    if (TOB294.detail.count <> 0) then
      TOB294.InsertDB (nil);
 
  finally
    // Libération des TOB
    TOB292.Free;
    TOB293.Free;
    TOB143.Free;
    TOB152.Free;
    TOB151.Free;
    TOB144.Free;
    TOB294.Free;
  end;
end;

procedure W_SAV_INTER_COMPL;
var i: integer;
begin
  for i := 0 to 4 do InsertChoixCode('R7Z', 'BL' + intTostr(i), 'Décision libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R7Z', 'DL' + intTostr(i), 'Date libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R7Z', 'CL' + intTostr(i), 'Table libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R7Z', 'TL' + intTostr(i), 'Texte libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R7Z', 'VL' + intTostr(i), 'Valeur libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R7Z', 'ML' + intTostr(i), 'Multi-choix libre ' + intTostr(i + 1), '', '');
end;




procedure UpdateChoixCodeLIBABR(sType, sCode, sLibelle, sAbrege : String);
begin
  if ExisteSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="'+sType+'" AND CC_CODE="'+sCode+'"')then
    ExecuteSQL('UPDATE CHOIXCOD SET CC_LIBELLE="'+ sLibelle + '", CC_ABREGE="'+ sAbrege + '" WHERE CC_CODE="'+sCode+'" AND CC_TYPE="'+sType+'"');
end;

procedure W_SAV_COMPETENCE_UPDATE;
var i : Integer ;
begin
  for I := 1 to 3 do UpdateChoixCodeLIBABR ('ZLI', 'CP' + intTostr(i), 'Table libre ' + intTostr(i), 'Table libre ' + intTostr(i));
  for i := 1 to 5 do UpdateChoixCodeLIBABR ('ZLI', 'RH' + intTostr(i), 'Table libre ' + intTostr(i), 'Table libre ' + intTostr(i));
End ;

procedure GPMajPdrType;
var
  TSoc, TRef: THTable;
begin
  If  not ExisteSql ('select 1 from WPDRTYPE ') then
  begin
    TSoc := THTable.Create(Application);
    TSoc.DatabaseName := DBSoc.DatabaseName;
    TSoc.Tablename := 'WPDRTYPE';
    TSoc.Indexname := 'WRT_CLE1';
    TSoc.Open;
    TRef := OpenTableREF('WPDRTYPE', 'WRT_CLE1');
    while not TRef.EOF do
    begin
      AddParamGc(TSoc,TRef);
      TRef.Next;
    end;
    Ferme(TRef);
    TSoc.Close;
    TSoc.Free;
  end;
end;

// Mise à jour de la table xWizards //
procedure GPMajxWizards;
var TSoc, TRef: THTable;
  REFQualif: string;
begin
  TSoc := THTable.Create(Application);
  TSoc.DatabaseName := DBSoc.DatabaseName;
  TSoc.Tablename := 'XWIZARDS';
  TSoc.Indexname := 'XWI_CLE1';
  TSoc.Open;
  TRef := OpenTableREF('XWIZARDS', 'XWI_CLE1');
  while not TRef.EOF do
  begin
    REFQualif := TRef.FindField('XWI_CODEXWI').AsString;
    If not TSoc.Findkey([REFQualif]) then
      AddParamGc(TSoc,TRef);
    TRef.Next;
  end;
  Ferme(TRef);
  TSoc.Close;
  TSoc.Free;
end;

procedure RT_InsertLibelleInfoComplArticle;
var i: integer;
begin
  for i := 0 to 4 do InsertChoixCode('R4Z', 'BL' + intTostr(i), 'Décision libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R4Z', 'DL' + intTostr(i), 'Date libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R4Z', 'CL' + intTostr(i), 'Table libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R4Z', 'TL' + intTostr(i), 'Texte libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R4Z', 'VL' + intTostr(i), 'Valeur libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R4Z', 'ML' + intTostr(i), 'Multi-choix libre ' + intTostr(i + 1), '', '');
  InsertChoixCode('R4Z', 'BLO', 'Bloc-notes ', '', '');
end;

procedure RT_InsertLibelleInfoComplLigne;
var i: integer;
begin
  for i := 0 to 4 do InsertChoixCode('R8Z', 'BL' + intTostr(i), 'Décision libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R8Z', 'DL' + intTostr(i), 'Date libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R8Z', 'CL' + intTostr(i), 'Table libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R8Z', 'TL' + intTostr(i), 'Texte libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R8Z', 'VL' + intTostr(i), 'Valeur libre ' + intTostr(i + 1), '', '');
  for i := 0 to 4 do InsertChoixCode('R8Z', 'ML' + intTostr(i), 'Multi-choix libre ' + intTostr(i + 1), '', '');
  InsertChoixCode('R8Z', 'BLO', 'Bloc-notes ', '', '');
end ;

procedure CFE5_RegulLiacodedo;
begin
	ExecuteSQL('update l01_valdonnee set l01_liacodedoc= "CFEM0A"  where l01_liaindice in (select l00_liaindice from l00_refdonnee where l00_liacodedoc = "CFEM0A" )');
	ExecuteSQL('update l01_valdonnee set l01_liacodedoc= "CFEM0B"  where l01_liaindice in (select l00_liaindice from l00_refdonnee where l00_liacodedoc = "CFEM0B" )');
	ExecuteSQL('update l01_valdonnee set l01_liacodedoc= "CFEM0P"  where l01_liaindice in (select l00_liaindice from l00_refdonnee where l00_liacodedoc = "CFEM0P" )');
	ExecuteSQL('update l01_valdonnee set l01_liacodedoc= "CFEM0P2" where l01_liaindice in (select l00_liaindice from l00_refdonnee where l00_liacodedoc = "CFEM0P2")');

	ExecuteSQL('update l01_valdonnee set l01_liacodedoc= "CFEM2A"  where l01_liaindice in (select l00_liaindice from l00_refdonnee where l00_liacodedoc = "CFEM2A" )');
	ExecuteSQL('update l01_valdonnee set l01_liacodedoc= "CFEM2B"  where l01_liaindice in (select l00_liaindice from l00_refdonnee where l00_liacodedoc = "CFEM2B" )');
	ExecuteSQL('update l01_valdonnee set l01_liacodedoc= "CFEM2P"  where l01_liaindice in (select l00_liaindice from l00_refdonnee where l00_liacodedoc = "CFEM2P" )');

	ExecuteSQL('update l01_valdonnee set l01_liacodedoc= "CFEM3A"  where l01_liaindice in (select l00_liaindice from l00_refdonnee where l00_liacodedoc = "CFEM3A" )');
	ExecuteSQL('update l01_valdonnee set l01_liacodedoc= "CFEM3A2" where l01_liaindice in (select l00_liaindice from l00_refdonnee where l00_liacodedoc = "CFEM3A2")');

	ExecuteSQL('update l01_valdonnee set l01_liacodedoc= "CFEM3B"  where l01_liaindice in (select l00_liaindice from l00_refdonnee where l00_liacodedoc = "CFEM3B" )');
	ExecuteSQL('update l01_valdonnee set l01_liacodedoc= "CFEM3B2" where l01_liaindice in (select l00_liaindice from l00_refdonnee where l00_liacodedoc = "CFEM3B2")');

	ExecuteSQL('update l01_valdonnee set l01_liacodedoc= "CFEM4"   where l01_liaindice in (select l00_liaindice from l00_refdonnee where l00_liacodedoc = "CFEM4"  )');
	ExecuteSQL('update l01_valdonnee set l01_liacodedoc= "CFEM4P"  where l01_liaindice in (select l00_liaindice from l00_refdonnee where l00_liacodedoc = "CFEM4P" )');

end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 04/02/2005
Modifié le ... :   /  /    
Description .. : Mise à jour des définitions de tables libres de la table 
Suite ........ : COMMUN vers CHOIXCOD 
Mots clefs ... : 
*****************************************************************}
procedure CPMajTLVersCHOIXCOD;
var TTl : TOB;
    i : integer;
    stCode, stLibelle, stAbrege,stLibre : string;
    bReinit : boolean;
begin
{$IFDEF AGL580}
 // bNull :=  V_PGI.SuppressionDesNullsDeLaTob;
 // V_PGI.SuppressionDesNullsDeLaTob := True;
{$ENDIF AGL580}
  { Si il existe des enregistrements tables libres dans CHOIXCOD, on ne fait pas le traitement }
  if ExisteSQL ('SELECT * FROM CHOIXCOD WHERE CC_TYPE="NAT"') then exit;
  TTl := TOB.Create ('', nil, -1);
  if (ctxPCL in V_PGI.PGIContexte) then v_pgi.enableDEShare := True;
  try
    { Chargement des enregistrements de COMMUN }
    TTl.LoadDetailFromSQL ('SELECT * FROM COMMUN WHERE CO_TYPE="NAT"');
    for i :=0 to TTl.Detail.Count - 1 do
    begin
      { Mise à jour de l'enregistrement CHOIXCOD }
      stCode := TTl.Detail[i].GetValue('CO_CODE');
      if (ctxPCL in V_PGI.PGIContexte) then // Dans le cas PCL, on regarde si la table libre était gérée dans le dossier
        bReinit := not ExisteSQL ('SELECT NT_TYPECPTE FROM NATCPTE WHERE NT_TYPECPTE="'+stCode+'"')
      else bReinit := False;
      if bReinit then
      begin
        stLibelle := TraduireMemoire('Table n°')+IntToStr(StrToInt(Copy(stCode,2,2))+1);
        stAbrege := '-';
      end else
      begin
        stLibelle := TTl.Detail[i].GetValue('CO_LIBELLE');
        stAbrege := TTl.Detail[i].GetValue('CO_ABREGE');
      end;
      stLibre := TTl.Detail[i].GetValue('CO_LIBRE');
      { Insertion dans la base }
      ExecuteSQL ('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) '+
        'VALUES ("NAT","'+stCode+'","'+stLibelle+'","'+stAbrege+'","'+stLibre+'")');
    end;
  finally
    TTl.Free;
    if (ctxPCL in V_PGI.PGIContexte) then v_pgi.enableDEShare := False;
{$IFDEF AGL580}
   // V_PGI.SuppressionDesNullsDeLaTob := bNull;
{$ENDIF AGL580}
  end;
end;

procedure UpdateNatureSuivante;
var iPos : integer;
NatureP,NatureSuiv : string;
QPP : TQuery;
begin
  QPP := OpenSQL('SELECT GPP_NATUREPIECEG,GPP_NATURESUIVANTE FROM PARPIECE',true);
  while not QPP.EOF do
  begin
    NatureP := QPP.FindField('GPP_NATUREPIECEG').AsString;
    NatureSuiv := QPP.FindField('GPP_NATURESUIVANTE').AsString;
    iPos := pos(NatureP+';',NatureSuiv);
    if iPos > 0 then
    begin
      system.Delete(NatureSuiv,iPos,Length(NatureP)+1);
      ExecuteSQL('UPDATE PARPIECE SET GPP_NATURESUIVANTE="' + NatureSuiv + '" WHERE GPP_NATUREPIECEG="' + NatureP + '"');
    end;
    QPP.Next;
  end;
  Ferme(QPP);
end;


procedure GPCopyWPRINTTV;
var TSoc, TRef: THTable;
  REFCode: string;
begin
  TSoc := THTable.Create(Application);
  TSoc.DatabaseName := DBSoc.DatabaseName;
  TSoc.Tablename := 'WPRINTTV';
  TSoc.Indexname := 'WTV_CLE1';
  TSoc.Open;
  TRef := OpenTableREF('WPRINTTV', 'WTV_CLE1');
  while not TRef.EOF do
  begin
    REFCode := TRef.FindField('WTV_CODETV').AsString;
    If not TSoc.Findkey([REFCode]) then
      AddParamGc(TSoc,TRef);
    TRef.Next;
  end;
  Ferme(TRef);
  TSoc.Close;
  TSoc.Free;
end;

procedure RT_InsertInfoCompl;
var i : integer;
    StConc,Code : String;
begin
  StConc:='9;B;D;H;J;K;N;O;Q;V;W;X;Y;Z';;
  Repeat
  Code:=ReadTokenSt(StConc) ;
  if Code <> '' then
  begin
    for i := 0 to 4 do InsertChoixCode('R'+Code+'Z', 'BL' + intTostr(i), 'Décision libre ' + intTostr(i + 1), '', '');
    for i := 0 to 4 do InsertChoixCode('R'+Code+'Z', 'DL' + intTostr(i), 'Date libre ' + intTostr(i + 1), '', '');
    for i := 0 to 4 do InsertChoixCode('R'+Code+'Z', 'CL' + intTostr(i), 'Table libre ' + intTostr(i + 1), '', '');
    for i := 0 to 4 do InsertChoixCode('R'+Code+'Z', 'TL' + intTostr(i), 'Texte libre ' + intTostr(i + 1), '', '');
    for i := 0 to 4 do InsertChoixCode('R'+Code+'Z', 'VL' + intTostr(i), 'Valeur libre ' + intTostr(i + 1), '', '');
    for i := 0 to 4 do InsertChoixCode('R'+Code+'Z', 'ML' + intTostr(i), 'Multi-choix libre ' + intTostr(i + 1), '', '');
    InsertChoixCode('R'+Code+'Z', 'BLO', 'Bloc-notes ', '', '');
  end;
  until  StConc='' ;
end;



procedure W_SAV_ArticlesParc;
var i: integer;
begin
  for i := 1 to 3 do InsertChoixCode('ZLP', 'WB' + intTostr(i), 'Décision libre ' + intTostr(i), '', '');
  for i := 1 to 3 do InsertChoixCode('ZLP', 'WC' + intTostr(i), 'Texte libre ' + intTostr(i), '', '');
  for i := 1 to 3 do InsertChoixCode('ZLP', 'WD' + intTostr(i), 'Date libre ' + intTostr(i), '', '');
  for i := 1 to 3 do InsertChoixCode('ZLP', 'WM' + intTostr(i), 'Montant libre ' + intTostr(i), '', '');
  for i := 1 to 9 do InsertChoixCode('ZLP', 'WT' + intTostr(i), 'Table libre ' + intTostr(i), '', '');
  InsertChoixCode('ZLP', 'WTA', 'Table libre 10', '', '');
end;


procedure RT_InsertLibelleInfoComplDixArtLigne;
var i: integer;
begin
  for i := 5 to 9 do InsertChoixCode('R4Z', 'CL' + intTostr(i), 'Table libre ' + intTostr(i + 1), '', '');
  for i := 5 to 9 do InsertChoixCode('R4Z', 'ML' + intTostr(i), 'Multi-choix libre ' + intTostr(i + 1), '', '');
  for i := 5 to 9 do InsertChoixCode('R8Z', 'CL' + intTostr(i), 'Table libre ' + intTostr(i + 1), '', '');
  for i := 5 to 9 do InsertChoixCode('R8Z', 'ML' + intTostr(i), 'Multi-choix libre ' + intTostr(i + 1), '', '');
end;


procedure YY_InitChampsSystemes;
var SS : string;
begin
    SS := UsDateTime(IDate1900);

    // scriptsp
    ExecuteSQL('update scriptsp set ysc_datecreation="' + SS + '" where ysc_datecreation is null');
    ExecuteSQL('update scriptsp set ysc_datemodif="' + SS + '" where ysc_datemodif is null');
    ExecuteSQL('update scriptsp set ysc_createur="" where ysc_createur is null');
    ExecuteSQL('update scriptsp set ysc_utilisateur="" where ysc_utilisateur is null');
    ExecuteSQL('update scriptsp set ysc_numversion=0 where ysc_numversion is null');
    // formesp
    ExecuteSQL('update formesp set yfm_datecreation="' + SS + '" where yfm_datecreation is null');
    ExecuteSQL('update formesp set yfm_datemodif="' + SS + '" where yfm_datemodif is null');
    ExecuteSQL('update formesp set yfm_createur="" where yfm_createur is null');
    ExecuteSQL('update formesp set yfm_utilisateur="" where yfm_utilisateur is null');
    ExecuteSQL('update formesp set yfm_numversion=0 where yfm_numversion is null');
    // modeles
    ExecuteSQL('update modeles set mo_datecreation="' + SS + '" where mo_datecreation is null');
    ExecuteSQL('update modeles set mo_datemodif="' + SS + '" where mo_datemodif is null');
    ExecuteSQL('update modeles set mo_createur="" where mo_createur is null');
    ExecuteSQL('update modeles set mo_utilisateur="" where mo_utilisateur is null');
    ExecuteSQL('update modeles set mo_numversion=0 where mo_numversion is null');
    // liste
    ExecuteSQL('update liste set li_datecreation="' + SS + '" where li_datecreation is null');
    ExecuteSQL('update liste set li_datemodif="' + SS + '" where li_datemodif is null');
    ExecuteSQL('update liste set li_createur="" where li_createur is null');
    ExecuteSQL('update liste set li_utilisateur="" where li_utilisateur is null');
    ExecuteSQL('update liste set li_numversion=0 where li_numversion is null');
    // filtres
    ExecuteSQL('update filtres set fi_datecreation="' + SS + '" where fi_datecreation is null');
    ExecuteSQL('update filtres set fi_datemodif="' + SS + '" where fi_datemodif is null');
    ExecuteSQL('update filtres set fi_createur="" where fi_createur is null');
    ExecuteSQL('update filtres set fi_utilisateur="" where fi_utilisateur is null');
    // graphs
    ExecuteSQL('update graphs set gr_datecreation="' + SS + '" where gr_datecreation is null');
    ExecuteSQL('update graphs set gr_datemodif="' + SS + '" where gr_datemodif is null');
    ExecuteSQL('update graphs set gr_createur="" where gr_createur is null');
    ExecuteSQL('update graphs set gr_utilisateur="" where gr_utilisateur is null');
end;

{Passage de la V6.0 à V7.0}
procedure MaJdesPrixGSMFromBlocNote;
var
  TobGSM  : tob;
  iCpt    : integer;
  sSQL, sBlocNote : string;
begin {MaJdesPrixGSMFromBlocNote}
  {$IFNDEF PGIMAJVER}
  if (PGIAsk('Confirmez-vous la mise à jour des prix des mouvements à partir du bloc note ? ', '') = mrYes) then
  {$ENDIF PGIMAJVER}
  begin
    TobGSM := Tob.Create('_STKMOUVEMENT_', nil, -1);
    try

      //SAUZET Jean-Luc Demande N° 1238 (Ajout du champ GSM_PRIXSAISIS dans la requête)
      {GPAO_BUG700_MM_GP13979}
      //sSQL := 'SELECT GSM_STKTYPEMVT, GSM_QUALIFMVT, GSM_GUID, GSM_DPA, GSM_DPR, GSM_PMAP, GSM_PMRP,'
      //      + iif(isOracle, 'GSM_BLOCNOTE', 'SUBSTRING(GSM_BLOCNOTE,0,200) AS GSM_BLOCNOTE')
      //      + ' FROM STKMOUVEMENT'
      //      + ' WHERE GSM_QUALIFMVT IN ("EIN", "CLO") AND (GSM_BLOCNOTE IS NOT NULL )'
      //      ;
      sSQL := 'SELECT GSM_STKTYPEMVT, GSM_QUALIFMVT, GSM_GUID, GSM_DPA, GSM_DPR, GSM_PMAP, GSM_PMRP, GSM_PRIXSAISIS,'
            + iif(isOracle, 'GSM_BLOCNOTE', 'SUBSTRING(GSM_BLOCNOTE,0,200) AS GSM_BLOCNOTE')
            + ' FROM STKMOUVEMENT'
            + ' WHERE GSM_QUALIFMVT IN ("EIN", "CLO") AND (GSM_BLOCNOTE IS NOT NULL )'
            ;


      TobGSM.LoadDetailDBFromSQL('STKMOUVEMENT', sSQL);
      if Assigned(TobGSM) then
      begin
        for iCpt := 0 to TobGSM.Detail.Count-1 do
        begin
          sBlocNote := TobGSM.Detail[iCpt].GetString('GSM_BLOCNOTE');
          if (copy(sBlocNote,1,3)='DPA') then
          begin
            TobGSM.Detail[iCpt].SetDouble('GSM_DPA' , Valeur(GetArgumentValue(sBlocNote, 'DPA' )));
            TobGSM.Detail[iCpt].SetDouble('GSM_DPR' , Valeur(GetArgumentValue(sBlocNote, 'DPR' )));
            TobGSM.Detail[iCpt].SetDouble('GSM_PMAP', Valeur(GetArgumentValue(sBlocNote, 'PMAP')));
            TobGSM.Detail[iCpt].SetDouble('GSM_PMRP', Valeur(GetArgumentValue(sBlocNote, 'PMRP')));

            //SAUZET Jean-Luc Demande N° 1238 (Initialisation de GSM_PRIXSAISIS)
            {GPAO_BUG700_MM_GP13979}
            TobGSM.Detail[iCpt].SetBoolean('GSM_PRIXSAISIS', True);

            TobGSM.Detail[iCpt].SetString('GSM_BLOCNOTE', '');
          end;
        end;
        TobGSM.UpdateDB(True);
      end;
    finally
      TobGSM.Free;
    end;
  end;
end; {MaJdesPrixGSMFromBlocNote}

//Procédure
//Mise à jour QSIMULATION en fonction de WCBNEVOLUTION et WPARAM
procedure RecupParamCBN;
var
 sQl : string;
 TWPF, TWPA, T : Tob;
 i : integer;
begin
  if not ExisteSQL('SELECT QSM_CODESIMU FROM QSIMULATION') then
  begin
    ExecuteSQL('UPDATE WCBNEVOLUTION SET WEV_CTX=TRIM(SUBSTRING(WEV_CTX,2,5))');
    ExecuteSQL('INSERT INTO QSIMULATION  (QSM_CTX, QSM_CODESIMU, QSM_LIBSIMU, QSM_ORIGINE, QSM_ENVIRONNEMENT, QSM_DATEMAJIMPORT,'
             + 'QSM_USERMAJIMPORT, QSM_OKMAJIMPORT, QSM_DATEVALIDER, QSM_USERVALIDER, QSM_CALCUL, QSM_DATECALCUL,'
             + 'QSM_USERCALCUL, QSM_MODESIMU)'
             + ' SELECT "0",WEV_CTX,WEV_CTX,"","0","' + UsDateTime(iDate1900) + '","","-","' + UsDateTime(iDate1900) + '","","","' + UsDateTime(iDate1900) + '"'
             + ',"","2"'
             + ' FROM WCBNEVOLUTION '
             + ' GROUP BY WEV_CTX ');

    Sql := 'SELECT "CALCULBESOINSNETS" AS WPF_CODEFONCTION, PGIGUID AS WPF_GUID, QSM_CODESIMU AS WPF_CONTEXTE, "" AS WPF_BLOCNOTE '
         + ' FROM QSIMULATION'
         + ' WHERE QSM_CTX="0" AND QSM_MODESIMU="2"'
         ;
    TWPF := Tob.Create('WPF', nil, -1);
    try
      if TWPF.LoadDetailDBFromSql('WPARAMFONCTION', Sql) then
      begin
        Sql := 'SELECT WPA_CODEPARAM, WPA_UTILISATEUR, WPA_BLOCNOTE '
             + ' FROM WPARAM'
             + ' WHERE WPA_CODEPARAM="CALCULBESOINSNETS"'
             ;
        TWPA := Tob.Create('WPA', nil, -1);
        try
          if TWPA.LoadDetailDBFromSql('WPARAM', Sql) then
          begin
            for i := 0 to (TWPF.Detail.count) -1 do
            begin
              T := TWPA.FindFirst(['WPA_CODEPARAM', 'WPA_UTILISATEUR'], ['CALCULBESOINSNETS', TWPF.Detail[i].GetString('WPF_CONTEXTE')], False);
              if assigned(T) then
              begin
                TWPF.Detail[i].SetString('WPF_BLOCNOTE', T.GetString('WPA_BLOCNOTE'));
                sQl := 'SELECT 1 FROM WPARAMFONCTION '
                     + ' WHERE WPF_CODEFONCTION="' + TWPF.Detail[i].GetString('WPF_CODEFONCTION')
                     + '" AND WPF_GUID="' + TWPF.Detail[i].GetString('WPF_GUID') + '"';
                if not ExisteSQL(sQl) then
                  TWPF.Detail[i].InsertDB(nil, False);
              end;
            end;
          end;
        finally
          TWPA.Free;
        end;
      end;
    finally
      TWPF.free;
    end;
  end;
end;



procedure RT_InsertTablettesSAV;
var
  iNum,iInc: integer;
  sPrefixe,sCode,sType,sLibelle : string;
begin
      //-- Champs libres //
      for iNum:=8 to 9 do
      begin
         case iNum of
            8 : sPrefixe := 'WPN';
            9 : sPrefixe := 'WPV';
         end ;
         for iInc := 1 to 3 do
         begin
            //-- Booleens libres //
            sType    := 'W' + IntToStr(iNum) + 'Y' ;
            sCode    := 'T' + sPrefixe + '_BOOLLIBRE' + IntToStr(iInc) ;
            sLibelle := 'Décision libre ' + IntToStr(iInc);
            InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
            //-- Char libres --//
            sType    := 'W' + IntToStr(iNum) + 'Z' ;
            sCode    := 'T' + sPrefixe + '_CHARLIBRE' + IntToStr(iInc) ;
            sLibelle := 'Texte libre ' + IntToStr(iInc);
            InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
            //-- Dates libres --//
            sType    := 'W' + IntToStr(iNum) + 'X' ;
            sCode    := 'T' + sPrefixe + '_DATELIBRE' + IntToStr(iInc) ;
            sLibelle := 'Date libre ' + IntToStr(iInc);
            InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
            //-- Valeur libres --//
            sType    := 'W' + IntToStr(iNum) + 'W' ;
            sCode    := 'T' + sPrefixe + '_VALLIBRE' + IntToStr(iInc) ;
            sLibelle := 'Valeur libre ' + IntToStr(iInc);
            InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
         end;
         //-- Libellés libres --//
         for iInc := 1 to 10 do
         begin
            case iInc of
            1..9 : begin
                     sType    := 'W' + IntToStr(iNum) + 'V' ;
                     sCode    := 'T' + sPrefixe + '_LIBRE'  + sPrefixe + IntToStr(iInc) ;
                     sLibelle := 'Table libre ' + IntToStr(iInc);
                     InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
                  end;
            10  : begin
                     sType    := 'W' + IntToStr(iNum) + 'V' ;
                     sCode    := 'T' + sPrefixe + '_LIBRE' + sPrefixe +  'A' ;
                     sLibelle := 'Table libre 10';
                     InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
                  end;
            end;
         end;
      end;
end;

// Mise à jour de la GP suite à l'intégration de MES
Procedure GPAOMajMES;
begin
  ExecuteSql('update wnomelig set wnl_DateApp = "'+UsdateTime(iDate1900)+'", wnl_DatePer =  "'+UsdateTime(iDate2099)+'"');
  ExecuteSql('UPDATE WGAMMELIG SET WGL_SUIVIOPMES="OF",WGL_ANTERIORITEMES="",WGL_GROUPAGEMES=0,WGL_DateApp = "'+
    UsdateTime(iDate1900)+'", WGL_DatePer = "'+
    UsdateTime(iDate2099)+'" WHERE WGL_SUIVIOPMES IS NULL');

  ExecuteSql('UPDATE WGAMMERES'
         + ' SET WGR_ORDRE=1+(SELECT COUNT(*)'
         +                  ' FROM WGAMMERES W2'
         +                  ' WHERE W2.WGR_NATURETRAVAIL=WGAMMERES.WGR_NATURETRAVAIL'
         +                  ' AND W2.WGR_ARTICLE=WGAMMERES.WGR_ARTICLE'
         +                  ' AND W2.WGR_MAJEUR=WGAMMERES.WGR_MAJEUR'
         +                  ' AND W2.WGR_OPEITI=WGAMMERES.WGR_OPEITI'
         +                  ' AND W2.WGR_NUMOPERGAMME=WGAMMERES.WGR_NUMOPERGAMME'
         +                  ' AND W2.WGR_CIRCUIT=WGAMMERES.WGR_CIRCUIT'
         +                  ' AND W2.WGR_FAMRESSOURCE<WGAMMERES.WGR_FAMRESSOURCE)'
          );

// gj 200906  ExecuteSql('UPDATE WGAMMERES SET WGR_NATURERESMES=IIF(WGR_ORDRE=1,"CPT","AUT"), WGR_RESORIGINEMES=""');
  ExecuteSql('UPDATE WGAMMERES SET WGR_NATURERESMES=IIF(WGR_ORDRE=1,"PIL","AUT"), WGR_RESORIGINEMES="" WHERE WGR_RESORIGINEMES IS NULL');
  ExecuteSql('UPDATE WGAOPER SET WGO_TYPEOPERATION="", WGO_SUIVIOPMES="" WHERE WGO_SUIVIOPMES IS NULL');
  ExecuteSql('UPDATE WGAOPERRES SET WGS_ORDRE=1+(SELECT COUNT(*) FROM WGAOPERRES W2'
         +                  ' WHERE W2.WGS_CODEOPERATION=WGAOPERRES.WGS_CODEOPERATION'
         +                  ' AND W2.WGS_RESSOURCE<WGAOPERRES.WGS_RESSOURCE)'
          );

//gj 200906  ExecuteSql('UPDATE WGAOPERRES SET WGS_NATURERESMES=IIF(WGS_ORDRE=1,"CPT","AUT"), WGS_RESORIGINEMES=""');
  ExecuteSql('UPDATE WGAOPERRES SET WGS_NATURERESMES=IIF(WGS_ORDRE=1,"PIL","AUT"), WGS_RESORIGINEMES="" WHERE WGS_RESORIGINEMES IS NULL');
//gerard jugde 200906
  ExecuteSql('UPDATE WGAMMETET'
         + ' SET WGT_TPREVHHCC=ISNULL(iif(WGT_QLOTSTOC=0,1,WGT_QLOTSTOC)*(SELECT SUM(WGL_TPREVHHCC/iif(WGL_QLOTSTOC=0,1,WGL_QLOTSTOC)'
         +       '/iif((GA_QECOPROD=0 OR (WGL_TYPEOPERATION IN ("EXE","PAS"))),1,GA_QECOPROD))'
         +       ' FROM WGAMMELIG'
         +       ' LEFT JOIN ARTICLE ON (GA_ARTICLE=WGL_ARTICLE)'
         +       ' WHERE WGL_NATURETRAVAIL=WGT_NATURETRAVAIL'
         +       ' AND WGL_ARTICLE=WGT_ARTICLE'
         +       ' AND WGL_MAJEUR=WGT_MAJEUR'
         +       ' AND WGL_CIRCUIT="" AND WGL_TYPEOPERATION<>"CPT"'
         +       '),0)'
         );

  ExecuteSql('UPDATE WORDRELIG'
         + ' SET WOL_LIGNEORDREGEN=1+(SELECT COUNT(*)'
         +                          ' FROM WORDRELIG W2'
         +                          ' WHERE WORDRELIG.WOL_IDENTIFIANT>=W2.WOL_IDENTIFIANT),'
         + ' WOL_TPREVHHCC=ISNULL((SELECT SUM(WOG_TPREVHHCC)'
         +                ' FROM WORDREGAMME'
         +                ' WHERE WOG_NATURETRAVAIL=WOL_NATURETRAVAIL'
         +                ' AND WOG_LIGNEORDRE=WOL_LIGNEORDRE),0),'
         + ' WOL_QBACSTDMES=0,WOL_QBACMINMES=0,wol_DateDec = Wol_DateCreation WHERE WOL_LIGNEORDREGEN IS NULL'
          );


  ExecuteSql('INSERT INTO WJETONS'
         + ' (WJT_PREFIXE,WJT_CONTEXTE,WJT_JETON,WJT_CREATEUR,WJT_UTILISATEUR,WJT_DATECREATION,WJT_DATEMODIF)'
         + ' SELECT DISTINCT "WOL" AS WJT_PREFIXE,'
         + ' "GEN" AS WJT_CONTEXTE,'
         + ' (SELECT 1+MAX(WOL_LIGNEORDREGEN) FROM WORDRELIG) AS WJT_JETON,'
         + ' "" AS WJT_CREATEUR,'
         + ' "" AS WJT_UTILISATEUR,'
         + ' "' + UsDateTime(iDate1900) + '" AS WJT_DATECREATION,'
         + ' "' + UsDateTime(iDate1900) + '" AS WJT_DATEMODIF'
         + ' FROM WJETONS WHERE NOT EXISTS (SELECT 1 FROM WJETONS WHERE WJT_PREFIXE="WOL" AND WJT_CONTEXTE="GEN")');


  ExecuteSql('UPDATE WORDREGAMME SET WOG_GROUPAGEMES=0,WOG_SUIVIOPMES="OF",WOG_ANTERIORITEMES="",'
         + ' WOG_RANGOPMES = 1+(SELECT COUNT(*)'
         +                    ' FROM WORDREGAMME W2'
         +                    ' WHERE W2.WOG_NATURETRAVAIL=WORDREGAMME.WOG_NATURETRAVAIL'
         +                    ' AND W2.WOG_LIGNEORDRE=WORDREGAMME.WOG_LIGNEORDRE'
         +                    ' AND (W2.WOG_OPECIRC||W2.WOG_NUMOPERGAMME)<(WORDREGAMME.WOG_OPECIRC||WORDREGAMME.WOG_NUMOPERGAMME)) WHERE WOG_GROUPAGEMES IS NULL'
         );


  Executesql('UPDATE WORDRERES SET WOR_NATURERESMES="",WOR_RESORIGINEMES="",WOR_NBRESSOURCE=1,WOR_UNITEPROD="",WOR_QUALIFUNITESTO="",WOR_COEFPROD=1,WOR_QPRODSAIS=0,WOR_QPRODSTOC=0 WHERE WOR_NATURERESMES IS NULL');

{gj 200906
  ExecuteSql('UPDATE WORDRERES SET WOR_NATURERESMES="CPT"'
         + ' WHERE WOR_NATURERESMES="" AND NOT EXISTS (SELECT 1'
         +                                            ' FROM WORDRERES R2'
         +                                            ' WHERE R2.WOR_NATURETRAVAIL=WORDRERES.WOR_NATURETRAVAIL'
         +                                            ' AND R2.WOR_LIGNEORDRE=WORDRERES.WOR_LIGNEORDRE'
         +                                            ' AND R2.WOR_OPECIRC=WORDRERES.WOR_OPECIRC'
         +                                            ' AND R2.WOR_NUMOPERGAMME=WORDRERES.WOR_NUMOPERGAMME'
         +                                            ' AND R2.WOR_NATURERESMES="CPT")'
         + ' AND WORDRERES.WOR_FAMRESSOURCE=(SELECT MIN(R3.WOR_FAMRESSOURCE)'
         +                                 ' FROM WORDRERES R3'
         +                                 ' WHERE R3.WOR_NATURETRAVAIL=WORDRERES.WOR_NATURETRAVAIL'
         +                                 ' AND R3.WOR_LIGNEORDRE=WORDRERES.WOR_LIGNEORDRE'
         +                                 ' AND R3.WOR_OPECIRC=WORDRERES.WOR_OPECIRC'
         +                                 ' AND R3.WOR_NUMOPERGAMME=WORDRERES.WOR_NUMOPERGAMME'
         +                                 ' AND R3.WOR_NATURERESMES="")'
          );
}
  ExecuteSql('UPDATE WORDRERES SET WOR_NATURERESMES="PIL"'
         + ' WHERE WOR_NATURERESMES="" AND NOT EXISTS (SELECT 1'
         +                                            ' FROM WORDRERES R2'
         +                                            ' WHERE R2.WOR_NATURETRAVAIL=WORDRERES.WOR_NATURETRAVAIL'
         +                                            ' AND R2.WOR_LIGNEORDRE=WORDRERES.WOR_LIGNEORDRE'
         +                                            ' AND R2.WOR_OPECIRC=WORDRERES.WOR_OPECIRC'
         +                                            ' AND R2.WOR_NUMOPERGAMME=WORDRERES.WOR_NUMOPERGAMME'
         +                                            ' AND R2.WOR_NATURERESMES="PIL")'
         + ' AND WORDRERES.WOR_FAMRESSOURCE=(SELECT MIN(R3.WOR_FAMRESSOURCE)'
         +                                 ' FROM WORDRERES R3'
         +                                 ' WHERE R3.WOR_NATURETRAVAIL=WORDRERES.WOR_NATURETRAVAIL'
         +                                 ' AND R3.WOR_LIGNEORDRE=WORDRERES.WOR_LIGNEORDRE'
         +                                 ' AND R3.WOR_OPECIRC=WORDRERES.WOR_OPECIRC'
         +                                 ' AND R3.WOR_NUMOPERGAMME=WORDRERES.WOR_NUMOPERGAMME'
         +                                 ' AND R3.WOR_NATURERESMES="")'
          );

  ExecuteSql('UPDATE WORDRERES SET WOR_NATURERESMES="AUT" WHERE (WOR_NATURERESMES="") OR (WOR_NATURERESMES is null)');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 23/11/2005
Modifié le ... :   /  /
Description .. : Génération du champ GSM_GUIDORI
Mots clefs ... :
*****************************************************************}
procedure Maj_GSM_GUIDORI;
var
  sql: string;
begin
  { Pour les numéros de série }
  sql := 'UPDATE STKMOUVEMENT PHY'
       + ' SET GSM_GUIDORI = (SELECT MAX(GSM_GUID)'
       +                    ' FROM STKMOUVEMENT ATT'
       +                    ' WHERE ATT.GSM_STKTYPEMVT="ATT"'
       +                    ' AND ATT.GSM_QUALIFMVT="APR"'
       +                    ' AND ATT.GSM_DEPOT=PHY.GSM_DEPOT'
       +                    ' AND ATT.GSM_ARTICLE=PHY.GSM_ARTICLE'
       +                    ' AND ATT.GSM_PREFIXEORI=PHY.GSM_PREFIXEORI'
       +                    ' AND ATT.GSM_NATUREORI=PHY.GSM_NATUREORI'
       +                    ' AND ATT.GSM_NUMEROORI=PHY.GSM_NUMEROORI'
       +                    ' AND ATT.GSM_SERIEINTERNE=PHY.GSM_SERIEINTERNE)'
       + ' WHERE GSM_STKTYPEMVT="PHY"'
       + ' AND GSM_QUALIFMVT="EPR"'
       + ' AND GSM_GUIDORI=""'
       + ' AND (SELECT MAX(GSM_GUID)'
       +      ' FROM STKMOUVEMENT ATT'
       +      ' WHERE ATT.GSM_STKTYPEMVT="ATT"'
       +      ' AND ATT.GSM_QUALIFMVT="APR"'
       +      ' AND ATT.GSM_DEPOT=PHY.GSM_DEPOT'
       +      ' AND ATT.GSM_ARTICLE=PHY.GSM_ARTICLE'
       +      ' AND ATT.GSM_PREFIXEORI=PHY.GSM_PREFIXEORI'
       +      ' AND ATT.GSM_NATUREORI=PHY.GSM_NATUREORI'
       +      ' AND ATT.GSM_NUMEROORI=PHY.GSM_NUMEROORI'
       +      ' AND ATT.GSM_SERIEINTERNE=PHY.GSM_SERIEINTERNE)<>""'
       ;
  ExecuteSql(Sql);

  { Pour les numéros de lot }
  sql := 'UPDATE STKMOUVEMENT PHY'
       + ' SET GSM_GUIDORI = (SELECT MAX(GSM_GUID)'
       +                    ' FROM STKMOUVEMENT ATT'
       +                    ' WHERE ATT.GSM_STKTYPEMVT="ATT"'
       +                    ' AND ATT.GSM_QUALIFMVT="APR"'
       +                    ' AND ATT.GSM_DEPOT=PHY.GSM_DEPOT'
       +                    ' AND ATT.GSM_ARTICLE=PHY.GSM_ARTICLE'
       +                    ' AND ATT.GSM_PREFIXEORI=PHY.GSM_PREFIXEORI'
       +                    ' AND ATT.GSM_NATUREORI=PHY.GSM_NATUREORI'
       +                    ' AND ATT.GSM_NUMEROORI=PHY.GSM_NUMEROORI'
       +                    ' AND ATT.GSM_LOTINTERNE=PHY.GSM_LOTINTERNE)'
       + ' WHERE GSM_STKTYPEMVT="PHY"'
       + ' AND GSM_QUALIFMVT="EPR"'
       + ' AND GSM_GUIDORI=""'
       + ' AND (SELECT MAX(GSM_GUID)'
       +      ' FROM STKMOUVEMENT ATT'
       +      ' WHERE ATT.GSM_STKTYPEMVT="ATT"'
       +      ' AND ATT.GSM_QUALIFMVT="APR"'
       +      ' AND ATT.GSM_DEPOT=PHY.GSM_DEPOT'
       +      ' AND ATT.GSM_ARTICLE=PHY.GSM_ARTICLE'
       +      ' AND ATT.GSM_PREFIXEORI=PHY.GSM_PREFIXEORI'
       +      ' AND ATT.GSM_NATUREORI=PHY.GSM_NATUREORI'
       +      ' AND ATT.GSM_NUMEROORI=PHY.GSM_NUMEROORI'
       +      ' AND ATT.GSM_LOTINTERNE=PHY.GSM_LOTINTERNE)<>""'
       ;
  ExecuteSql(Sql);

  { Pour les attendus qui restent }
  sql := 'UPDATE STKMOUVEMENT PHY'
       + ' SET GSM_GUIDORI = (SELECT MAX(GSM_GUID)'
       +                    ' FROM STKMOUVEMENT ATT'
       +                    ' WHERE ATT.GSM_STKTYPEMVT="ATT"'
       +                    ' AND ATT.GSM_QUALIFMVT="APR"'
       +                    ' AND ATT.GSM_DEPOT=PHY.GSM_DEPOT'
       +                    ' AND ATT.GSM_ARTICLE=PHY.GSM_ARTICLE'
       +                    ' AND ATT.GSM_PREFIXEORI=PHY.GSM_PREFIXEORI'
       +                    ' AND ATT.GSM_NATUREORI=PHY.GSM_NATUREORI'
       +                    ' AND ATT.GSM_NUMEROORI=PHY.GSM_NUMEROORI)'
       + ' WHERE GSM_STKTYPEMVT="PHY"'
       + ' AND GSM_QUALIFMVT="EPR"'
       + ' AND GSM_GUIDORI=""'
       + ' AND (SELECT MAX(GSM_GUID)'
       +      ' FROM STKMOUVEMENT ATT'
       +      ' WHERE ATT.GSM_STKTYPEMVT="ATT"'
       +      ' AND ATT.GSM_QUALIFMVT="APR"'
       +      ' AND ATT.GSM_DEPOT=PHY.GSM_DEPOT'
       +      ' AND ATT.GSM_ARTICLE=PHY.GSM_ARTICLE'
       +      ' AND ATT.GSM_PREFIXEORI=PHY.GSM_PREFIXEORI'
       +      ' AND ATT.GSM_NATUREORI=PHY.GSM_NATUREORI'
       +      ' AND ATT.GSM_NUMEROORI=PHY.GSM_NUMEROORI)<>""'
       ;
   ExecuteSql(Sql);

  { Pour les réservés qui restent }
  sql := 'UPDATE STKMOUVEMENT PHY'
       + ' SET GSM_GUIDORI = (SELECT MAX(GSM_GUID)'
       +                    ' FROM STKMOUVEMENT RES'
       +                    ' WHERE RES.GSM_STKTYPEMVT="RES"'
       +                    ' AND RES.GSM_QUALIFMVT="RPR"'
       +                    ' AND RES.GSM_DEPOT=PHY.GSM_DEPOT'
       +                    ' AND RES.GSM_ARTICLE=PHY.GSM_ARTICLE'
       +                    ' AND RES.GSM_PREFIXEORI=PHY.GSM_PREFIXEORI'
       +                    ' AND RES.GSM_NATUREORI=PHY.GSM_NATUREORI'
       +                    ' AND RES.GSM_NUMEROORI=PHY.GSM_NUMEROORI)'
       + ' WHERE GSM_STKTYPEMVT="PHY"'
       + ' AND GSM_QUALIFMVT="SPR"'
       + ' AND GSM_GUIDORI=""'
       + ' AND (SELECT MAX(GSM_GUID)'
       +      ' FROM STKMOUVEMENT RES'
       +      ' WHERE RES.GSM_STKTYPEMVT="RES"'
       +      ' AND RES.GSM_QUALIFMVT="RPR"'
       +      ' AND RES.GSM_DEPOT=PHY.GSM_DEPOT'
       +      ' AND RES.GSM_ARTICLE=PHY.GSM_ARTICLE'
       +      ' AND RES.GSM_PREFIXEORI=PHY.GSM_PREFIXEORI'
       +      ' AND RES.GSM_NATUREORI=PHY.GSM_NATUREORI'
       +      ' AND RES.GSM_NUMEROORI=PHY.GSM_NUMEROORI)<>""'
       ;
   ExecuteSql(Sql);

  //Jean-Luc SAUZET Le 25/05/2008 Version 8.1.850.80 Demande N° 2235
  { Pour le Reste }
  sql := 'UPDATE STKMOUVEMENT'
       + ' SET GSM_GUIDORI = ""'
       + ' WHERE GSM_GUIDORI IS NULL'
       ;

  ExecuteSql(Sql);

end;

{***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 12/12/2005
Modifié le ... :   /  /    
Description .. : fct qui bascule, pour le bureau, la table annuinterloc dans la 
Suite ........ : table contact
Suite ........ : la table annuinterloc sera ensuite supprimée.
Mots clefs ... : ANNUINTERLOC;CONTACT
*****************************************************************}
procedure SuppAnnuInterloc ;
Var TobAnu, TobContact,TobDet,TobDetCon : tob;
  StSql : string;
  ii    : integer;
begin
  if Not ExisteSql ('Select ani_datecreation from ANNUINTERLOC') then exit;
     //obligation de tout prendre, car changement de  table
    // stSql := 'Select annuinterloc.* ,ann_tiers,ann_natureauxi,jtc_ttcivilite from ANNUINTERLOC left join ANNUAIRE on ann_GUIDper=ani_GUIDper left join JUTYPECIVIL on jtc_typeciv=ani_cv';
  stSql := 'Select annuinterloc.* ,ann_tiers,ann_natureauxi,ann_guidper,jtc_ttcivilite from ANNUINTERLOC left join ANNUAIRE on ann_codeper=ani_codeper left join JUTYPECIVIL on jtc_typeciv=ani_cv';
  TobAnu := Tob.create('annu_interloc', nil, -1);
  TobAnu.LoadDetailDBFromSql('annu_interloc',stSql);
  TobContact := Tob.create ('les contacts',nil,-1);
  for ii:=0 to TobAnu.detail.count-1 do
    begin
    tobdet:=TobAnu.detail[ii];
    TobDetCon := TOB.Create ('CONTACT', TobContact ,-1) ;
    if Tobdet.getvalue('ANN_TIERS') <> '' then
      begin
      TobDetCon.putvalue('C_NUMEROCONTACT',TobDet.GetInteger('ANI_NOCONT')+100); //pournepas ecraser éventule contatc tiers, on met +100
      TobDetCon.putvalue('C_TYPECONTACT','T');
      TobDetCon.putvalue('C_TIERS',Tobdet.getString('ANN_TIERS'));
      TobDetCon.putvalue('C_AUXILIAIRE',TiersAuxiliaire(Tobdet.getString('ANN_TIERS'),false));
      TobDetCon.putvalue('C_NATUREAUXI',Tobdet.getString('ANN_NATUREAUXI'));
      end
    else begin // cas tiers n'existe pas
      TobDetCon.putValue('C_NUMEROCONTACT',TobDet.GetInteger('ANI_NOCONT'));
      TobDetCon.putvalue('C_TYPECONTACT','ANN');
      TobDetCon.putvalue('C_AUXILIAIRE',Tobdet.getValue('ANN_GUIDPER'));
      end;
      //copie zones commune
    TobDetCon.putvalue('C_GUIDPER',Tobdet.getValue('ANN_GUIDPER'));
    TobDetCon.putvalue('C_PRINCIPAL',Tobdet.getValue('ANI_PRINCIPAL'));
    TobDetCon.putvalue('C_NOM',Tobdet.getValue('ANI_NOM'));
    TobDetCon.putvalue('C_PRENOM',Tobdet.getValue('ANI_PRENOM'));
    TobDetCon.putvalue('C_CIVILITE',Tobdet.getValue('JTC_TTCIVILITE'));
    TobDetCon.putvalue('C_TEXTELIBRE3',Tobdet.getValue('ANI_CV'));
    TobDetCon.putvalue('C_TELEPHONE',Tobdet.getValue('ANI_TEL1'));
    TobDetCon.putvalue('C_CLETELEPHONE',Tobdet.getValue('ANI_CLETELEPHONE'));
    TobDetCon.putvalue('C_RVA',Tobdet.getValue('ANI_EMAIL'));
    TobDetCon.putvalue('C_FAX',Tobdet.getValue('ANI_FAX'));
    TobDetCon.putvalue('C_BLOCNOTE',Tobdet.getValue('ANI_NOTECON'));
    TobDetCon.putvalue('C_SERVICE',Tobdet.getValue('ANI_SERVICE'));
    TobDetCon.putvalue('C_FONCTION',Tobdet.getValue('ANI_FONCTIONJ'));
    TobDetCon.putvalue('C_TELEX',Tobdet.getValue('ANI_TEL2'));
    TobDetCon.putvalue('C_DATECREATION',Tobdet.getValue('ANI_DATECREATION'));
    TobDetCon.putvalue('C_UTILISATEUR',Tobdet.getValue('ANI_UTILISATEUR'));
    TobDetCon.putvalue('C_TEXTELIBRE1',Tobdet.getValue('ANI_ASSIST'));
    TobDetCon.putvalue('C_TEXTELIBRE2',Tobdet.getValue('ANI_TELASSIST'));
    end;
  TobContact.SetallModifie(true); //pour OK blob
  TobContact.InsertOrUpdateDB (False); // ecriture table contact
  TobContact.free;
  TobAnu.free;

end;

procedure MajMarkMESApres;
var
  i : integer ;
  //-- Libellé libres de la table CATALOGU //
Begin
  for i := 1 to 10 do
  begin
    if i < 3 then
    begin
      InsertChoixExt('GCX', 'TGCA_DATELIBRE' + IntToStr(i), 'Date libre '     + IntToStr(i), 'Date libre '     + IntToStr(i), '');
      InsertChoixExt('GCW', 'TGCA_VALLIBRE'  + IntToStr(i), 'Valeur libre '   + IntToStr(i), 'Valeur libre '   + IntToStr(i), '');
      InsertChoixExt('GCY', 'TGCA_BOOLLIBRE' + IntToStr(i), 'Décision libre ' + IntToStr(i), 'Décision libre ' + IntToStr(i), '');
      InsertChoixExt('GCZ', 'TGCA_CHARLIBRE' + IntToStr(i), 'Texte libre '    + IntToStr(i), 'Texte libre '    + IntToStr(i), '');
    end;
    Case i of
      1..9: InsertChoixExt('GCU', 'TGCA_LIBREGCA' + IntToStr(i), 'Table libre ' + IntToStr(i), 'Table libre ' + IntToStr(i), '');
      10  : InsertChoixExt('GCU', 'TGCA_LIBREGCAA'             , 'Table libre 10'            , 'Table libre 10'            , '');
    end;
  end;

  //-- Libellé libres MES //
  for i := 1 to 10 do
  begin
    if i < 3 then
    begin
      InsertChoixExt('Q1X', 'TQWB_DATELIBRE' + IntToStr(i), 'Date libre '     + IntToStr(i), 'Date libre '     + IntToStr(i), '');
      InsertChoixExt('Q1W', 'TQWB_VALLIBRE'  + IntToStr(i), 'Valeur libre '   + IntToStr(i), 'Valeur libre '   + IntToStr(i), '');
      InsertChoixExt('Q1Y', 'TQWB_BOOLLIBRE' + IntToStr(i), 'Décision libre ' + IntToStr(i), 'Décision libre ' + IntToStr(i), '');
      InsertChoixExt('Q1Z', 'TQWB_CHARLIBRE' + IntToStr(i), 'Texte libre '    + IntToStr(i), 'Texte libre '    + IntToStr(i), '');
    end;
    Case i of
      1..9: InsertChoixExt('Q1V', 'TQWB_LIBREQWB' + IntToStr(i), 'Table libre ' + IntToStr(i), 'Table libre ' + IntToStr(i), '');
      10  : InsertChoixExt('Q1V', 'TQWB_LIBREQWBA'             , 'Table libre 10'            , 'Table libre 10'            , '');
    end;
  end;
end;

procedure RT_InsertTablettesWIV;
var
  iInc: integer;
  iNum,sPrefixe,sCode,sType,sLibelle : string;
begin
      sPrefixe := 'WIV';
      iNum:='B';
         for iInc := 1 to 3 do
         begin
            //-- Booleens libres //
            sType    := 'W' + iNum + 'Y' ;
            sCode    := 'T' + sPrefixe + '_BOOLLIBRE' + IntToStr(iInc) ;
            sLibelle := 'Décision libre ' + IntToStr(iInc);
            InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
            //-- Char libres --//
            sType    := 'W' + iNum + 'Z' ;
            sCode    := 'T' + sPrefixe + '_CHARLIBRE' + IntToStr(iInc) ;
            sLibelle := 'Texte libre ' + IntToStr(iInc);
            InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
            //-- Dates libres --//
            sType    := 'W' + iNum + 'X' ;
            sCode    := 'T' + sPrefixe + '_DATELIBRE' + IntToStr(iInc) ;
            sLibelle := 'Date libre ' + IntToStr(iInc);
            InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
            //-- Valeur libres --//
            sType    := 'W' + iNum + 'W' ;
            sCode    := 'T' + sPrefixe + '_VALLIBRE' + IntToStr(iInc) ;
            sLibelle := 'Valeur libre ' + IntToStr(iInc);
            InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
         end;
         //-- Libellés libres --//
         for iInc := 1 to 10 do
         begin
            case iInc of
            1..9 : begin
                     sType    := 'W' + iNum + 'V' ;
                     sCode    := 'T' + sPrefixe + '_LIBRE'  + sPrefixe + IntToStr(iInc) ;
                     sLibelle := 'Table libre ' + IntToStr(iInc);
                     InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
                  end;
            10  : begin
                     sType    := 'W' + iNum + 'V' ;
                     sCode    := 'T' + sPrefixe + '_LIBRE' + sPrefixe +  'A' ;
                     sLibelle := 'Table libre 10';
                     InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
                  end;
            end;
         end;
end;

procedure GCTraite_TIERSIMPACTPIECE;
var DateFin, Flux : string;
 
    function IsExiste(Element : string) : boolean;
    begin
      Result := ExisteSQL('SELECT * FROM TIERSIMPACTPIECE WHERE GTI_DATEPARAM = "'+ DateFin
                        + '" AND GTI_FLUX = "'+ Flux +'" AND GTI_ELEMENTFC = "'+ Element +'"');
    end;
 
    procedure ExecuteInsert(NatTiers, Element : string);
    begin
      ExecuteSql ('INSERT INTO TIERSIMPACTPIECE (GTI_FLUX, GTI_TYPENATTIERS, GTI_ELEMENTFC, GTI_DATEPARAM)'
                + ' VALUES("' + Flux + '", "' + NatTiers + '", "' + Element +'", "' + DateFin + '")');
    end;
 
begin
  DateFin := UsDateTime(iDate2099);
  Flux := 'VEN';
  if not IsExiste('AUX') then ExecuteInsert('FAC', 'AUX');
  if not IsExiste('CNF') then ExecuteInsert('FAC', 'CNF');
  if not IsExiste('CNS') then ExecuteInsert('FAC', 'CNS');
  if not IsExiste('COL') then ExecuteInsert('FAC', 'COL');
  if not IsExiste('DEV') then ExecuteInsert('PRI', 'DEV');
  if not IsExiste('ENC') then ExecuteInsert('PRI', 'ENC');
  if not IsExiste('ESC') then ExecuteInsert('PRI', 'ESC');
  if not IsExiste('FAR') then ExecuteInsert('FAC', 'FAR');
  if not IsExiste('FCO') then ExecuteInsert('PRI', 'FCO');
  if not IsExiste('FHT') then ExecuteInsert('PRI', 'FHT');
  if not IsExiste('JPA') then ExecuteInsert('PRI', 'JPA');
  if not IsExiste('LIB') then ExecuteInsert('FAC', 'LIB');
  if not IsExiste('MRR') then ExecuteInsert('PRI', 'MRR');
  if not IsExiste('MRT') then ExecuteInsert('PRI', 'MRT');
  if not IsExiste('REG') then ExecuteInsert('PRI', 'REG');
  if not IsExiste('REM') then ExecuteInsert('PRI', 'REM');
  if not IsExiste('RFI') then ExecuteInsert('PRI', 'RFI');
  if not IsExiste('RIB') then ExecuteInsert('FAC', 'RIB');
  if not IsExiste('TPF') then ExecuteInsert('PRI', 'TPF');
  Flux := 'ACH';
  if not IsExiste('AUX') then ExecuteInsert('FAC', 'AUX');
  if not IsExiste('CNF') then ExecuteInsert('FAC', 'CNF');
  if not IsExiste('CNS') then ExecuteInsert('FAC', 'CNS');
  if not IsExiste('COL') then ExecuteInsert('FAC', 'COL');
  if not IsExiste('DEV') then ExecuteInsert('PRI', 'DEV');
  if not IsExiste('ENC') then ExecuteInsert('PRI', 'ENC');
  if not IsExiste('ESC') then ExecuteInsert('PRI', 'ESC');
  if not IsExiste('FAR') then ExecuteInsert('FAC', 'FAR');
  if not IsExiste('FCO') then ExecuteInsert('PRI', 'FCO');
  if not IsExiste('FHT') then ExecuteInsert('PRI', 'FHT');
  if not IsExiste('JPA') then ExecuteInsert('PRI', 'JPA');
  if not IsExiste('LIB') then ExecuteInsert('FAC', 'LIB');
  if not IsExiste('MRR') then ExecuteInsert('PRI', 'MRR');
  if not IsExiste('MRT') then ExecuteInsert('PRI', 'MRT');
  if not IsExiste('REG') then ExecuteInsert('PRI', 'REG');
  if not IsExiste('REM') then ExecuteInsert('PRI', 'REM');
  if not IsExiste('RFI') then ExecuteInsert('PRI', 'RFI');
  if not IsExiste('RIB') then ExecuteInsert('FAC', 'RIB');
  if not IsExiste('TPF') then ExecuteInsert('PRI', 'TPF');
end;

//kb 270706
procedure InsertTablettesRQ;
Var
	iNum, iInc:integer;
  sPrefixe, sType, sCode, sLibelle: string;
begin
  //-- Champs libres //
  for iNum:=0 to 3 do
  begin
     case iNum of
        {Laurent Abélard le 22/11/2007 Version 8.1.850.15 FQ14585
        0 : sPrefixe := 'RQD';
        1 : sPrefixe := 'RQN';
        2 : sPrefixe := 'RQP';}
        0 : sPrefixe := 'RQN';
        1 : sPrefixe := 'RQP';
        2 : sPrefixe := 'RQD';
     end ;
     for iInc := 1 to 3 do
     begin
        //-- Booleens libres //
        sType    := 'R' + IntToStr(iNum) + 'Y' ;
        sCode    := 'T' + sPrefixe + '_BOOLLIBRE' + IntToStr(iInc) ;
        sLibelle := 'Décision libre ' + IntToStr(iInc);
        InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
        //-- Char libres --//
        sType    := 'R' + IntToStr(iNum) + 'Z' ;
        sCode    := 'T' + sPrefixe + '_CHARLIBRE' + IntToStr(iInc) ;
        sLibelle := 'Texte libre ' + IntToStr(iInc);
        InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
        //-- Dates libres --//
        sType    := 'R' + IntToStr(iNum) + 'X' ;
        sCode    := 'T' + sPrefixe + '_DATELIBRE' + IntToStr(iInc) ;
        sLibelle := 'Date libre ' + IntToStr(iInc);
        InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
        //-- Valeur libres --//
        sType    := 'R' + IntToStr(iNum) + 'W' ;
        sCode    := 'T' + sPrefixe + '_VALLIBRE' + IntToStr(iInc) ;
        sLibelle := 'Valeur libre ' + IntToStr(iInc);
        InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
     end;
     //-- Libellés libres --//
     for iInc := 1 to 10 do
     begin
        case iInc of
        1..9 : begin
                 sType    := 'R' + IntToStr(iNum) + 'V' ;
                 sCode    := 'T' + sPrefixe + '_LIBRE'  + sPrefixe + IntToStr(iInc) ;
                 sLibelle := 'Table libre ' + IntToStr(iInc);
                 InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
              end;
        10  : begin
                 sType    := 'R' + IntToStr(iNum) + 'V' ;
                 sCode    := 'T' + sPrefixe + '_LIBRE' + sPrefixe +  'A' ;
                 sLibelle := 'Table libre 10';
                 InsertChoixExt (sType,sCode,sLibelle,sLibelle,'');
              end;
        end;
     end;
  end;
end;



// D.Koza
procedure wConsolideCalculParam;
var
  TWCP  : Tob;
  sQl   : string;
begin
  // test fait en amont
  //if not IsDossierPCL then
  //begin
    if not ExisteSQL('SELECT 1 FROM WCALCULPARAM') then
    begin
      if ExisteSQL('SELECT 1 FROM WCBNPARAM') then
      begin
        sQl := 'SELECT IIF(QSM_MODESIMU="2","CBN","CBB") AS WPP_ORIGINE, WCP_CTX AS WPP_CTX,'
             + ' WCP_DEBUTCALCUL AS WPP_DEBUTCALCUL, WCP_FINCALCUL AS WPP_FINCALCUL, WCP_DATECREATION AS WPP_DATECREATION,'
             + ' WCP_DATEMODIF AS WPP_DATEMODIF, WCP_CREATEUR AS WPP_CREATEUR, WCP_UTILISATEUR AS WPP_UTILISATEUR,'
             + ' WCP_SOCIETE AS WPP_SOCIETE, WCP_BLOCNOTE AS WPP_BLOCNOTE'
             + ' FROM WCBNPARAM, QSIMULATION'
             + ' WHERE QSM_CTX="0" AND TRIM(QSM_CODESIMU)=TRIM(WCP_CTX)';
        TWCP := Tob.Create('WCP', nil, -1);
        try
          if TWCP.LoadDetailDBFromSql('WCALCULPARAM', Sql) then
          begin
            TWCP.SetAllModifie(True);    // Obligatoire pour que le Memo se duplique
            TWCP.InsertDB(nil, False);
          end;
        finally
          TWCP.Free;
        end;
      end;
      if ExisteSQL('SELECT 1 FROM WSCMPARAM') then
      begin
        sQl := 'SELECT "SCM" AS WPP_ORIGINE, WSP_CTX AS WPP_CTX,'
             + ' WSP_DEBUTCALCUL AS WSP_DEBUTCALCUL, WSP_FINCALCUL AS WPP_FINCALCUL, WSP_DATECREATION AS WPP_DATECREATION,'
             + ' WSP_DATEMODIF AS WPP_DATEMODIF, WSP_CREATEUR AS WPP_CREATEUR, WSP_UTILISATEUR AS WPP_UTILISATEUR,'
             + ' WSP_SOCIETE AS WPP_SOCIETE, WSP_BLOCNOTE AS WPP_BLOCNOTE'
             + ' FROM WSCMPARAM';
        TWCP := Tob.Create('WCP', nil, -1);
        try
          if TWCP.LoadDetailDBFromSql('WCALCULPARAM', Sql) then
          begin
            TWCP.SetAllModifie(True);    // Obligatoire pour que le Memo se duplique
            TWCP.InsertDB(nil, False);
          end;
        finally
          TWCP.Free;
        end;
      end;
    end;
  //end;
end;

// JTR

function UpdateAcpteCptaDiff2 : boolean ;
var TobAcomptesUpd, TobTmp : TOB;
    Cpt, Nbre : integer;
    LaRefPiece, Flux, NatureCpta : string;
    Nature, Souche : string;
    NumP, Indice : integer;
    Qry : TQuery;
begin
  { Charge les acomptes en TOB}
  TobAcomptesUPD := TOB.Create('ACOMPTES', nil, -1);
  Qry := OpenSQL('SELECT ACOMPTES.*, PIECE.GP_DATEPIECE AS DATEPIECE, PIECE.GP_VENTEACHAT AS FLUX FROM ACOMPTES '
               + 'INNER JOIN PIECE ON GP_NATUREPIECEG=GAC_NATUREPIECEG AND GP_SOUCHE=GAC_SOUCHE AND GP_NUMERO=GAC_NUMERO AND GP_INDICEG=GAC_INDICEG', true);
  TobAcomptesUPD.LoadDetailDB ('ACOMPTES', '', '', Qry, True);
  Ferme(Qry);
  Nbre := 0;
  Flux := '';
  Nature := '';
  Souche := '';
  NumP := 0;
  Indice := 0;
  TobAcomptesUPD.Detail.Sort('GAC_NATUREPIECEG;GAC_SOUCHE;GAC_NUMERO;GAC_INDICEG;GAC_JALECR;GAC_NUMECR');
  for Cpt := 0 to TobAcomptesUPD.detail.count -1 do
  begin
    TobTmp := TobAcomptesUPD.Detail[Cpt];
    Flux := TobTmp.GetString('FLUX');
    Nature := TobTmp.GetString('GAC_NATUREPIECEG');
    Souche := TobTmp.GetString('GAC_SOUCHE');
    NumP := TobTmp.GetInteger('GAC_NUMERO');
    Indice := TobTmp.GetInteger('GAC_INDICEG');
    Nbre := TobTmp.GetInteger('GAC_NUMORDRE');
    if Flux <> '' then
    begin
      { Test si comptabilisé }
      LaRefPiece := Nature + ';' + Souche + ';' + FormatDateTime('ddmmyyyy',TobTmp.GetValue('DATEPIECE')) + ';'
                    + IntToStr(NumP) + ';' + IntToStr(Indice)+';';
      if Flux = 'VEN' then
      begin
        if TobTmp.GetBoolean('GAC_ISREGLEMENT') then
          NatureCpta := 'RC'
          else
          NatureCpta := 'OC';
      end else
      if Flux = 'ACH' then
      begin
        if TobTmp.GetBoolean('GAC_ISREGLEMENT') then
          NatureCpta := 'RF'
          else
          NatureCpta := 'OF';
      end;
      { Mise à jour de l'acompte }
      ExecuteSQL('UPDATE ACOMPTES SET GAC_CPTADIFF'
               + '                = IIF((SELECT 1 FROM COMPTADIFFEREE WHERE GCD_REFPIECE = "' + LaRefPiece + '"'
               + '                      AND GCD_NATURECOMPTA="' + NatureCpta + '" AND GCD_RANG = ' + IntToStr(Nbre)
               + '                      AND GCD_COMPTABILISE= "-")=1,"X","-")'
               + ' WHERE GAC_NATUREPIECEG = "' + Nature + '" AND GAC_SOUCHE = "' + Souche + '" AND GAC_NUMERO = ' + IntToStr(NumP)
               + ' AND GAC_INDICEG = ' + IntToStr(Indice) + ' AND GAC_JALECR = "' + TobTmp.GetString('GAC_JALECR') + '"'
               + ' AND GAC_NUMECR = ' + IntToStr(TobTmp.GetInteger('GAC_NUMECR')));
      { Delete la compta diff }
      ExecuteSQL('DELETE FROM COMPTADIFFEREE WHERE GCD_REFPIECE = "' + LaRefPiece + '"'
               + ' AND GCD_NATURECOMPTA="' + NatureCpta + '" AND GCD_RANG = ' + IntToStr(Nbre));
    end;
  end;
  { Suppression de la compta différée déjà transmise }
  ExecuteSQL('DELETE FROM COMPTADIFFEREE WHERE GCD_NATURECOMPTA IN("OC", "RC", "OF", "RF") AND GCD_COMPTABILISE = "X"');
end;

procedure MajEngagement;
Begin
 // DEBUT MAJ QBPBIBLIOAXE MARCHE ENGAGEMENT
  ExecuteSqlNoPCl ('DELETE FROM QBPBIBLIOAXE WHERE QBX_CODEMARCHE="#PGIENGAGE"');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2, QBX_NOMTABLETTE,QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION, '+
  'QBX_DATEMODIF, QBX_CREATEUR,QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("003","Etablissement","#PGIENGAGE","ETABLISS","ET_ETABLISSEMENT","ET_LIBELLE","","","TTETABLISSEMENT","GL_ETABLISSEMENT","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","GCMULETABLISS","DAL_ETABLISSEMENT","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION, '+
  'QBX_DATEMODIF,QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("004","BU","#PGIENGAGE","CHOIXCOD","CC_CODE","CC_LIBELLE","CC_TYPE","PSQ","PGCODESTAT","HIERARCHIE","","","","","X",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPPCODESTAT","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("005","Sous BU","#PGIENGAGE","CHOIXCOD","CC_CODE","CC_LIBELLE","CC_TYPE","PL4","PGLIBREPCMB4","HIERARCHIE","","","","","X",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPPLIBRE4","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2, QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("006","Organisation","#PGIENGAGE","CHOIXCOD","CC_CODE","CC_LIBELLE","CC_TYPE","PL1","PGLIBREPCMB1","HIERARCHIE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPPLIBRE1","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("007","Tablette libre 2","#PGIENGAGE","CHOIXCOD","CC_CODE","CC_LIBELLE","CC_TYPE","PL2","PGLIBREPCMB2","HIERARCHIE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPPLIBRE2","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("008","Tablette libre 3","#PGIENGAGE","CHOIXCOD","CC_CODE","CC_LIBELLE","CC_TYPE","PL3","PGLIBREPCMB3","HIERARCHIE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPPLIBRE3","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("009","Lieu de travail 1","#PGIENGAGE","CHOIXCOD","CC_CODE","CC_LIBELLE","CC_TYPE","PAG","PGTRAVAILN1", "HIERARCHIE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPPWORKN1","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("010","Lieu de travail 2","#PGIENGAGE","CHOIXCOD","CC_CODE","CC_LIBELLE","CC_TYPE","PST","PGTRAVAILN2","HIERARCHIE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPPWORKN2","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF,QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("011","Lieu de travail 3","#PGIENGAGE","CHOIXCOD","CC_CODE","CC_LIBELLE","CC_TYPE","PUN","PGTRAVAILN3","HIERARCHIE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPPWORKN3","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES ' +
  '("012","Lieu de travail 4","#PGIENGAGE","CHOIXCOD","CC_CODE","CC_LIBELLE","CC_TYPE","PSV","PGTRAVAILN4","HIERARCHIE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPPWORKN4","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("013","Libre hierarchie 1","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LH1","GCLIBREHIE1","HIERARCHIE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("014","Libre hierarchie 2","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LH2","GCLIBREHIE2","HIERARCHIE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  ' QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("015","Libre hierarchie 3","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LH3","GCLIBREHIE3","HIERARCHIE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("016","Libre hierarchie 4","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LH4","GCLIBREHIE4","HIERARCHIE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,' +
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES ' +
  '("017","Libre hierarchie 5","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LH5","GCLIBREHIE5","HIERARCHIE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES ' +
  '("018","Libre hierarchie 6","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LH6","GCLIBREHIE6","HIERARCHIE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("019","Libre hierarchie 7","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LH7","GCLIBREHIE7","HIERARCHIE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  ' QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("020","Libre hierarchie 8","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LH8","GCLIBREHIE8","HIERARCHIE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION, ' +
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("021","Libre hierarchie 9","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LH9","GCLIBREHIE9","HIERARCHIE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  ' QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("022","Libre hierarchie 10","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LHA","GCLIBREHIEA","HIERARCHIE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("024","LIBRE ARTICLE 1","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LA1","GCLIBREART1","GL_LIBREART1","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","DAL_LIBREART1","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("025","LIBRE ARTICLE 2","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LA2","GCLIBREART2","GL_LIBREART2","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","DAL_LIBREART2","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("026","LIBRE ARTICLE 3","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LA3","GCLIBREART3","GL_LIBREART3","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","DAL_LIBREART3","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("027","LIBRE ARTICLE 4","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LA4","GCLIBREART4","GL_LIBREART4","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","DAL_LIBREART4","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("028","LIBRE ARTICLE 5","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LA5","GCLIBREART5","GL_LIBREART5","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","DAL_LIBREART5","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("029","LIBRE ARTICLE 6","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LA6","GCLIBREART6","GL_LIBREART6","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","DAL_LIBREART6","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("030","LIBRE ARTICLE 7","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LA7","GCLIBREART7","GL_LIBREART7","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","DAL_LIBREART7","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("033","LIBRE ARTICLE 10","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LAA","GCLIBREARTA", "GL_LIBREARTA","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","DAL_LIBREARTA","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION, '+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("034","Article","#PGIENGAGE","ARTICLE","GA_ARTICLE","GA_LIBELLE","","","GCARTICLE", "GL_ARTICLE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","GCMULARTINFOS","DAL_ARTICLE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES ' +
  '("035","Salariés","#PGIENGAGE","SALARIES","PSA_SALARIE","PSA_LIBELLE","","","PGSALARIE","HIERARCHIE","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","PGSALARIES","HIERARCHIE","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  ' QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("046","Type de DA","#PGIENGAGE","TYPEDA","DAT_TYPEDA","DAT_LIBELLE","","","TTTYPEDA","GLC_TYPEDA",'+
  '"LEFT JOIN LIGNECOMPL & ON &.GLC_NATUREPIECEG=GL_NATUREPIECEG AND &.GLC_SOUCHE=GL_SOUCHE AND &.GLC_NUMERO=GL_NUMERO '+  'AND &.GLC_INDICEG=GL_INDICEG AND &.GLC_NUMORDRE=GL_NUMORDRE	", "","","","",'+  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","GCTYPEDA","DAL_TYPEDA","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("047","Service","#PGIENGAGE","SERVICES","PGS_CODESERVICE","PGS_NOMSERVICE","","","PGSERVICE","GLC_CODESERVICE",'+
  '"LEFT JOIN LIGNECOMPL & ON &.GLC_NATUREPIECEG=GL_NATUREPIECEG AND &.GLC_SOUCHE=GL_SOUCHE AND &.GLC_NUMERO=GL_NUMERO AND &.GLC_INDICEG=GL_INDICEG'+  ' AND &.GLC_NUMORDRE=GL_NUMORDRE	","","","","",'+'"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) +  '","","000","PGSERVICES","DA_CODESERVICE","LEFT JOIN PIECEDA & ON &.DA_SOUCHE=DAL_SOUCHE AND &.DA_NUMERO=DAL_NUMERO AND &.DA_TYPEDA=DAL_TYPEDA")');
  ExecuteSQLNoPCL ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES ' +
  '("048","FAMILLE NIVEAU 1","#PGIENGAGE","CHOIXCOD","CC_CODE","CC_LIBELLE","CC_TYPE","FN1","GCFAMILLENIV1","GL_FAMILLENIV1","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXCOD","DAL_FAMILLENIV1","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  ' QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+  '("032","LIBRE ARTICLE 9","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LA9","GCLIBREART9","GL_LIBREART9","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","DAL_LIBREART9","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("031","LIBRE ARTICLE 8","#PGIENGAGE","CHOIXEXT","YX_CODE","YX_LIBELLE","YX_TYPE","LA8","GCLIBREART8","GL_LIBREART8","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXEXT","DAL_LIBREART8","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2,QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,' +
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES '+
  '("049","FAMILLE NIVEAU 2","#PGIENGAGE","CHOIXCOD","CC_CODE","CC_LIBELLE","CC_TYPE","FN2","GCFAMILLENIV2", "GL_FAMILLENIV2","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXCOD","DAL_FAMILLENIV2","")');
  ExecuteSqlNoPCl ('INSERT INTO QBPBIBLIOAXE (QBX_CODEAXE, QBX_LIBAXE, QBX_CODEMARCHE, QBX_NOMTABLEAXE, QBX_CHPCODE, QBX_CHPLIB, QBX_NOMCHP2,'+
  'QBX_VALCHP2, QBX_NOMTABLETTE, QBX_NOMCHPREQUETE, QBX_LEFTJOINR, QBX_COMPLWHERE, QBX_FAMILLECRIT, QBX_NUMCRIT, QBX_ORDRCRIT, QBX_DATECREATION,'+
  'QBX_DATEMODIF, QBX_CREATEUR, QBX_UTILISATEUR, QBX_LISTE, QBX_NOMCHPPR, QBX_LEFTJOINRPR) VALUES ' +
  '("050","FAMILLE NIVEAU 3","#PGIENGAGE","CHOIXCOD","CC_CODE","CC_LIBELLE","CC_TYPE","FN3","GCFAMILLENIV3","GL_FAMILLENIV3","","","","","",'+
  '"' + UsDateTime_(iDate1900) + '","' + UsDateTime_(iDate1900) + '","","000","QULBPCHOIXCOD","DAL_FAMILLENIV3","")');
  // FIN MAJ QBPBIBLIOAXE MARCHE ENGAGEMENT
End;

procedure UpdateColRTYPEIDBQ;
  {Provient de UtilRIB}
  Function ErreurDansIban(Rib : String) : boolean;
  var
    St2, St4, cleIBAN, cleIBAN2, ret, strInter : String ;
    ii : Byte ;
    cleL, i : integer ;
  begin
    Result:=True;
    CleIban2 := Copy(Rib,3,2);
    St2 := Copy(RIB,5,length(Rib)-4) + Copy(RIB,1,2) + '00' ;
    if Length(St2)<10 then exit ;
    St2:=UpperCase(St2) ;
    //Transforme les lettres en chiffres selon le NEE 5.3
    i:=1 ;
    while i<Length(St2) do
    begin
      if St2[i] in ['A'..'Z'] then
      BEGIN
        ii:=Ord(St2[i])-65 ;
        st4:= copy(st2,1,i-1) + inttostr(10+ii) + copy(st2,i+1, length(st2));
        st2:=st4 ;
      END ;
      inc(i);
    end ;
    ret := '' ;
    cleL := 0 ;
    st4:='';
    //On découpe par tranche de 9
    //On calcul la clé via mod 97 puis on fait clé + reste du rib
    for i:=1 to (length(st2) div 9)+1 do
    begin
      st4 := copy(st2,1,9) ;
      delete(st2,1,9);
      strInter := inttostr(cleL)+st4 ;
      cleL := strtoint64(strinter) mod 97 ;
    end ;
    //une fois fini, on calcul 98-clé
    cleIBAN := inttostr(98-(cleL  mod 97));
    if length(cleIBAN)=1 then  cleIBAN := '0' + cleIBAN ;
    Result := Not (CleIBAN = CleIban2);
  end;

  {Provient de UtilRIB}
  function IsIBAN(IBAN : string) : boolean;
  begin
    try
      Result := not (ErreurDansIban(IBAN));
    except on e:
      EConvertError do Result:=false;
    end;
  end;
var
  Q:   TQuery;
  SQL: String;
  sType: String;
begin
  Q := OpenSQL('SELECT R_AUXILIAIRE, R_NUMERORIB, R_CODEIBAN FROM RIB '+
      ' WHERE (R_AUXILIAIRE<>"") AND NOT(R_AUXILIAIRE IS NULL)'+
      ' AND NOT (R_NUMERORIB IS NULL)', False);
  try
    while not Q.EOF do
      begin
      sType := '';
      if IsIBAN(Q.FindField('R_CODEIBAN').AsString) then
        sType := 'IBAN'
      else if Trim(Q.FindField('R_CODEIBAN').AsString) <> '' then
        sType := 'BBAN';
      SQL := 'UPDATE RIB SET R_TYPEIDBQ = "'+sType+'"' +
             ' WHERE R_AUXILIAIRE = "'+Q.FindField('R_AUXILIAIRE').AsString+'"' +
             '   AND R_NUMERORIB = '+Q.FindField('R_NUMERORIB').AsString;
      ExecuteSQL(SQL);
      Q.Next;
      end;
  finally
    Ferme(Q);
    end;
end;

procedure CPActivationPointageAvance ;
var
  SQL             : String ;
begin
  try
    BEGINTRANS ;

  //****************************************************************************
  //* Initalisation de la liste paramétrable du pointage manuel
  //****************************************************************************

    //* Comptabilité

    CreerListeParametrablePourPointageManuel('CPM') ;

    //* Trésorerie

    CreerListeParametrablePourPointageManuel('CPT') ;

    CommitTrans;
  except
    on E: Exception do
    begin
      if  (v_pgi.SAV) then
        BEGIN
        PgiError(E.message);
        END ;
      trace.traceError('ERROR','Exception (Initalisation de la liste paramétrable du pointage manuel)'+' '+e.message);

      RollBack;
    end;

  end;

  try
    BEGINTRANS ;

  //****************************************************************************
  //* Initialisation des souches bancaire sur l'exercice en cours et sur le suivant.
  //****************************************************************************

    LancerInitialiserEnSerieDesSouchesBancaires() ;

    CommitTrans;
  except
    on E: Exception do
    begin
      if  (v_pgi.SAV) then
        BEGIN
        PgiError(E.message);
        END ;
      trace.traceError('ERROR','Exception (Initialisation des souches bancaire sur l''exercice en cours et sur le suivant)'+' '+e.message );

      RollBack;
    end;

  end;

  // Si rappro déjà fait, pas de traitement à lancer :

  If GetparamSocSecur('SO_CP_NEWRAPPRO',FALSE) then Exit ;

  //****************************************************************************
  //* Mise à jour du détail des lignes
  //****************************************************************************

  try
    BEGINTRANS ;
  SQL := 'UPDATE EEXBQLIG SET CEL_DEVISE = (SELECT MAX(EE_DEVISE) FROM EEXBQ '
      +  'WHERE CEL_REFPOINTAGE = EE_REFPOINTAGE '
      +  'AND CEL_NUMRELEVE = EE_NUMERO) '
      +  'WHERE (CEL_DEVISE = "" OR CEL_DEVISE IS NULL)' ;
  ExecuteSQL(SQL);

  SQL := 'UPDATE EEXBQLIG SET CEL_REFPOINTAGE = "" '
      + ' WHERE (CEL_DATEPOINTAGE = "' + usDateTime_(iDate1900) + '" OR CEL_DATEPOINTAGE IS NULL)' ;
  ExecuteSQL(SQL);



  SetParamSoc('SO_CP_NEWRAPPRO','X') ;
    CommitTrans;

  except
    on E: Exception do
    begin
      if  (v_pgi.SAV) then
        BEGIN
        PgiError(E.message);
        END ;
      trace.traceError('ERROR','Exception (Bascule Rapprochement automatique)'+' '+e.message );

      RollBack;
      SetParamSoc('SO_CP_NEWRAPPRO','-') ;
    end;

  end;


end ;

procedure CPInitialisationEtablissementEtAgencesBancaires ;
begin

  //****************************************************************************
  //* Phase d'initialisation des banques, agences et comptes bancaires         *
  //****************************************************************************

  try
    BEGINTRANS ;

      InitialiserBanques() ;

    CommitTrans;
  except
    on E: Exception do
    begin
      if  (v_pgi.SAV) then
        BEGIN
        PgiError(E.message);
        END ;
      trace.traceError('ERROR','Exception (Initialisation des établissements et agences bancaires)'+' '+e.message );

      RollBack;
    end;

  end;
end ;

end.

