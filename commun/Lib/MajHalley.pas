unit MajHalley;

interface

uses ComCtrls, StdCtrls,BTMajStruct,ChangeVersions,uBTPVerrouilleDossier,galPatience;

function MAJHalleyAvant(VSoc: Integer; MajLab1, MajLab2: TLabel; MajJauge: TProgressBar): boolean;
function MAJHalleyApres(VSoc: Integer; MajLab1, MajLab2: TLabel; MajJauge: TProgressBar): boolean;
function MAJHalleyIsPossible(VSoc: Integer; MajLab1, MajLab2: TLabel; MajJauge: TProgressBar): boolean;
function MAJHalleyPendant(VSoc: Integer; NomTable: string): boolean;

procedure MAJStandardPaie;

//Ludovic MONTAVON Demande N° 1418
Procedure BasculeExpertEtConseil;

{ TS : 23/3/7 : Pour éviter de toujours faire un " if not IsDossierPCL then ... " dixit A.R. }
function ExecuteSQLNoPCL(const Sql: WideString): Integer;
function ExecuteSQLContOnExcept(const Sql: WideString): Integer;
procedure UpDateDecoupeLigneOuvPlat(SetSQL: string; SetCondition: string='');

//Pour Executer une Commande CL via SQL sous OS400 (DB2)
Function OS400_QCMDEXC( CommandeCL: String ): String;

//Marie-Christine DESSEIGNET Le 21/01/2008 Version 8.1.850.18 Demande N° 2069
Procedure ModifPlanningGA;

implementation

uses HEnt1, HCtrls, MajTable, Forms, DB,
{$IFDEF VER150} {D7}
  Variants,
{$ENDIF}
  {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
  MajStruc, HQry, SysUtils, Classes,
  DBCtrls, ParamSoc, Dialogs,
  MajHalleyUtil, MajHalleyProc,
  uHListe,Ent1, UtilGa, HmsgBox (*, STKMoulinette, yTarifsBascule, Wjetons,yTarifsCommun, GcLienPiece *),
  uTob , {utablesged,} wcommuns,majenrafale,{BasculeArticleLie,}YYBundle_tof,utilpgi ,galsystem, sql7util,
  cbptrace, UtilOracle,UtilCbp,DpJurOutilsBlob{,MajPontTax},Wjetons,UBTPSequenceCpta,uPGFnWorkRh,Bobgestion;

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

  stSavVerrou,stSavUser:string; // js1 120706 verrou
  // MCD
  AfplanCharge: boolean;

  //Marie-Christine DESSEIGNET Demande N° 1196
  afSurbooking: Boolean;



procedure UpDateDecoupeAna(SetSQL: string ; WhereSupJal : String = ''; WhereSupExo : String = '' ; WheresupEcr : String = '' ; DecoupeParExo : Boolean = TRUE);
var DMin, DMax, DD1, DD2: TDateTime;
  ListeJ: TStrings;
  Q, QExo: TQuery;
  i, iper, Delta, IperMax: integer;
  St  :String ;
begin
  // Lecture des journaux
  ListeJ := TStringList.Create;
  St:='Select J_JOURNAL from JOURNAL' ; If WhereSupJal<>'' Then st:=St+' WHERE '+WhereSupJal ;
  Q := OpenSQL(St, True);
  while not Q.EOF do
  begin
    ListeJ.Add(Q.Fields[0].AsString);
    Q.Next;
  end;
  Ferme(Q);
  // Balayage des écritures avec découpe
  for i := 0 to ListeJ.Count - 1 do
  begin
    St:='Select EX_EXERCICE, EX_DATEDEBUT,EX_DATEFIN from EXERCICE' ; If WhereSupExo<>'' Then st:=St+' WHERE '+WhereSupExo ;
    QExo := OpenSQl(St, True);
    while not QExo.EOF do
    begin
      DMin := QExo.Fields[1].AsDateTime;
      DMax := QExo.Fields[2].AsDateTime;
      Delta := Round((DMax - DMin) / 10);
      If DecoupeParExo then IperMax:=1 Else iPerMax:=10 ;
//      for iper := 1 to 10 do
      for iper := 1 to iPerMax do
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
        If DecoupeParExo then BEGIN DD1:=DMin ; DD2:=DMax ; END ;
        BeginTrans;
        St:='UPDATE ANALYTIQ SET ' + SetSQL + ' WHERE Y_JOURNAL="' + ListeJ[i] + '" AND Y_EXERCICE="' + QExo.Fields[0].AsString +
          '" AND Y_DATECOMPTABLE>="' + UsDateTime_(DD1) + '" AND Y_DATECOMPTABLE<="' + UsDateTime_(DD2) + '" ' ;
        If WhereSupEcr<>'' Then st:=St+' AND ('+WhereSupEcr+') ' ;
        ExecuteSQLContOnExcept(St);
        CommitTrans;
      end;
      QExo.Next;
    end;
    Ferme(QExo);
  end;
  ListeJ.Clear;
  ListeJ.Free;
end;

//Beatrice MERIAUX Demande 2971 ANNULE ET REMPLACE LA PRECEDENTE PROCEDURE (//Béatrice Mériaux Demande N° 2441)
// Transformation du stockage des conventions :
//- actuellement découpées en 3 parties
//- regroupement des différentes zones et stockage sous forme de blob dans la table LIENSOLE
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 13/05/2008
Modifié le ... :   /  /
Description .. : Transformation de l'affichage des conventions
Mots clefs ... :
*****************************************************************}
procedure MajConventions;
var
   FPatience_l : TFPatience;
   OBConventions_l, OBConvLibres_l, OBAnnulien_c : TOB;
   bMajOK_l : boolean;
   iAnlInd_l : integer;
   sGuidConv_l, sEmploiBlob_l, sPrefixe_l, sRangBlob_l, sLibelle_l, sConvention_l, sConvLibre_l : string;
   sAnlConvTxt_l, sAnlRefLien_l, sAnlNoteLien_l : string;
   bAnlConv_l, bAnlConvLib_l, bAnlConvSuite_l : boolean;
begin
   bMajOK_l := GetParamSocSecur('SO_JUCONVOK', false);
   if not bMajOK_l then exit;

   FPatience_l := FenetrePatience('Mise à jour de la gestion des conventions en cours, veuillez patientez...');
   FPatience_l.lCreation.Caption := 'Exécution de la procédure';
   FPatience_l.lAide.Caption     := 'sur la base commune ' + V_PGI.DbName;
   FPatience_l.Invalidate;
   Application.ProcessMessages;

   OBAnnulien_c := TOB.Create('ANNULIEN', nil, -1);
   OBAnnulien_c.LoadDetailDBFromSQL('ANNULIEN', 'SELECT * FROM ANNULIEN');

   for iAnlInd_l := 0 to OBAnnulien_c.Detail.Count - 1 do
   begin
      if OBAnnulien_c.Detail[iAnlInd_l].GetString('ANL_GUIDCONV') <> '' then Continue;

      sConvention_l := '';
      sConvLibre_l := '';

      bAnlConv_l      := OBAnnulien_c.Detail[iAnlInd_l].GetString('ANL_CONV') = 'X';
      bAnlConvLib_l   := OBAnnulien_c.Detail[iAnlInd_l].GetString('ANL_CONVLIB') = 'X';
      bAnlConvSuite_l := OBAnnulien_c.Detail[iAnlInd_l].GetString('ANL_CONVSUITE') = 'X';

      sAnlConvTxt_l := OBAnnulien_c.Detail[iAnlInd_l].GetString('ANL_CONVTXT');
      sAnlRefLien_l := OBAnnulien_c.Detail[iAnlInd_l].GetString('ANL_REFLIEN');
      sAnlNoteLien_l := OBAnnulien_c.Detail[iAnlInd_l].GetString('ANL_NOTELIEN');
      sAnlNoteLien_l := GetRTFStringText(sAnlNoteLien_l);
      if sAnlNoteLien_l = #0 then
         sAnlNoteLien_l := '';

      sConvention_l := sAnlConvTxt_l;

      if sAnlRefLien_l <> '' then
      begin
         if not bAnlConvLib_l then
            sConvention_l := sConvention_l + #$D#$A + sAnlRefLien_l
         else
            sConvLibre_l := sAnlRefLien_l;
         sAnlRefLien_l := '';
      end;

      if bAnlConvSuite_l and (sAnlNoteLien_l <> '') then
      begin
         sConvention_l := sConvention_l + #$D#$A + sAnlNoteLien_l;
            sAnlNoteLien_l := '';
      end;

      sGuidConv_l := AglGetGuid;

      // Convention
      BlobGetCode('OLE_CONVENTION', sEmploiBlob_l, sPrefixe_l, sRangBlob_l, sLibelle_l);
      OBConventions_l := BlobCreation(sPrefixe_l, sGuidConv_l, sRangBlob_l, sEmploiBlob_l,
                                       sLibelle_l, sConvention_l);
      OBConventions_l.InsertDB(nil);
      FreeAndNil(OBConventions_l);

      // Convention libre
      BlobGetCode('OLE_CONVLIBRE', sEmploiBlob_l, sPrefixe_l, sRangBlob_l, sLibelle_l);
      OBConvLibres_l := BlobCreation(sPrefixe_l, sGuidConv_l, sRangBlob_l, sEmploiBlob_l,
                                    sLibelle_l, sConvLibre_l);
      OBConvLibres_l.InsertDB(nil);
      FreeAndNil(OBConvLibres_l);

      OBAnnulien_c.Detail[iAnlInd_l].PutValue('ANL_GUIDCONV', sGuidConv_l);
      OBAnnulien_c.Detail[iAnlInd_l].UpdateDB();

   end;
   OBAnnulien_c.Free;
   if FPatience_l <> Nil then
   begin
      FPatience_l.Close;
      FPatience_l.Free;
   end;
   SetParamSoc('SO_JUCONVOK', false);
end;

  {==============================================================================}
  {=========================== PROCEDURES MAJ AVANT =============================}
  {==============================================================================}

Procedure MajAvant900;
Begin

  //Marie-Noëlle GARNIER Demande N° 1765
  //Demande déjà intégrée en MajAvant850 mais seulement dans la version 810 du PGIMajVer
  //c'est à dire pour la version Business qui est sortie décalée de la version v800 de PCL.
  //Donc en Entreprise les clients auront bien eût le DELETE puis la BOB ramenant la nouvelle
  //liste, par contre en PCL il faut bien la détruire avant que la nouvelle SOCREF v900 import
  //cette nouvelle liste.
  If V_PGI.ModePCL='1' then
    ExecuteSQLContOnExcept( 'DELETE FROM LISTE WHERE LI_LISTE="RTMULACTIONACTGEN"' );

  //Jean-Luc SAUZET Le 06/03/2008 Version 9.0.900.22 Demande N° 2217
  (*
  ExecuteSQLContOnExcept( 'DELETE FROM LISTE WHERE LI_LISTE LIKE "YTARIFSMUL%"' );
  ExecuteSQLContOnExcept( 'DELETE FROM LISTE WHERE LI_LISTE LIKE "YTARIFSFSL%"' );
  *)
  //Stéphane BOUSSERT Le 19/03/2008 Version 9.0.900.22 Demande N° 2231
  ExecuteSQLContOnExcept( 'DELETE FROM LISTE WHERE LI_LISTE IN ("CPEEXBQ", "CPEEXBQ2", "CPMULEEXBQLIG", "CPMULCETEBAC")' );

  //Stéphane BOUSSERT Le 19/03/2008 Version 9.0.900.22 Demande N° 2254
  ExecuteSQLContOnExcept( 'DELETE FROM LISTE WHERE LI_LISTE in ("CPVALIDEECR", "CPMULCETEBAC", "CPLISTETRACFONB")' );
  //
End;

Procedure MajAvant902;
Begin

  //Jean-Luc SAUZET Le 02/04/2008 Version 9.0.902.25 Demande N° 2311
  (*
  ExecuteSQLContOnExcept( 'DELETE FROM DECOMBOS WHERE DO_COMBO = "WPREFIXEREV"' );
  ExecuteSQLContOnExcept( 'DELETE FROM LISTE WHERE LI_LISTE="WJETONS"' );
  *)
  ExecuteSQLContOnExcept( 'DELETE FROM LISTE WHERE LI_LISTE="RQDEMDEROG"' );

End;

Procedure MajAvant903;
Begin

  //Nicolas FOURNEL Le 09/04/2008 Version 9.0.903.26 Demande N° 2322
  ExecuteSQLContOnExcept( 'DELETE FROM LISTE WHERE LI_LISTE IN ("PSE_MULFLUX","PSE_MULIMPORTS","PSE_MULEXPORTS", "GCNOMENLIG")' );

  //Nicolas FOURNEL Le 09/04/2008 Version 9.0.903.26 Demande N° 2331
  //Je Fusionne avec la ligne du dessus
  //ExecuteSQLContOnExcept( 'DELETE FROM LISTE WHERE LI_LISTE IN ("GCNOMENLIG")' );

End;

Procedure MajAvant904;
Begin

  //Nicolas FOURNEL Le 16/04/2008 Version 9.0.904.32 Demande N° 2364
  ExecuteSQLContOnExcept( 'DELETE FROM LISTE WHERE LI_LISTE IN ("PSE_MULFLUX","PSE_MULIMPORTS","PSE_MULEXPORTS")' );

  //Nicolas FOURNEL Le 16/04/2008 Version 9.0.904.32 Demande N° 2366
  //Je Fusionne avec la ligne du dessus
  //ExecuteSQLContOnExcept( 'DELETE FROM LISTE WHERE LI_LISTE IN ("GCGROUPEPIECEACH")');

End;

Procedure MajAvant905;
Begin

//Jean-Luc Sauzet Le 23/04/2008 Version 9.0.905.36 Demande N° 2376
//  ExecuteSQLContOnExcept( 'DELETE FROM LISTE WHERE LI_LISTE="GCSTKPICKING" ');

End;

Procedure MajAvant910;
Begin
  //M Faudel Le 04/06/2008 Version 9.0.911.16 Demande N° 2534
  if TableExiste('DUEMTRAVAIL') then ExecuteSQLContOnExcept('DELETE FROM DUEMTRAVAIL');
  ExecuteSQLContOnExcept('DELETE FROM COMMUN WHERE CO_TYPE = "DU2"');

  AglNettoieListes('PGDUELISTES','PUS_NODOSSIER',nil);
  AglNettoieListes('PGHISTODETAIL',' PHD_DATEAPPLIC;PHD_ETABLISSEMENT;PHD_PGTYPEHISTO;PHD_PGTYPEINFOLS',nil);

  //M Faudel Le 04/06/2008 Version 9.0.911.16 Demande N° 2536
  ExecuteSQLContOnExcept('delete from DECOMBOs where DO_COMBO="PGCOMURBAINE"');
  ExecuteSQLContOnExcept('delete from DECOMBOs where DO_COMBO="PGcodeCATEGORIE2"');

  //M DESGOUTTE Le 04/06/2008 Version 9.0.911.16 Demande N° 2493
  if TableExiste('YFORMESINSEE') then ExecuteSQLContOnExcept('DELETE FROM YFORMESINSEE');

End;

Procedure MajAvant911;
Begin
  //N CHAVANNE Le 04/06/2008 Version 9.0.911.16 Demande N° 2625
//  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_CONTROLE = (DH_CONTROLE || "1")  WHERE DH_NOMCHAMP = "GA2_NUMEROSERIEGR" AND DH_CONTROLE NOT LIKE "%1%" ');

End;

Procedure MajAvant912;
Begin
  //M MORRETTON Le 11/06/2008 Version 9.0.912.3 Demande N° 2662
//  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE LIKE "YTARIFS%601"');
End;

Procedure MajAvant913;
Begin
End;

Procedure MajAvant917;
Begin
  //M FAUDEL Le 23/07/2008 Version 9.0.917.1 Demande N° 2930
  ExecuteSQLContOnExcept ('delete from modeles where mo_Nature="PAY" and mo_code="PGD" and mo_langue="FRF"');
  ExecuteSQLContOnExcept ('delete from modeles where mo_Nature="PSA" and mo_code="PHS" and mo_langue="FRF"');
  //S BOUSSERT Le 23/07/2008 Version 9.0.917.1 Demande N° 2932
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE ="CPEEXBQ"');

End;

Procedure MajAvant918;
Begin
  //M VERMOT GAUCHY Le 28/07/2008 Version 9.0.918.1 Demande N°2999
  if TableExiste('YPARAMEDITION') then
    ExecuteSQLContOnExcept('delete from yparamedition where yed_predefini="CEG" and yed_tproduit="IMM"');
End;

Procedure MajAvant919;
Begin
  //S BOUSSERT le 27/08/2008 Dem n°3031
  // Suppression des listes CPRLANCETRA et CPRELANCEDIV
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE="CPRLANCETRA"');
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE="CPRELANCEDIV"');
//  //M MORRETTON le 27/08/2008 Dem n°3067
//  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE LIKE "YTARIFSMAJ%"');
end;

Procedure MajAvant920;
Begin
  //D SCLAVOPOULOS Dem n°3136
  ExecuteSQLContOnExcept('DELETE FROM MODELES WHERE MO_TYPE="E" AND MO_NATURE="WO4" AND MO_CODE="OL0"');
End;

Procedure MajAvant921;
Begin
  //S BOUSSERT 3195
  // Ajout du champs E_MODESAISIE dans la liste CPRECHERCHEECR
  AglNettoieListes('CPRECHERCHEECR', 'E_MODESAISIE', nil);

  //D SCLAVOPOULOS 3206
  ExecuteSQLContOnExcept('DELETE FROM MODEDATA WHERE MD_CLE LIKE "EWO4OL0%"');

end;

Procedure MajAvant922;
Begin
  //JLS 3226
  AGLNettoieListesPlus('YYMULJNALEVENT','',nil,False,'GEV_BLOCNOTE');
//  AGLNettoieListes('GCLISTINVLIGNE','GIL_VALLIBRE1;GIL_VALLIBRE2;GIL_VALLIBRE3',nil);

  //M FAUDEL 3256
  // Liste PGVENTILAREM champ PVS_NODOSSIER obligatoire
  AglNettoieListes('PGVENTILAREM', 'PVS_NODOSSIER',nil);

  // Liste PGVENTILAORG champ PVO_NODOSSIER obligatoire
  AglNettoieListes('PGVENTILAORG ', 'PVO_NODOSSIER',nil);

  // Liste PGVENTILAORGAN champ PVO_NODOSSIER obligatoire
  AglNettoieListes('PGVENTILAORGAN', 'PVO_NODOSSIER',nil);

  //M FAUDEL
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE IN ("PGAFFECTDUCS","PGAFFECTDUCSDB0")');

End;

Procedure MajAvant923;
Begin
  //D SCLAVOPOULOS 3275
//  AGLNettoieListes( 'UORDRELIG', 'WOL_WBMEMO');

  //MC DESSEIGNET 3280
  //ces listes ne sont pas personnalisable. on peut les detruire pour les changer
//  executesqlNopcl ('delete FROM LISTE WHERE LI_LISTE like "AFSAISIEACTRES%"');

  //D SCLAVOPOULOS 3302
  (*
  AglNettoieListes('WGAMMECIR','WGC_GUIDWGL', nil,'WGC_IDENTIFIANTWGL');
  AglNettoieListes('WGAMMELIG;WGAMMELIG1;WGAMMELIGMODIFLOT','WGL_GUID',nil,'WGL_IDENTIFIANT');
  AglNettoieListes('WGAMMERESMODIFLOT','WGR_GUID',nil,'WGR_IDENTIFIANT');
  AglNettoieListes('WNOMECPD','WNL_GUID', nil,'WNL_IDENTIFIANT');
  AglNettoieListes('WNOMELIG','WNL_GUID', nil,'WNL_IDENTIFIANT');
  AglNettoieListes('WNOMEDEC','WND_GUIDWNL',nil,'WND_IDENTIFIANTWNL');
  AglNettoieListes('WNOMELIGDEM','WNL_GUID',nil,'WNL_IDENTIFIANT');
  AglNettoieListes('WNOMETET;WNOMETETMODIFLOT;WREVISIONWNT;WDELETEWNT','WNT_GUID',nil,'WNT_IDENTIFIANT');
  AglNettoieListes('WOB_RUPTURES','WOB_GUID', nil,'WOB_IDENTIFIANT');
  AglNettoieListes('WORDREBES2CPD;WORDREBES;WORDREBES2;WORDREBESCPD;WORDREBESDEM;WORDREBESL','WOB_GUID',nil,'WOB_IDENTIFIANT');
  AglNettoieListes('WORDRECMP','GL_GUIDWOL', nil,'GL_IDENTIFIANTWOL');
  AglNettoieListes('WORDREGAMME1;WORDREGAMME','WOG_GUID',nil,'WOG_IDENTIFIANT');
  AglNettoieListes('WORDRELIG;WORDRELIG2','WOL_GUID',nil,'WOL_IDENTIFIANT');
  AglNettoieListes('WORDREPHASE;WORDREPHASE2','WOP_GUID',nil,'WOP_IDENTIFIANT');
  AglNettoieListes('WORDRETET','WOT_GUID',nil,'WOT_IDENTIFIANT');
  AglNettoieListes('WREVISION','WJA_GUID;WJA_GUIDORI',nil,'WJA_IDENTIFIANT');
  AglNettoieListes('WGAMMETET;WGAMMETET2;WGAMMETETMODIFLOT;WREVISIONWGT;WDELETEWGT','WGT_GUID',nil,'WGT_IDENTIFIANT');
  AglNettoieListes('WRECODIFARTICLE','WJA_GUID;WJA_GUIDACTION;WJA_GUIDORI', nil,'WJA_IDENTIFIANT;WJA_IDACTION;WJA_IDENTIFIANTWXX');
  AglNettoieListes('UORDREBES','WOB_GUID', nil, 'WOB_IDENTIFIANT');
  AglNettoieListes('UORDRELIG','WOL_GUID');
  AglNettoieListes('WJOURNALACTION','WJA_GUID;WJA_GUIDORI;WJA_GUIDACTION');
  *)

End;

Procedure MajAvant924;
Begin
End;

Procedure MajAvant925;
Begin
  //MC DESSEIGNET 3427
  //liste AfSaisi* pas modifiable. obligation de détruire les uaters car ajout table
//  ExecuteSQLContOnExcept('delete from liste where li_liste in ("afsaisieActRestts","afsaisieActReshts",'
//  +'"afsaisieActReshtt","afsaisieActResttt","afsaisieActaffht", "afsaisieActafftt","afdptiersmailing","afmulcontmailing","afaffmailingcon")');

  // R BOENINGEN 3357
//  AGLNettoieListes('GCMULLIGNE','GL_DATEPIECE',nil);

  //C DUMAS 3388
(*
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE IN ("RTMULTIERSMAILING","RTMULCONTMAILING","RTAFFMAILINGCON"'
                   +',"RTMULLIGNMAILING","RTTIERSMAILINGFIC","RTCONTMAILINGFIC","RTAFFMAILINGFIC","RTLIGNMAILINGFIC"'
                   +',"RFMULTIERSMAILING","RFMULLIGNMAILING","RFTIERSMAILINGFIC","RFLIGNMAILINGFIC")');
*)
End;

Procedure MajAvant926;
Begin
{  // M VERMOT GAUCHY 3441
  ExecuteSQLContOnExcept('delete from yparamedition where yed_predefini="CEG" and yed_tproduit="IMM"');
  Pour l'instant on ne fait pas car problème uniquement observé en dev}
  // M MORETTON 3446
//  ExecuteSQLNoPCL('DELETE FROM PARAMSOC WHERE SOC_NOM IN ("SCO_PDRDEFAUTGA","SCO_PDRMETHVALO","SCO_PDRGENEREWPL","SCO_PDRCONSOAVANCE")');
  // JLS 3451
  // JLS 09/10/2008: La liste a changé de référenciel table. Je suis obligé de la supprimer car le AGLNettoieListe ne permet pas celà
//  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE="GCSTKDISPODETAIL"');
  // S BOUSSERT 3481
  // MAJ du champ G_CUTOFFCOMPTE de la table GENERAUX suite mauvaise initialisation
  ExecuteSQLContOnExcept('UPDATE GENERAUX SET G_CUTOFFCOMPTE="" WHERE G_CUTOFFCOMPTE="-"') ;
End;

Procedure MajAvant927;
Begin
//ATTENTION INTEGRATION EN 927 DES ELEMENTS DE STRCUTURE KPMG MULTI ENTITES
  //N FOURNEL 3479
//  AglNettoieListes('GCGROUPEMANVTE;GCGROUPEMANACH','GP_DOMAINE',nil);
End;

Procedure MajAvant928;
Begin
  //MCD 3522
//  ExecuteSqlContOnExcept('delete from liste where li_liste in ("AFTACHEGEN","AFMULPLANNINGGENE")');
  //N FOURNEL 3535
  ExecuteSQLContOnExcept('DELETE FROM COMMUN WHERE CO_TYPE = "GP2"');
  // MD3 3528 
  // Nettoyage des listes suite au remplacement de DOR_NONTRAITE par DOS_NONTRAITE
  AGLNettoieListes ('DPDOSSIEROBLIG','',nil,'DOR_NONTRAITE');
  AGLNettoieListes ('YYMULANNDOSS','',nil,'DOR_NONTRAITE');
  AGLNettoieListes ('YYMULANNDOSSLITE','',nil,'DOR_NONTRAITE');
  AGLNettoieListes ('YYMULSELDOSS','',nil,'DOR_NONTRAITE');
  AGLNettoieListes ('YYMULSELDOSSLITE','',nil,'DOR_NONTRAITE');
  // DBR 3554
//  AglNettoieListes ('GCDUPLICPIECE', 'GP_DATEPIECE', nil);

End;

Procedure MajAvant929;
Begin
  //MCD 3598
  // pas OK en 928 à refaire en 929
//  ExecuteSqlContOnExcept('delete from liste where li_liste in ("aftachegen","afmulplanninggene")');
End;

Procedure MajAvant930;
Begin
  //M DESGOUTTE 3610
  // Suppression de paramsoc obsolètes
  ExecuteSQLContOnExcept('delete from paramsoc where soc_nom like "%DIODE%" and SOC_TREE like "001;014;003;%"');
End;

Procedure MajAvant931;
Begin

End;

Procedure MajAvant932;
Begin
//  if not IsDossierPCL() then
//    EdiMajVer(932).MajAvant();
  //D SCLAVOPOULOS 3693
//  AglNettoieListes('WCHANGEMENTWOL','WCH_GUID;WCH_GUIDORI',nil,'WCH_IDENTIFIANT');
  //T PETETIN 3710
//  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE="WUNITE"');

End;

Procedure MajAvant933;
Begin
  //M MORRETTON 3770
  (*
  AGLNettoieListesPlus('YTARIFSMULTIE101', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULTIE201', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULTIE211', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULTIE301', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULTIE401', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULTIE501', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULTIE601', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULART101', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULART201', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULART211', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULART301', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULART401', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULART501', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULART601', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULMAJ101', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULMAJ201', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULMAJ211', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULMAJ301', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULMAJ401', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULMAJ501', 'YTS_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULMAJ601', 'YTS_PROFILCOMPO', nil, False);
  *)
End;

Procedure MajAvant934;
Begin
  //DBR 3824
  (*
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE ="WPARCNOME_TRT"');
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE ="WPARCELEM"');
  *)
End;

Procedure MajAvant935;
Begin
  //D SCLAVOPOULOS 3900
  AglNettoieListes('RQNONCONF', 'RQN_AFFAIRE', nil);

End;

Procedure MajAvant936;
Begin
  //JP LAURENT 3908
  AGLNettoieListes('QULBPARBRE004;QULBPARBRECA1;QULBPARBRECA2;QULBPARBRECA3;QULBPARBRECA4 ;QULBPARBRECA5; QULBPARBRECA6 ; QULBPARBREQTE','QBR_NIVALAFF');
  //MCD 3921
//  AglNettoieListesPlus('AFALIGNAFF', 'AFF_TIERS', nil, true);
  //D SCLAVOPOULOS 3922
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE="WORDRECMP"');
  //D SCLAVOPOULOS 3925
//  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE="GCSTKDISPATCH"');
  //D KOZA 3941
  ExecuteSqlNoPCL('DELETE FROM QMACROGAM');
  ExecuteSqlNoPCL('DELETE FROM QMAGENC');
  ExecuteSqlNoPCL('DELETE FROM QNOMENCLA WHERE QNO_CTX="0"');
  //D SCLAVOPOULOS 3956
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE = "WORDREBES3"');
  //D SCLAVOPOULOS 3963
  ExecuteSQLContOnExcept( 'DELETE FROM LISTE WHERE LI_LISTE in ("RQDEMDEROG", "RQNONCONF", "RQPLANCORR")' );

End;

Procedure MajAvant937;
Begin
  //T SUBLET 4011
  { Suppression de la liste EDICADENCE }
//  ExecuteSqlNoPCL('DELETE FROM LISTE WHERE LI_LISTE="EDICADENCE"');

  { Suppression de la liste EDICALAGE }
//  ExecuteSqlNoPCL('DELETE FROM LISTE WHERE LI_LISTE="EDICALAGE"');

  { Suppression de la liste EDIBLREC }
//  ExecuteSqlNoPCL('DELETE FROM LISTE WHERE LI_LISTE="EDIBLREC"');

  //N FOURNEL 4038
//  AglNettoieListes('GCGROUPEPIECEACH', 'GP_DOMAINE',nil);
End;

Procedure MajAvant938;
Begin
  //D SCLAVOPOULOS 4060
//  ExecuteSqlContOnExcept('DELETE FROM LISTE WHERE LI_LISTE IN ("GCLISTINVLIGNECTM","GCLISTINVLIGNE")');
  // G JUGDE4082
  AGLNettoieListes ('QMESHISTORES;QMESHISTORESO','QWH_DEPOT;QWH_DATERELEVE;QWH_RESSOURCE;QWH_SITE;QWH_RELEVE;QWH_NATURETRAVAIL;QWH_LIGNEORDREGEN;QWH_LIGNEORDRE;QWH_BAC;QWH_ALEA;QWH_CODEARTICLE;QWH_OPECIRC;QWH_NUMOPERGAMME;QWH_LIGNERELEVE;QWH_SOUSLIGNE;QWH_PHASE;QWH_TYPELIGNEPDR;QWH_POSTE',nil);
  //M MORRETTON 4086
  (*
  AGLNettoieListesPlus('YTARIFSPARAMETRES', 'YFO_PROFILCOMPO', nil, False);
  AGLNettoieListesPlus('YTARIFSMULTIE101', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULTIE201', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULTIE211', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULTIE301', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULTIE401', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULTIE501', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULTIE601', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULART101', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULART201', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULART211', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULART301', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULART401', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULART501', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULART601', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULMAJ101', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULMAJ201', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULMAJ211', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULMAJ301', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULMAJ401', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULMAJ501', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  AGLNettoieListesPlus('YTARIFSMULMAJ601', 'YTS_APPORTEUR;YTS_PRESCRIPTEUR;YTS_CATALOGUE', nil, False);
  *)
End;

Procedure MajAvant939;
Begin
  // D SCLAVOPOULOS 4119
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE IN ("UORDREBESLASSAIS","UORDREBESLASSTOC")');

  //S BOUSSERT 4107
  // MAJ de la liste FEMPRUNT par ajout d'un champ obligatoire
  AglNettoieListes('FEMPRUNT', 'EMP_NOCALCREGUL', nil ) ;

End;

Procedure MajAvant940;
Begin
  //MCD 4184
//  AglNettoieListesPlus('AFMULMODIFLOTAFFA', 'RDW_CLEDATA', nil, true);
End;

Procedure MajAvant941;
Begin
  //MCD 4198
  //liste AfSaisi* pas modifiable. obligation de les détruire dans les bases car ajout champs
//  ExecuteSQLContOnExcept( 'DELETE FROM LISTE WHERE li_liste LIKE "afsaisieact%"' );

  //G JUGDE 4212
  SUPPRIMEETAT ('E','QGQ','Q03');
End;

Procedure MajAvant942;
Begin
End;

Procedure MajAvant943;
Begin
  //D SCLAVOPOULOS 4309
//  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE ="WORDREBES3"');
  //MCD 4350
  //les listes suivantes sont à détruire
//  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE IN ("AFPLANPARAMMUL","AFLISTETACHES","AFPLALIGNE_MUL")');

//  AGLNettoieListesPlus( 'RTLTIERSMULCON' ,'T_AUXILIAIRE', Nil, True );
//  AGLNettoieListesPlus( 'RTLTIERSMULCON2' ,'T_AUXILIAIRE', Nil, True );

End;

Procedure MajAvant944;
Begin

  //MNG 4378
  { ajout 3 champs obligatoires dans la liste }
//  AglNettoieListes('RTANNUAIRE', 'T_PARTICULIER;T_TIERS;C_FERME',nil);

  //G JUDGE 4417
//  AglNettoieListesPlus( 'WGAMMELIG1', '', nil, False, 'WGL_CIRCUIT' );
//  AglNettoieListesPlus( 'WGAMMELIG1', 'WGL_CIRCUIT;WGL_SITE;WGL_DEPOT;WGL_TIERS;WGL_SOCIETEGROUPE', nil, False, '' );

End;

Procedure MajAvant947;
Begin
  //MNG 4529
  { ajout 4 chps oblig de type formule }
//  ExecuteSQLContOnExcept( 'DELETE FROM LISTE WHERE LI_LISTE="RTTIERS_PILOTACTC"' );

End;

Procedure MajAvant948;
Begin
  //MCD 4647
//  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE ="AFPLALIGNE_MUL"');
  //MCD 4753
//  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE ="AFLISTETACHES"');
End;

Procedure MajAvant951;
Begin
  //MNG 4824
  // liste repassée en socref
//  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE ="RTAM_RESMARKET_M"');
  //G JUGDE 4862
//  ExecuteSQLContOnExcept( 'DELETE FROM LISTE WHERE LI_LISTE IN ("WGAMMERES","WORDRERES","WGAMLIGEXCEPT","QWHISTORES")' );
End;

Procedure MajAvant952;
Begin
  // D SCLAVOPOULOS 4927
//    ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE ="WORDREPHASE2"');
End;

Procedure MajAvant953;
Begin
  // D KOZA 4977
  ExecuteSqlNoPCL('DELETE FROM LISTE WHERE LI_LISTE IN ("QULBTNPROOPE002","QULBTNPROPHA002","QULCHXPRODTR004","QULDETCIRCPH002","QULDETSCENA003","QULGROUPE002"'
              + ',"QULGRPTYPE15001","QULGRPTYPE2001","QULGRPTYPE3001","QULINFOBAR005","QULLOGACTIONS001","QULLOGANOMALIE","QULPHASECAR004","QULPLANGP8","QULPLANPROD"'
              + ',"QULPROFILCA001","QULPROINDUS1","QULREGRPMENU002","QULREGRPPLAN003","QULREPARTCIRC001","QULREPARTFABCIRC","QULSATTRIB003","QULSATTRIBBIS004"'
              + ',"QULSCENARCA002","QULSELCHXRAL003","QULSELECTOF006","QULSIMULATION002","QULVALIDSIMU002","QULVARIANTECAD001","QULPARAMTYPN","QULTRANDETD2","QULFILTREBESOIN","QULITI0")');

End;

Procedure MajAvant954;
Begin
  // MCD 5004
  //GIGA liste à recréer suite ajout phase
//  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE ="AFLISTETACHES"');
End;

Procedure MajAvant955;
Begin
  //T SUBLET 5054
  { suppression des listes de la prépa. d'expé. }
//  ExecuteSQLNoPCL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCPREPAEXPAFC", "GCPREPAEXPAFF", "GCPREPAEXPEXT")');
End;

Procedure MajAvant956;
Begin
  // Gm 5100
//  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE ="AFLISTETACHES"');
  //LEK 5106
  SupprimeEtat('E','BAB','BA2');
  SupprimeEtat('E','BAB','BA3');
  SupprimeEtat('E','BAB','BA4');
  SupprimeEtat('E','BAB','BA5');
  SupprimeEtat('E','BAB','BA6');
  SupprimeEtat('E','CBR','CBG');
  SupprimeEtat('E','CBX','CA2');
  SupprimeEtat('E','CGL','CTV');
  SupprimeEtat('E','CGL','CVP');
  SupprimeEtat('E','CJ6','CJ0');
  SupprimeEtat('E','CJ6','CJ1');
  SupprimeEtat('E','CJ6','CJ2');
  SupprimeEtat('E','CJ6','CJ3');
  SupprimeEtat('E','CJ6','CJ4');
  SupprimeEtat('E','CJ6','CJ6');
  SupprimeEtat('E','CJ6','CJ7');
  SupprimeEtat('E','JAL','JCD');
  SupprimeEtat('E','JAL','JCJ');
  SupprimeEtat('E','JAL','JDD');
  SupprimeEtat('E','JAL','JDE');
  SupprimeEtat('E','JAL','JDL');
  SupprimeEtat('E','JAL','JDP');
  SupprimeEtat('E','JAL','JDV');
  SupprimeEtat('E','JAL','JPJ');
  SupprimeEtat('E','JAL','JPN');
  SupprimeEtat('E','JAL','JPP');

End;

Procedure MajAvant957;
Begin
  //C PARWEZ 5110
//  ExecuteSQLNoPCL('SELECT * FROM PARAMSOC WHERE SOC_TREE  LIKE "001;020;003;%" AND (SOC_NOM LIKE "%WXC%" OR SOC_NOM LIKE "%CFG%")');
End;

Procedure MajAvant958;
Begin
End;

Procedure MajAvant959;
Begin
  //CD 5237
//  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE ="GCCORRESIMPORT"');
End;

Procedure MajAvant960;
Begin
{  // MNG 5234   / 5270
  // vidage table qui contient trop de mauvaises traductions
  ExecuteSQLContOnExcept('DELETE FROM YTRADMETIER');}

End;

Procedure MajAvant961;
Begin
  //MNG 5321
  // 1-changement de vue 2-suppression jointure sur RTINFOS006 qui est mise dans la vue
//  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE IN ("RTAM_RESMARKET_M","RTMULQUALITECONTA")');
  //JPL 5346
//  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE in ("GCMULPIECEDATR","GCMULPIECEDAVAL")');
  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE in ("QBPAXEARBRE","QULBPPNEWSAL","QULBPPBOOL","QBPINDEMBDEF")');
  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE in ("QULBPARBREPAIE","QULBPARBREPAIE1","QULBPARBREPAIE2","QULBPARBREPAIE3","QULBPARBREPAIE4","QULBPARBREPAIE5","QULBPARBREPAIE6","QULBPARBREPAIE7")');
  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE in ("QULBPARBRE","QULBPARBRE002","QULBPARBRE003","QULBPARBRE004","QULBPARBRECA1","QULBPARBRECA2","QULBPARBRECA3","QULBPARBRECA4")');
  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE in ("QULBPARBRECA5","QULBPARBRECA6","QULBPARBREQTE","YYMULTIDOSSIER")');
  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE in ("QUVBPARBRE1","QUVBPARBRE2","QUVBPARBRE3","QUVBPARBRE4","QUVBPARBRE5","QUVBPARBRE6","QUVBPARBRE7")');
  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE in ("QBPACCESAXE","QBPSEQBUDGET","QULBPCHXENSEIGN","QULBPCHXGROUP")');
End;

Procedure MajAvant962;
Begin
  //N FOURNEL 5409
  ExecuteSQLContOnExcept( 'DELETE FROM LISTE WHERE LI_LISTE="PSE_MULCATALOGUE" ');
End;

Procedure MajAvant963;
Begin
End;

Procedure MajAvant964;
Begin
End;

Procedure MajAvant965;
Begin
End;

Procedure MajAvant966;
Begin
  //JPL 5649
  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE in ("QBPSEQBUDGET","QBPACCESAXE")');

End;

Procedure MajAvant967;
Begin
  //MNG 5664
  // CRM : pour Maj jointure sur contact
//  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE IN ("RTPERSPHISTO","RTVISACONTACT")');
End;

Procedure MajAvant968;
Begin
  //MNG 5748
  // mng : ajout du champ siret obligatoire
//  AglNettoieListesPlus('RTMULSUSPECT', 'RSU_SIRET',nil,true);
  //MMORRETTON 5782
//  ExecuteSQLNoPCL('DELETE FROM CHOIXEXT WHERE YX_TYPE="WTR"');
End;

Procedure MajAvant971;
Begin
  //LEK 6023
  SupprimeEtat('E','TVA','CR2');
  SupprimeEtat('E','TVA','EAC');
  SupprimeEtat('E','TVA','SUI');
  SupprimeEtat('E','TVA','CTR');
End;

Procedure MajAvant972;
Begin
  //C DUMAS 6101
  ExecuteSQLContOnExcept ('delete from paramsoc where soc_tree = "001;003;023;034;" or soc_tree = "001;003;023;036;" or soc_tree = "001;003;023;038;" or soc_tree = "001;003;023;039;" or soc_tree = "001;003;023;049;" or soc_tree = "001;003;023;050;"');

  //LEK CHHOEU 6128
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE IN ("SUPRANA","CPMULCUTOFF","CPVALIDEBAP","CPPRORATAPARAMLST","CPMULCFONB","CPMULCETEBAC",'
  +'"CPLISTVISEUR","CPLISTETYPEVISA","CPLISTESUIVIBAP","CPLISTERELANCEBAP","CPLISTEPURGEBAP","CPLISTEMODIFBAP","CPLISTECIRCUITBAP","CPHISTOBAPVISEUR",'
  +'"CPGENEANAL2","CPAFFICHEALLBAP","CPHISTOBAP")');
End;

Procedure MajAvant973;
Begin
  //D SCLAVOPOULOS 6140
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE="WJOURNALACTION"');
  //T PETETIN 6150
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE IN ("RQACTIONS","RQACTIONSNC","RQACTIONSPC")');
  //P DAMINETTE 6215
  If TableExiste('SWVIGNETTES') then  ExecuteSQLContOnExcept('DELETE from swvignettes where swv_typevignette = "CP" and swv_codevignette in ("CPVIGNETTEMULBAP", "CPVIGSUIVIBAP") ');
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE IN ("CPVALIDEBAP","CPHISTOBAPVISEUR","CPHISTOBAP","CPCONSECR")');

End;

Procedure MajAvant974;
Begin
  //MNG 6258
  // supprimer dans MajAvant962 le delete from ytradmetier
  // vidage table qui contient l'expression "opération(s)
  ExecuteSQLContOnExcept('DELETE FROM YTRADMETIER');
  //LEK 6273
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE IN ("CPVALIDEBAP" ,"CPSUPPRESSIONBAP" )');
End;

Procedure MajAvant975;
Begin
  //T PETETIN 6325
//  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE IN ("WFORMCONVVARDEF","WFORMCONVVAR","WFORMCONV")');

End;

Procedure MajAvant976;
Begin
  //M MORRETTON 6457
//  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE="GCLISTEINVLIGPRIX"');
End;

Procedure MajAvant979;
Begin
  //M MORRETTON
  ExecuteSQLNoPCL('DELETE FROM CHOIXEXT WHERE YX_TYPE="WTR"');

End;

Procedure MajAvant980;
Begin
  //M MORRETON
  ExecuteSQLContOnExcept('DELETE FROM GRAPHS WHERE GR_GRAPHE IN ("TWPDRVIEWER3","TWPDRCUBE3")  AND GR_LANGUE="---"');
  //JPL
//  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE in ("QULBPSESSIONPGIDA","GCHISTODA","GCMULPIECEDA","GCMULPIECEDAVAL")');
  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE in ("QULBPSESSIONPGIDA")');

End;

Procedure MajAvant981;
Begin
  //JP LAURENT
//  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE in ("QULBPSESSIONPGIDA","GCHISTODA","GCMULPIECEDA","GCMULPIECEDAVAL",'
//  +'"GCHIERARACHIEPAR","HIERARCHIEVAL","QBPIMPLIG","QBPIMPPAR","QBPREVISION")');
  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE in ("QULBPSESSIONPGIDA","HIERARCHIEVAL","QBPIMPLIG","QBPIMPPAR","QBPREVISION")');
  //M MORRETTON
  ExecuteSQLContOnExcept('DELETE FROM GRAPHS WHERE GR_GRAPHE IN ("TWPDRCUBE3")  AND GR_LANGUE="---"');
  //L CHHOEU
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE IN("CPCOMPEN01","CPCOMPEN01SV","CPCOMPEN02","CPCOMPEN02SV")');
  //D KOZA
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE = "WPF_CBNGRPROP"');
End;

Procedure MajAvant982;
Begin
  { JTR FQ;010;10424. Delete de la liste pour la "réimporter" par les bobs }
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE = "PSE_MULEWNOMETET"');
  // TS
  ExecuteSqlNoPCL('DELETE FROM LISTE WHERE LI_LISTE="WORDRECMP"');
  //JPL
//  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE in ("QULBPSESSIONPGIDA","GCHISTODA","GCMULPIECEDA","GCMULPIECEDAVAL",'
//  +'"GCHIERARACHIEPAR","HIERARCHIEVAL","QBPIMPLIG","QBPIMPPAR","QBPREVISION","GCMULPIECEEXPE")');
  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE in ("QULBPSESSIONPGIDA","HIERARCHIEVAL","QBPIMPLIG","QBPIMPPAR","QBPREVISION")');
End;

Procedure MajAvant988;
Begin
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE IN ("CPMULCETEBAC","CPMULCFONB")');
End;

Procedure MajAvant995;
Begin
  //M MORRETTON
  ExecuteSQLContOnExcept('DELETE FROM GRAPHS WHERE GR_GRAPHE IN ("TWPDRCUBE3")');

  //D KOZA
//  ExecuteSqlNoPCL('DELETE FROM LISTE WHERE LI_LISTE="GCMULARTINFOS"');

  //C DUMAS
  SupprimeEtat( 'E','ACE','ACR');

  //N FOURNEL
//  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE="GCGENEREFAA"');

  //D KOZA
//  ExecuteSQLNoPCL('DELETE FROM LISTE WHERE LI_LISTE IN ("GCLISTINVENT","GCLISTINVLIGNE")');

  //D SCLAVOPOULOS
//  ExecuteSQLContOnExcept('DELETE FROM MODELES WHERE MO_TYPE="E" AND MO_NATURE="STK" AND MO_CODE="STD"');
//  ExecuteSQLContOnExcept('DELETE FROM MODEDATA WHERE MD_CLE LIKE "ESTKSTD%"');

  //LEK CHHOEU
  ExecuteSQLNoPCL('DELETE FROM LISTE WHERE LI_LISTE = "CPCLOTPERJOU"');

End;

Procedure MajAvant996;
Begin
//7547
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE="QMESPOINTAGE"');
//7555
//  ExecuteSqlContOnExcept('DELETE FROM LISTE WHERE LI_LISTE = "WJOURNALACTION"');
//7594
//  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE = "WNOMELIGMODIFLOT"');
End;

Procedure MajAvant997;
Begin

  //COMPTA
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_liste IN ("SUPPRANA","ILISTIMO_SUB")');
  //DOMG
//  AglNettoieListes('AFMULFACTIERSAFFP', 'GLA_REGROUPEFACT', nil);
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE LIKE "QUVBPREALISE%"');
  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE in ("QULQBPREALISE","QBPSEQBUDGET","QBPSEQBUDGETMUL","QBPLIENSIMUL")');
//  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE="AFVALINDICE"');

{// BIDOUILLE POUR PERSONNALISATION
// on charge une TOB pour sauvegarde tablette afin de la recharger en majver997
  TobSidos := Tob.Create('TOB SIDOS', nil, -1);
  Sql := 'SELECT * FROM COMMUN WHERE CO_TYPE="TRA" AND CO_LIBRE<>"---"';
  TobSidos.LoadDetailDBFromSql('COMMUN', Sql);}

// 2eme version BIDOUILLE POUR PERSONNALISATION
// on charge une table tempo pour sauvegarde tablette afin de la recharger en majver997

  if isMssql then
  begin
    // ======== MSSQL ===============
    if TableExiste('TmpCommun')
    then ExecuteSQLContOnExcept('drop table TmpCommun');

    ExecuteSQLContOnExcept('select CO_TYPE, Co_CODE, CO_LIBELLE, CO_ABREGE, Co_LIBRE '+
    'into TmpCommun '+
    'from (select * from commun where co_type="TRA" and co_libre<>"---") TCO');
  end
  else if isOracle then
  begin
  // ======== ORACLE =================
    ExecuteSQLContOnExcept('begin '+
    'execute immediate "drop table TmpCommun" ; '+
    'exception '+
    'when others '+
    'then null; '+
    'end;');

    ExecuteSQLContOnExcept('create table TmpCommun as '+
    'select CO_TYPE, Co_CODE, CO_LIBELLE, CO_ABREGE, Co_LIBRE '+
    'from (select * from commun where co_type="TRA" and co_libre<>"---") TCO');
  end;

end;

Procedure MajAvant998;
Begin
  DropVuesStandard;
  DropEtatsStandard;
  DropListesStandard;
  (* --- BTP LSE --- *)
  ExecuteSQLContOnExcept('DELETE FROM DECOMBOS WHERE DO_COMBO="WNATURETRAVAIL"');
  (* --- CCS59001D013 --- *)
  ExecuteSQLContOnExcept ('DELETE FROM NFILES WHERE NFI_FILEGUID IN (SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="CISX" AND YFS_CRIT2="COMPTA" AND YFS_EXTFILE=".CIX" AND YFS_PREDEFINI="CEG")');
  ExecuteSQLContOnExcept ('DELETE FROM NFILEPARTS WHERE NFS_FILEGUID IN (SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="CISX" AND YFS_CRIT2="COMPTA" AND YFS_EXTFILE=".CIX" AND YFS_PREDEFINI="CEG")');
	ExecuteSQLContOnExcept ('DELETE FROM YFILESTD WHERE YFS_CODEPRODUIT="CISX" AND YFS_CRIT2="COMPTA" AND YFS_PREDEFINI="CEG"');
  ExecuteSQLContOnExcept ('DELETE FROM CPGZIMPREQ WHERE (CIS_DOMAINE="E" OR CIS_DOMAINE="O") AND (CIS_TABLE2="BALANCE" OR CIS_TABLE2="JOURNAL") AND (CIS_CLEPAR="STD CEGID")');
  (* --- CCS50998D003  --- *)
  ExecuteSqlContOnExcept ('DELETE FROM NFILEPARTS WHERE NFS_FILEGUID IN (SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="COMPTA" AND YFS_CRIT1="ECH" AND YFS_CRIT2="MOP")');
  ExecuteSqlContOnExcept ('DELETE FROM NFILES WHERE NFI_FILEGUID IN (SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="COMPTA" AND YFS_CRIT1="ECH" AND YFS_CRIT2="MOP")');
	ExecuteSqlContOnExcept ('DELETE FROM YFILESTD WHERE YFS_CODEPRODUIT="COMPTA" AND YFS_CRIT1="ECH" AND YFS_CRIT2="MOP"');
	ExecuteSqlContOnExcept ('DELETE FROM NFILES WHERE NFI_FILEGUID IN (SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="CISX" AND YFS_CRIT2="COMPTA" AND YFS_EXTFILE=".CIX" AND YFS_PREDEFINI="CEG")');
	ExecuteSqlContOnExcept ('DELETE FROM NFILEPARTS WHERE NFS_FILEGUID IN (SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="CISX" AND YFS_CRIT2="COMPTA" AND YFS_EXTFILE=".CIX" AND YFS_PREDEFINI="CEG")');
	ExecuteSqlContOnExcept ('DELETE FROM YFILESTD WHERE YFS_CODEPRODUIT="CISX" AND YFS_CRIT2="COMPTA" AND YFS_PREDEFINI="CEG"');
  ExecuteSqlContOnExcept ('DELETE FROM CPGZIMPREQ WHERE (CIS_DOMAINE="E" OR CIS_DOMAINE="O") AND (CIS_TABLE2="BALANCE" OR CIS_TABLE2="JOURNAL") AND (CIS_CLEPAR="STD CEGID")');
  (* --- CCS59001D013 --- *)
  ExecuteSqlContOnExcept ('DELETE FROM NFILES WHERE NFI_FILEGUID IN (SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="CISX" AND YFS_CRIT2="COMPTA" AND YFS_EXTFILE=".CIX" AND YFS_PREDEFINI="CEG")');
  ExecuteSqlContOnExcept ('DELETE FROM NFILEPARTS WHERE NFS_FILEGUID IN (SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="CISX" AND YFS_CRIT2="COMPTA" AND YFS_EXTFILE=".CIX" AND YFS_PREDEFINI="CEG")');
	ExecuteSqlContOnExcept ('DELETE FROM YFILESTD WHERE YFS_CODEPRODUIT="CISX" AND YFS_CRIT2="COMPTA" AND YFS_PREDEFINI="CEG"');
  ExecuteSqlContOnExcept ('DELETE FROM CPGZIMPREQ WHERE (CIS_DOMAINE="E" OR CIS_DOMAINE="O") AND (CIS_TABLE2="BALANCE" OR CIS_TABLE2="JOURNAL") AND (CIS_CLEPAR="STD CEGID")');
  (* --- CPS50998F008 --- *)
  ExecuteSqlContOnExcept ('SELECT * FROM YFILESTD WHERE YFS_CODEPRODUIT="PAIE" AND YFS_PREDEFINI="CEG" AND YFS_NOM="PAIES5.XLA"');
	ExecuteSqlContOnExcept ('SELECT * FROM NFILES WHERE NFI_FILEGUID IN( SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="PAIE" AND YFS_PREDEFINI="CEG" AND YFS_NOM="PAIES5.XLA")');
	ExecuteSqlContOnExcept ('SELECT * FROM NFILEPARTS WHERE NFS_FILEGUID IN( SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="PAIE" AND YFS_PREDEFINI="CEG" AND YFS_NOM="PAIES5.XLA")');
  (* --- CPS50998D028 --- *)
  ExecuteSqlContOnExcept ('DELETE FROM FILTRES WHERE FI_TABLE="PGMULSALARIE"');
	(* --- CPS50998F066 --- *)
  ExecuteSqlContOnExcept ('DELETE FROM YFILESTD WHERE YFS_CODEPRODUIT="PAIE" AND YFS_PREDEFINI="CEG"');
	ExecuteSqlContOnExcept ('DELETE FROM NFILES WHERE NFI_FILEGUID IN( SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="PAIE" AND YFS_PREDEFINI="CEG")');
  ExecuteSqlContOnExcept ('DELETE FROM NFILEPARTS WHERE NFS_FILEGUID IN( SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="PAIE" AND YFS_PREDEFINI="CEG")');
  (* --- CPS50998D070 --- *)
  ExecuteSqlContOnExcept ('DELETE FROM PARAMSALARIE WHERE PPP_PGINFOSMODIF="PSA_SALARIETYPE"');
  (* --- CPS50998P070 --- *)
  ExecuteSqlContOnExcept ('DELETE FROM CUMULRUBRIQUE WHERE PCR_NATURERUB="AAA" AND PCR_RUBRIQUE="4030" AND (PCR_CUMULPAIE="05" OR PCR_CUMULPAIE="43")');
	ExecuteSqlContOnExcept ('DELETE FROM CUMULRUBRIQUE WHERE PCR_NATURERUB="COT" AND PCR_RUBRIQUE="8550" AND (PCR_CUMULPAIE="13" OR PCR_CUMULPAIE="15")');
	ExecuteSqlContOnExcept ('DELETE FROM PROFILRUB WHERE PPM_PROFIL="031" AND PPM_RUBRIQUE="7210"');
  ExecuteSqlContOnExcept ('DELETE FROM CUMULRUBRIQUE WHERE PCR_RUBRIQUE="8640" AND PCR_CUMULPAIE="35"');
  ExecuteSqlContOnExcept ('DELETE FROM CUMULRUBRIQUE WHERE PCR_RUBRIQUE="8644" AND PCR_CUMULPAIE="35"');
  ExecuteSqlContOnExcept ('DELETE FROM CUMULRUBRIQUE WHERE PCR_RUBRIQUE="8654" AND PCR_CUMULPAIE="35"');
  ExecuteSqlContOnExcept ('DELETE FROM CUMULRUBRIQUE WHERE PCR_RUBRIQUE="8646" AND PCR_CUMULPAIE="35"');
  ExecuteSqlContOnExcept ('DELETE FROM PROFILRUB WHERE PPM_RUBRIQUE="8300" AND PPM_NATURERUB="COT" AND PPM_PREDEFINI="CEG"');
  ExecuteSqlContOnExcept ('DELETE FROM PROFILRUB WHERE PPM_RUBRIQUE="8136" AND PPM_NATURERUB="COT" AND PPM_PREDEFINI="CEG"');
  ExecuteSqlContOnExcept ('DELETE FROM PROFILRUB WHERE PPM_RUBRIQUE="8404" AND PPM_NATURERUB="COT" AND PPM_PREDEFINI="CEG"');
  ExecuteSqlContOnExcept ('DELETE FROM PUBLICOTIS WHERE PUO_PREDEFINI="CEG"');
  ExecuteSqlContOnExcept ('DELETE FROM PROFILRUB WHERE PPM_PROFIL="019" AND PPM_RUBRIQUE="4970"');
  ExecuteSqlContOnExcept ('DELETE FROM PROFILRUB WHERE PPM_PROFIL="374" AND PPM_RUBRIQUE="3578"');
  ExecuteSqlContOnExcept ('DELETE FROM PROFILRUB WHERE PPM_PROFIL="065" AND PPM_RUBRIQUE="4972"');
  ExecuteSqlContOnExcept ('DELETE FROM PROFILRUB WHERE PPM_PROFIL="228" AND PPM_RUBRIQUE="4982"');
  ExecuteSqlContOnExcept ('DELETE FROM PROFILRUB WHERE PPM_PROFIL="096" AND PPM_RUBRIQUE="4968"');
  ExecuteSqlContOnExcept ('DELETE FROM PROFILRUB WHERE PPM_PROFIL="097" AND PPM_RUBRIQUE="4976"');
  ExecuteSqlContOnExcept ('DELETE FROM CUMULRUBRIQUE WHERE PCR_NATURERUB="AAA" AND PCR_RUBRIQUE="3000" AND (PCR_CUMULPAIE="05" OR PCR_CUMULPAIE="06")');
  ExecuteSqlContOnExcept ('DELETE FROM CUMULRUBRIQUE WHERE PCR_NATURERUB="AAA" AND PCR_RUBRIQUE="3030" AND (PCR_CUMULPAIE="05" OR PCR_CUMULPAIE="06")');
  ExecuteSqlContOnExcept ('DELETE FROM CUMULRUBRIQUE WHERE PCR_NATURERUB="AAA" AND PCR_RUBRIQUE="3060" AND (PCR_CUMULPAIE="05" OR PCR_CUMULPAIE="06")');
	ExecuteSqlContOnExcept ('DELETE FROM CUMULRUBRIQUE WHERE PCR_NATURERUB="AAA" AND PCR_RUBRIQUE="3240" AND (PCR_CUMULPAIE="05" OR PCR_CUMULPAIE="06")');
 	ExecuteSqlContOnExcept ('DELETE FROM CUMULRUBRIQUE WHERE PCR_NATURERUB="AAA" AND PCR_RUBRIQUE="3270" AND (PCR_CUMULPAIE="05" OR PCR_CUMULPAIE="06")');
  ExecuteSqlContOnExcept ('DELETE FROM CUMULRUBRIQUE WHERE PCR_NATURERUB="AAA" AND PCR_RUBRIQUE="3300" AND (PCR_CUMULPAIE="05" OR PCR_CUMULPAIE="06")');
	ExecuteSqlContOnExcept ('DELETE FROM CUMULRUBRIQUE WHERE PCR_NATURERUB="AAA" AND PCR_RUBRIQUE="1048" AND (PCR_CUMULPAIE="11" OR PCR_CUMULPAIE="12")');
  // Edition 7 compta
  SupprimeEtat('E', 'I58', '054', True);
  SupprimeEtat('E', 'I58', '055', True);
  SupprimeEtat('E', 'I58', 'A54', True);
  SupprimeEtat('E', 'I58', 'A55', True);
  SupprimeEtat('E', 'I58', '855', True);
  SupprimeEtat('E', 'I58', '954', True);
  SupprimeEtat('E', 'I58', '955', True);
  SupprimeEtat('E', 'I90', '054', True);
  SupprimeEtat('E', 'I90', '055', True);
  SupprimeEtat('E', 'I90', 'A54', True);
  SupprimeEtat('E', 'I90', 'A55', True);
  // Edition 6 paie
  ExecuteSqlContOnExcept ('DELETE FROM HISTOSAISRUB WHERE PSD_RUBRIQUE="" AND PSD_ORIGINEMVT="SRB"');
  ExecuteSqlContOnExcept ('DELETE FROM CONTRATTRAVAIL WHERE PCI_SALARIE NOT IN (SELECT PSA_SALARIE FROM SALARIES)');
  ExecuteSQLContOnExcept('DELETE FROM CONTRATTRAVAIL WHERE PCI_SALARIE NOT IN (SELECT PSA_SALARIE FROM SALARIES)');
  ExecuteSQLContOnExcept('DELETE FROM PUBLICOTIS WHERE PUO_NATURERUB="REM" AND PUO_RUBRIQUE="09"');
end;


// =====   MAJAVANT ================================================

function MAJHalleyAvant(VSoc: Integer; MajLab1, MajLab2: TLabel; MajJauge: TProgressBar): boolean;
var sSql:string;
    QVerrou:TQuery;
    savEnabledeshare:boolean;
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

  if (V_PGI.ModePCL='1') then       //js1 traitements pour pgimajver batché expert
  begin
    LogPCL('DOSSIER=' + V_PGI.NoDossier,FicLogPCL);
    LogPCL('DEBUT=' + DateTimeToStr(Now),FicLogPCL);

    if (V_PGI.ModePCL='1') and V_PGI.RunWithParams and isVerrou then
      begin
      Application.terminate;
      Abort; //JS1 OUIEOUIEOUIE  !!!!
      end;

    savEnableDeshare:=V_PGI.EnableDeshare;
    V_PGI.EnableDeshare:=True;
    stSavVerrou:='';
    stSavUser:='';
    sSQL := 'SELECT DOS_VERROU,DOS_UTILISATEUR FROM DOSSIER WHERE DOS_NODOSSIER="'+V_PGI.NoDossier+'"';

    QVerrou := OpenSQL(sSQL, True);
    if not QVerrou.Eof then
    begin
      stSavVerrou := QVerrou.FindField('DOS_VERROU').AsString;
      stSavUser := QVerrou.FindField('DOS_UTILISATEUR').AsString;
    end;
    Ferme(QVerrou);

    if stSavVerrou <> 'MAJ' then
      begin
      ExecuteSQLContOnExcept('UPDATE DOSSIER SET DOS_VERROU="MAJ", DOS_UTILISATEUR="'+GetUser+'" WHERE DOS_NODOSSIER="'+V_PGI.NoDossier+'"');
      result:=false;
      end;
    V_PGI.EnableDeshare:=savEnableDeshare;
  end;

  if V_PGI.SAV then
  begin
    LogAGL('');
    LogAGL('**----------------- '+  DateTimeToStr(Now) + ' -----------------------');
    if V_PGI.ModePCL='1' then
      LogAGL('Traitement du dossier "' + V_PGI.NoDossier + '" en environnement PCL')
    else
      LogAGL('Traitement du dossier "' + V_PGI.NoDossier + '" en environnement Entreprise');
    if IsMonoOuCommune then
      LogAGL('    ==> base commune ou mono')
    else
      LogAGL('    ==> base dossier PCL') ;
    LogAGL('Version de la base   : '+IntToStr(VSoc));
    LogAGL('Version de la Socref : '+IntToStr(V_PGI.NumVersionBase));
    LogAGL('Début MajAvant ' + DateTimeToStr(Now));
  end;

  if (V_PGI.ModePCL <> '1') and not(V_PGI.RunWithParams)
    then
      if (okrafale=-1) then abort;   //js1 gestion du bouton annuler sur la maj en rafale

  //js1120506 : optimisation lancée systématiquement pour les bases qui ne le sont pas deja
{$IFDEF MAJPCL}
  if (V_PGI.DOSSIERPCL) and not (EstBasePclOptimisee) then
  begin
    if V_PGI.SAV then LogAgl (DateTimeToStr(Now) + ' : Lancement de l''optimisation.') ;
    OptimiseDossierPCL;
    if V_PGI.SAV then LogAgl (DateTimeToStr(Now) + ' : fin de l''optimisation.') ;
  end else if V_PGI.SAV then LogAgl ('Dossier déjà optimisé.') ;
{$ENDIF}
  // Mise en place des bobs derniere generation version 2008
  BOB_IMPORT_PCL_STD('COMPLETE-2008','Complet version 2008',False);

  //js1 220606 : on recharge le schéma juste après l'optimisation
  ConnecteTrueFalse(DBSOC);

  if V_PGI.SAV then LogAgl ('Après ConnecteTrueFalse') ;
  V_PGI.dedejacharge :=False;

  ChargeTablePrefixe(FALSE,TRUE);
  if V_PGI.SAV then LogAgl ('Après ChargeTablePrefixe') ;

  ForceVerrouille;
  // protection des menus
  ExecuteSQLContOnExcept('UPDATE MENU SET MN_VERSIONDEV="-" WHERE MN_VERSIONDEV="X"');
  // --

  // ----- V9 --------------------------------------
  If VSoc < 900 Then MajAvant900; //Nouvelle SOCREF V900
  If VSoc < 902 Then MajAvant902; //Pas de MajAvant901 ...
  If VSoc < 903 Then MajAvant903;
  If VSoc < 904 Then MajAvant904;
  If VSoc < 905 Then MajAvant905;
  If VSoc < 910 Then MajAvant910;
  If VSoc < 911 Then MajAvant911;
  If VSoc < 912 Then MajAvant912;
  If VSoc < 913 Then MajAvant913;
  If VSoc < 917 Then MajAvant917;
  If VSoc < 918 Then MajAvant918;
  If VSoc < 919 Then MajAvant919;
  If VSoc < 920 Then MajAvant920;
  If VSoc < 921 Then MajAvant921;
  If VSoc < 922 Then MajAvant922;
  If VSoc < 923 Then MajAvant923;
  If VSoc < 924 Then MajAvant924;
  If VSoc < 925 Then MajAvant925;
  If VSoc < 926 Then MajAvant926;
  If VSoc < 927 Then MajAvant927;
  If VSoc < 928 Then MajAvant928;
  If VSoc < 929 Then MajAvant929;
  If VSoc < 930 Then MajAvant930;
  If VSoc < 931 Then MajAvant931;
  If VSoc < 932 Then MajAvant932;
  If VSoc < 933 Then MajAvant933;
  If VSoc < 934 Then MajAvant934;
  If VSoc < 935 Then MajAvant935;
  If VSoc < 936 Then MajAvant936;
  If VSoc < 937 Then MajAvant937;
  If VSoc < 938 Then MajAvant938;
  If VSoc < 939 Then MajAvant939;
  If VSoc < 940 Then MajAvant940;
  If VSoc < 941 Then MajAvant941;
  If VSoc < 942 Then MajAvant942;
  If VSoc < 943 Then MajAvant943;
  If VSoc < 944 Then MajAvant944;
  If VSoc < 947 Then MajAvant947;
  If VSoc < 948 Then MajAvant948;
  If VSoc < 951 Then MajAvant951;
  If VSoc < 952 Then MajAvant952;
  If VSoc < 953 Then MajAvant953;
  If VSoc < 954 Then MajAvant954;
  If VSoc < 955 Then MajAvant955;
  If VSoc < 956 Then MajAvant956;
  If VSoc < 957 Then MajAvant957;
  If VSoc < 958 Then MajAvant958;
  If VSoc < 959 Then MajAvant959;
  If VSoc < 960 Then MajAvant960;
  If VSoc < 961 Then MajAvant961;
  If VSoc < 962 Then MajAvant962;
  If VSoc < 963 Then MajAvant963;
  If VSoc < 964 Then MajAvant964;
  If VSoc < 965 Then MajAvant965;
  If VSoc < 966 Then MajAvant966;
  If VSoc < 967 Then MajAvant967;
  If VSoc < 968 Then MajAvant968;
  If VSoc < 971 Then MajAvant971;
  If VSoc < 972 Then MajAvant972;
  If VSoc < 973 Then MajAvant973;
  If VSoc < 974 Then MajAvant974;
  If VSoc < 975 Then MajAvant975;
  If VSoc < 976 Then MajAvant976;
  If VSoc < 979 Then MajAvant979;
  If VSoc < 980 Then MajAvant980;
  If VSoc < 981 Then MajAvant981;
  If VSoc < 982 Then MajAvant982;
  If VSoc < 988 Then MajAvant988;
  If VSoc < 995 Then MajAvant995;
  If VSoc < 996 Then MajAvant996;
  If VSoc < 997 Then MajAvant997; //V9ED2
  If VSoc < 998 Then MajAvant998; //V9ED2

  Result := True;
end;

{==============================================================================}
{=========================== PROCEDURES MAJ APRES =============================}
{==============================================================================}

procedure UpDateDecoupeEcr(SetSQL: string ; WhereSupJal : String = ''; WhereSupExo : String = '' ; WheresupEcr : String = '' ; DecoupeParExo : Boolean = TRUE);
var DMin, DMax, DD1, DD2: TDateTime;
  ListeJ: TStrings;
  Q, QExo: TQuery;
  i, iper, Delta, IperMax: integer;
  St  :String ;
begin
  // Lecture des journaux
  ListeJ := TStringList.Create;
  St:='Select J_JOURNAL from JOURNAL' ; If WhereSupJal<>'' Then st:=St+' WHERE '+WhereSupJal ;
  Q := OpenSQL(St, True);
  while not Q.EOF do
  begin
    ListeJ.Add(Q.Fields[0].AsString);
    Q.Next;
  end;
  Ferme(Q);
  // Balayage des écritures avec découpe
  for i := 0 to ListeJ.Count - 1 do
  begin
    St:='Select EX_EXERCICE, EX_DATEDEBUT,EX_DATEFIN from EXERCICE' ; If WhereSupExo<>'' Then st:=St+' WHERE '+WhereSupExo ;
    QExo := OpenSQl(St, True);
    while not QExo.EOF do
    begin
      DMin := QExo.Fields[1].AsDateTime;
      DMax := QExo.Fields[2].AsDateTime;
      Delta := Round((DMax - DMin) / 10);
      If DecoupeParExo then IperMax:=1 Else iPerMax:=10 ;
      for iper := 1 to iPerMax do
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
        If DecoupeParExo then BEGIN DD1:=DMin ; DD2:=DMax ; END ;
        BeginTrans;
        St:='UPDATE ECRITURE SET ' + SetSQL + ' WHERE E_JOURNAL="' + ListeJ[i] + '" AND E_EXERCICE="' + QExo.Fields[0].AsString +
          '" AND E_DATECOMPTABLE>="' + UsDateTime_(DD1) + '" AND E_DATECOMPTABLE<="' + UsDateTime_(DD2) + '"' ;
        If WhereSupEcr<>'' Then st:=St+' AND ('+WhereSupEcr+') ' ;
        ExecuteSQLContOnExcept(St);
        CommitTrans;
      end;
      QExo.Next;
    end;
    Ferme(QExo);
  end;
  ListeJ.Clear;
  ListeJ.Free;
end;

procedure UpDateDecoupeLigneOuvPlat(SetSQL: string; SetCondition: string='');
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
  Q := OpenSQL('Select DISTINCT BOP_NATUREPIECEG, BOP_SOUCHE FROM LIGNEOUVPLAT', True);
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
    QMax := OpenSQL('SELECT MIN(BOP_NUMERO), MAX(BOP_NUMERO),COUNT(*) FROM LIGNEOUVPLAT WHERE BOP_NATUREPIECEG="' + ListeNat[i] + '" AND BOP_SOUCHE="' + ListeSouche[i] +
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
        QMax := OpenSQL('SELECT MIN(BOP_NUMERO), MAX(BOP_NUMERO),COUNT(*) FROM LIGNEOUVPLAT WHERE BOP_NATUREPIECEG="' + ListeNat[i] + '" AND BOP_SOUCHE="' +
          ListeSouche[i] + '" AND BOP_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND BOP_NUMERO<' + IntToStr(NbFinTranche) + '', True);
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
            ExecuteSQL('UPDATE LIGNEOUVPLAT SET ' + SetSQL + ' WHERE BOP_NATUREPIECEG="' + ListeNat[i] + '" AND BOP_SOUCHE="' + ListeSouche[i] + '" AND BOP_NUMERO>=' +
                        IntToStr(NN1) + ' AND BOP_NUMERO<=' + IntToStr(NN2) + ' ' + SetCondition );
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
      // sinon trou ds la numéraotation, update pour la souche
    end else ExecuteSQL('UPDATE LIGNEOUVPLAT SET ' + SetSQL + ' WHERE BOP_NATUREPIECEG="' + ListeNat[i] + '" AND BOP_SOUCHE="' + ListeSouche[i] + '"' + ' ' + SetCondition);
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
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
            ExecuteSQLContOnExcept('UPDATE LIGNE SET ' + SetSQL + ' WHERE GL_NATUREPIECEG="' + ListeNat[i] + '" AND GL_SOUCHE="' + ListeSouche[i] + '" AND GL_NUMERO>=' +
                        IntToStr(NN1) + ' AND GL_NUMERO<=' + IntToStr(NN2) + ' ' + SetCondition );
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
      // sinon trou ds la numéraotation, update pour la souche
    end else ExecuteSQLContOnExcept('UPDATE LIGNE SET ' + SetSQL + ' WHERE GL_NATUREPIECEG="' + ListeNat[i] + '" AND GL_SOUCHE="' + ListeSouche[i] + '"' + ' ' + SetCondition);
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;

procedure UpDateDecoupeLigneBase(SetSQL: string; SetCondition : string );
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
  Q := OpenSQL('Select DISTINCT BLB_NATUREPIECEG, BLB_SOUCHE FROM LIGNEBASE', True);
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
    QMax := OpenSQL('SELECT MIN(BLB_NUMERO), MAX(BLB_NUMERO),COUNT(*) FROM LIGNEBASE WHERE BLB_NATUREPIECEG="' + ListeNat[i] + '" AND BLB_SOUCHE="' + ListeSouche[i] +
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
        QMax := OpenSQL('SELECT MIN(BLB_NUMERO), MAX(BLB_NUMERO),COUNT(*) FROM LIGNEBASE WHERE BLB_NATUREPIECEG="' + ListeNat[i] + '" AND BLB_SOUCHE="' +
          ListeSouche[i] + '" AND BLB_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND BLB_NUMERO<' + IntToStr(NbFinTranche) + '', True);
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
            ExecuteSQL('UPDATE LIGNEBASE SET ' + SetSQL + ' WHERE BLB_NATUREPIECEG="' + ListeNat[i] + '" AND BLB_SOUCHE="' + ListeSouche[i] + '" AND BLB_NUMERO>=' +
                        IntToStr(NN1) + ' AND BLB_NUMERO<=' + IntToStr(NN2) + ' ' + SetCondition );
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
      // sinon trou ds la numéraotation, update pour la souche
    end else ExecuteSQL('UPDATE LIGNEBASE SET ' + SetSQL + ' WHERE BLB_NATUREPIECEG="' + ListeNat[i] + '" AND BLB_SOUCHE="' + ListeSouche[i] + '"' + ' ' + SetCondition);
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;

procedure UpDateDecoupeLigneFac(SetSQL: string; SetCondition : string ='');
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
  Q := OpenSQL('Select DISTINCT BLF_NATUREPIECEG, BLF_SOUCHE FROM LIGNEFAC', True);
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
    QMax := OpenSQL('SELECT MIN(BLF_NUMERO), MAX(BLF_NUMERO),COUNT(*) FROM LIGNEFAC WHERE BLF_NATUREPIECEG="' + ListeNat[i] + '" AND BLF_SOUCHE="' + ListeSouche[i] +
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
        QMax := OpenSQL('SELECT MIN(BLF_NUMERO), MAX(BLF_NUMERO),COUNT(*) FROM LIGNEFAC WHERE BLF_NATUREPIECEG="' + ListeNat[i] + '" AND BLF_SOUCHE="' +
          ListeSouche[i] + '" AND BLF_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND BLF_NUMERO<' + IntToStr(NbFinTranche) + '', True);
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
            ExecuteSQL('UPDATE LIGNEFAC SET ' + SetSQL + ' WHERE BLF_NATUREPIECEG="' + ListeNat[i] + '" AND BLF_SOUCHE="' + ListeSouche[i] + '" AND BLF_NUMERO>=' +
                        IntToStr(NN1) + ' AND BLF_NUMERO<=' + IntToStr(NN2) + ' ' + SetCondition );
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
      // sinon trou ds la numéraotation, update pour la souche
    end else ExecuteSQL('UPDATE LIGNEFAC SET ' + SetSQL + ' WHERE BLF_NATUREPIECEG="' + ListeNat[i] + '" AND BLF_SOUCHE="' + ListeSouche[i] + '"' + ' ' + SetCondition);
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;

procedure UpDateDecoupePiece(SetSQL:string; SetCondition: string='');
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
              IntToStr(NN1) + ' AND GP_NUMERO<=' + IntToStr(NN2)+ ' ' + SetCondition );
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
        + ' WHERE GP_NATUREPIECEG="' + ListeNat[i] + '" AND GP_SOUCHE="' + ListeSouche[i] + '"'+ ' ' + SetCondition );
    end;
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;

procedure UpDateDecoupePieceRG(SetSQL:string; SetCondition: string='');
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
  Q := OpenSQL('Select DISTINCT PRG_NATUREPIECEG, PRG_SOUCHE FROM PIECERG', True);
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
    QMax := OpenSQL('SELECT MIN(PRG_NUMERO), MAX(PRG_NUMERO),COUNT(*) FROM PIECERG WHERE PRG_NATUREPIECEG="' + ListeNat[i] + '" AND PRG_SOUCHE="' + ListeSouche[i] +
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
        QMax := OpenSQL('SELECT MIN(PRG_NUMERO), MAX(PRG_NUMERO),COUNT(*) FROM PIECERG WHERE PRG_NATUREPIECEG="' + ListeNat[i] + '" AND PRG_SOUCHE="' +
          ListeSouche[i] + '" AND PRG_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND PRG_NUMERO<' + IntToStr(NbFinTranche) + '', True);
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
            ExecuteSQL('UPDATE PIECERG SET ' + SetSQL + ' WHERE PRG_NATUREPIECEG="' + ListeNat[i] + '" AND PRG_SOUCHE="' + ListeSouche[i] + '" AND PRG_NUMERO>=' +
              IntToStr(NN1) + ' AND PRG_NUMERO<=' + IntToStr(NN2)+ ' ' + SetCondition );
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
      ExecuteSQL('UPDATE PIECERG SET ' + SetSQL
        + ' WHERE PRG_NATUREPIECEG="' + ListeNat[i] + '" AND PRG_SOUCHE="' + ListeSouche[i] + '"'+ ' ' + SetCondition );
    end;
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;

procedure UpDateDecoupePiedBase(SetSQL: string; SetCondition : string ='');
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
  Q := OpenSQL('Select DISTINCT GPB_NATUREPIECEG, GPB_SOUCHE FROM PIEDBASE', True);
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
    QMax := OpenSQL('SELECT MIN(GPB_NUMERO), MAX(GPB_NUMERO),COUNT(*) FROM PIEDBASE WHERE GPB_NATUREPIECEG="' + ListeNat[i] + '" AND GPB_SOUCHE="' + ListeSouche[i] +
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
        QMax := OpenSQL('SELECT MIN(GPB_NUMERO), MAX(GPB_NUMERO),COUNT(*) FROM PIEDBASE WHERE GPB_NATUREPIECEG="' + ListeNat[i] + '" AND GPB_SOUCHE="' +
          ListeSouche[i] + '" AND GPB_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND GPB_NUMERO<' + IntToStr(NbFinTranche) + '', True);
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
            ExecuteSQL('UPDATE PIEDBASE SET ' + SetSQL + ' WHERE GPB_NATUREPIECEG="' + ListeNat[i] + '" AND GPB_SOUCHE="' + ListeSouche[i] + '" AND GPB_NUMERO>=' +
              IntToStr(NN1) + ' AND GPB_NUMERO<=' + IntToStr(NN2)+ ' ' + SetCondition );
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
      ExecuteSQL('UPDATE PIEDBASE SET ' + SetSQL
        + ' WHERE GPB_NATUREPIECEG="' + ListeNat[i] + '" AND GPB_SOUCHE="' + ListeSouche[i] + '"'+ ' ' + SetCondition );
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
            ExecuteSQLContOnExcept('UPDATE PIEDECHE SET ' + SetSQL + ' WHERE GPE_NATUREPIECEG="' + ListeNat[i] + '" AND GPE_SOUCHE="' + ListeSouche[i] + '" AND GPE_NUMERO>=' +
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
      ExecuteSQLContOnExcept('UPDATE PIEDECHE SET ' + SetSQL
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
            ExecuteSQLContOnExcept('UPDATE ACOMPTES SET ' + SetSQL + ' WHERE GAC_NATUREPIECEG="' + ListeNat[i] + '" AND GAC_SOUCHE="' + ListeSouche[i] + '" AND GAC_NUMERO>=' +
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
      ExecuteSQLContOnExcept('UPDATE ACOMPTES SET ' + SetSQL
        + ' WHERE GAC_NATUREPIECEG="' + ListeNat[i] + '" AND GAC_SOUCHE="' + ListeSouche[i] + '"');
    end;
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;

// BTP
procedure UpDateDecoupeLigneCompl(SetSQL: string; SetCondition : string ='');
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
  Q := OpenSQL('Select DISTINCT GLC_NATUREPIECEG, GLC_SOUCHE FROM LIGNECOMPL', True);
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
    QMax := OpenSQL('SELECT MIN(GLC_NUMERO), MAX(GLC_NUMERO),COUNT(*) FROM LIGNECOMPL WHERE GLC_NATUREPIECEG="' + ListeNat[i] + '" AND GLC_SOUCHE="' + ListeSouche[i] +
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
        QMax := OpenSQL('SELECT MIN(GLC_NUMERO), MAX(GLC_NUMERO),COUNT(*) FROM LIGNECOMPL WHERE GLC_NATUREPIECEG="' + ListeNat[i] + '" AND GLC_SOUCHE="' +
          ListeSouche[i] + '" AND GLC_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND GLC_NUMERO<' + IntToStr(NbFinTranche) + '', True);
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
            ExecuteSQL('UPDATE LIGNECOMPL SET ' + SetSQL + ' WHERE GLC_NATUREPIECEG="' + ListeNat[i] + '" AND GLC_SOUCHE="' + ListeSouche[i] + '" AND GLC_NUMERO>=' +
                        IntToStr(NN1) + ' AND GLC_NUMERO<=' + IntToStr(NN2) + ' ' + SetCondition );
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
      // sinon trou ds la numéraotation, update pour la souche
    end else ExecuteSQL('UPDATE LIGNECOMPL SET ' + SetSQL + ' WHERE GLC_NATUREPIECEG="' + ListeNat[i] + '" AND GLC_SOUCHE="' + ListeSouche[i] + '"' + ' ' + SetCondition);
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;


procedure UpDateDecoupeLigneOUV(SetSQL: string; SetCondition : string ='');
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
  Q := OpenSQL('Select DISTINCT BLO_NATUREPIECEG, BLO_SOUCHE FROM LIGNEOUV', True);
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
    QMax := OpenSQL('SELECT MIN(BLO_NUMERO), MAX(BLO_NUMERO),COUNT(*) FROM LIGNEOUV WHERE BLO_NATUREPIECEG="' + ListeNat[i] + '" AND BLO_SOUCHE="' + ListeSouche[i] +
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
        QMax := OpenSQL('SELECT MIN(BLO_NUMERO), MAX(BLO_NUMERO),COUNT(*) FROM LIGNEOUV WHERE BLO_NATUREPIECEG="' + ListeNat[i] + '" AND BLO_SOUCHE="' +
          ListeSouche[i] + '" AND BLO_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND BLO_NUMERO<' + IntToStr(NbFinTranche) + '', True);
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
            ExecuteSQL('UPDATE LIGNEOUV SET ' + SetSQL + ' WHERE BLO_NATUREPIECEG="' + ListeNat[i] + '" AND BLO_SOUCHE="' + ListeSouche[i] + '" AND BLO_NUMERO>=' +
                        IntToStr(NN1) + ' AND BLO_NUMERO<=' + IntToStr(NN2) + ' ' + SetCondition );
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
      // sinon trou ds la numéraotation, update pour la souche
    end else ExecuteSQL('UPDATE LIGNEOUV SET ' + SetSQL + ' WHERE BLO_NATUREPIECEG="' + ListeNat[i] + '" AND BLO_SOUCHE="' + ListeSouche[i] + '"' + ' ' + SetCondition);
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;
// --

//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><
// DEBUT MAJ APRES
//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Procedure MajVer625;
var SS : string;
    i:integer;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  SS := UsDateTime(IDate1900);

  // BTP
  ExecuteSQL('UPDATE AFFAIRE SET AFF_OKSIZERO="-" WHERE AFF_OKSIZERO IS NULL' );
  // --
  //GRC
  { ajout 2 champs obligatoires dans la liste }
  AglNettoieListes('RTMULACTIONS', 'RAC_TYPEACTION;RAC_INTERVENANT',nil);
  { bug champs libres suspects : on remplace les codes TX0..2 par TL0..2 }
  if ExisteSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="RSZ" AND CC_CODE LIKE "TX%"') then
  begin
    //for i := 0 to 2 Do
       //ExecuteSQLContOnExcept('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE ,CC_LIBELLE,CC_ABREGE,CC_LIBRE) SELECT "RSZ","TL'+IntToStr(i)+'",CC_LIBELLE,CC_ABREGE,CC_LIBRE FROM CHOIXCOD WHERE CC_TYPE="RSZ" and cc_code="TX'+IntToStr(i)+'"');
    for i := 0 to 2 do InsertChoixCode('RSZ', 'TL' + intTostr(i), 'Texte libre ' + intTostr(i + 1), '', '');
    ExecuteSql ('DELETE FROM CHOIXCOD WHERE CC_TYPE="RSZ" AND CC_CODE LIKE "TX%"');
  end;

  if not IsDossierPCL then
  begin
    //GPAO
    ExecuteSQLContOnExcept('UPDATE WPDRTET  SET WPE_AVECQPCB="-"');
    ExecuteSQLContOnExcept('UPDATE WPDRTYPE SET WRT_AVECQPCB="-"');
    //GC
    ExecuteSQLContOnExcept('UPDATE PORT SET GPO_DATESUP="' + USDATETIME(iDate2099) + '" WHERE GPO_DATESUP is null');

    //-- AJOUT DE DONNEES DANS LA TABLE EDIFAMILLEEMG --//
    if not ExisteSQL('SELECT EFM_CODEFAMILLE FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="ARC"') then
      ExecuteSQLContOnExcept('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE, EFM_NATURESPIECE, EFM_SERIALISE) '
                +'VALUES("ARC", "Accusé de réception de commande EDI_ARC", "ACV,CC,CCE,CCR", "-")')
    else
      ExecuteSQLContOnExcept('UPDATE EDIFAMILLEEMG SET EFM_NATURESPIECE="ACV,CC,CCE,CCR", EFM_SERIALISE="-"'
                +'WHERE EFM_CODEFAMILLE="ARC"');

    if not ExisteSQL('SELECT EFM_CODEFAMILLE FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="BL"') then
      ExecuteSQLContOnExcept('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE, EFM_NATURESPIECE, EFM_SERIALISE) '
                +'VALUES("BL", "Bon de livraison standard EDI_BL", "ALF,ALV,APV,BLC,LCE,LCR,LFR,PRE", "-")')
    else
      ExecuteSQLContOnExcept('UPDATE EDIFAMILLEEMG SET EFM_NATURESPIECE="ALF,ALV,APV,BLC,LCE,LCR,LFR,PRE", EFM_SERIALISE="-"'
                +'WHERE EFM_CODEFAMILLE="BL"');

    if not ExisteSQL('SELECT EFM_CODEFAMILLE FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="CDE"') then
      ExecuteSQLContOnExcept('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE, EFM_NATURESPIECE, EFM_SERIALISE) '
                +'VALUES("CDE", "Commande standard EDI_CDE", "ACV,CC,CCE,CCR", "-")')
    else
      ExecuteSQLContOnExcept('UPDATE EDIFAMILLEEMG SET EFM_NATURESPIECE="ACV,CC,CCE,CCR", EFM_SERIALISE="-"'
                +'WHERE EFM_CODEFAMILLE="CDE"');

    if not ExisteSQL('SELECT EFM_CODEFAMILLE FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="FAC"') then
      ExecuteSQLContOnExcept('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE, EFM_NATURESPIECE, EFM_SERIALISE) '
                +'VALUES("FAC", "Facture EDI_FAC", "ABT,AF,AFP,AFS,APR,AVC,AVS,FAC,FBT", "-")')
    else
      ExecuteSQLContOnExcept('UPDATE EDIFAMILLEEMG SET EFM_NATURESPIECE="ABT,AF,AFP,AFS,APR,AVC,AVS,FAC,FBT", EFM_SERIALISE="-"'
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
  ExecuteSQLContOnExcept('UPDATE MODEPAIE SET MP_AVECINFOCOMPL="-", MP_AVECNUMAUTOR="-", MP_COPIECBDANSCTRL="-", MP_AFFICHNUMCBUS="-"');

  if not IsDossierPCL then
  begin
    UpdateDecoupePiedEche('GPE_CBINTERNET="", GPE_CBLIBELLE="", GPE_DATEEXPIRE="", GPE_TYPECARTE="", GPE_CBNUMCTRL="", GPE_CBNUMAUTOR="", GPE_NUMCHEQUE=""');
    UpdateDecoupeAcomptes('GAC_CBNUMCTRL="", GAC_CBNUMAUTOR=""');
    //-- Tables PARCAISSE --//
    ExecuteSQLContOnExcept('UPDATE PARCAISSE SET GPK_CLISAISIENOM="-"');   //114
    //BTP
      //-- Table AFFAIRE --//
    (*
    ExecuteSQL('UPDATE AFFAIRE SET AFF_PREPARE = "-" WHERE AFF_ETATAFFAIRE IN ("ENC","ACP","ACC","REF")');
    ExecuteSQL('UPDATE AFFAIRE SET AFF_PREPARE = "X" WHERE AFF_ETATAFFAIRE IN ("CLO","TER")');
    ExecuteSQL('UPDATE LIGNE SET GL_QTERESTE = GL_QTEFACT WHERE GL_NATUREPIECEG IN ("DBT","FBT","DAC","ABT","ETU")');
    ExecuteSQL('UPDATE BTPARDOC SET BPD_DESCREMPLACE = "-"');
    *)
    //GPAO
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_LIBELLE= "Réception SST d''achat" WHERE GPP_NATUREPIECEG="BSA"');
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
    ExecuteSQLContOnExcept('UPDATE PARCAISSE SET GPK_TPEREPERTOIRE="",GPK_CLAVIERPISTE="-"'); //115
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
    //GC   // js1 pour joel sich 230606 ACH;VEN; au lieu de ''
    ExecuteSQLContOnExcept('UPDATE CONDITIONNEMENT SET GCO_DEFAUT="-",GCO_TYPESFLUX="ACH;VEN;"');
  end; // not IsDossierPCL
End;


Procedure MajVer630;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSQLContOnExcept('update imoref set irf_statut="" where irf_statut is null');
    ExecuteSQLContOnExcept('update imovi1 set iV1_statut="" where iV1_statut is null');
    ExecuteSQLContOnExcept('update imovi2 set iV2_statut="" where iV2_statut is null');
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
       ExecuteSQLContOnExcept('UPDATE ARTICLECOMPL SET GA2_ESTUL="-" WHERE GA2_ESTUL IS NULL');
  end; //PasDossierPGI
End;

Procedure MajVer633;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  ExecuteSQLContOnExcept('UPDATE YFILES SET YFI_FILECOMPRESSED = "X", YFI_FILESTORAGE = -1, YFI_FILEDATECREATE = YFI_DATECREATION, YFI_FILEDATEMODIFY = YFI_DATECREATION WHERE YFI_FILESTORAGE IS NULL');
  if not IsDossierPCL then
  begin
    //---------- ARTICLE ------------------
    ExecuteSQLContOnExcept('update ARTICLE set ga_lasstraitance="MAN" where ga_lasstraitance is null');
    //---------- PROFILART ------------------
    ExecuteSQLContOnExcept('update PROFILART set gpf_lasstraitance="MAN" where gpf_lasstraitance is null');
    //---------- WNOMELIG ------------------
    ExecuteSQLContOnExcept('update WNOMELIG set wnl_lasstraitance="MAN" where wnl_lasstraitance is null');
    //---------- WORDREBES ------------------
    ExecuteSQLContOnExcept('update WORDREBES set wob_lasstraitance="MAN" where wob_lasstraitance is null');
    //---------- WNOMEDEC ------------------
    ExecuteSQLContOnExcept('update WNOMEDEC set wnd_lasstraitance="MAN" where wnd_lasstraitance is null');
    //--------- WINITCHAMPLIG -----------------------
    if not ExisteSQL('SELECT 1 FROM  WINITCHAMPLIG WHERE WIL_NOMTABLE="ARTICLE" and WIL_NOMCHAMP="GA_LASSTRAITANCE" and WIL_INITCTXTYPE="DEF" and  WIL_IDENTIFIANT=75 and WIL_NUMORDRE=12') then
      ExecuteSQLContOnExcept('insert into WINITCHAMPLIG (WIL_NOMTABLE,WIL_NOMCHAMP,WIL_INITCTXTYPE,WIL_INITCTXCODE'
               +',WIL_IDENTIFIANT,WIL_NUMORDRE,WIL_PRIORITE, WIL_C1OPERATEUR, WIL_C1VALUE'
               +',WIL_C2OPERATEUR, WIL_C2BRANCHEMENT,WIL_C2VALUE,WIL_C3OPERATEUR'
               +',WIL_C3BRANCHEMENT,WIL_C3VALUE,WIL_INITVALEUR'
               +',WIL_DEFINITVALEUR,WIL_OBLIGATOIRE,WIL_CUSTOMIZABLE,WIL_FAMCHAMP,WIL_BYPARAMSOC,WIL_NUMVERSIONTET) '
               +' values ("ARTICLE","GA_LASSTRAITANCE","DEF","DEFAUT",75,12,"","","","","","","","","","MAN","MAN","X","X","DIV","-",1)');
    //---------- WCHAMP ------------------
    ExecuteSQLContOnExcept('delete FROM WCHAMP where WCA_CONTEXTEPROFIL="PRO" and WCA_NOMTABLE="ARTICLE" and WCA_NOMCHAMP="GA_DELAIPROD"');
    if not ExisteSQL('SELECT 1 FROM   WCHAMP WHERE WCA_CONTEXTEPROFIL="PRO" and WCA_NOMTABLE="ARTICLE" and WCA_NOMCHAMP="GA_LASSTRAITANCE"') then
      ExecuteSQLContOnExcept('insert into WCHAMP (WCA_CONTEXTEPROFIL,WCA_NOMTABLE,WCA_NOMCHAMP) values ("PRO","ARTICLE","GA_LASSTRAITANCE")');
     //GC
    ExecuteSQLContOnExcept('UPDATE CONDITIONNEMENT SET GCO_SAISICOND="-",GCO_TYPESFLUX="VEN;",GCO_DECOND="-" WHERE GCO_TYPESFLUX is null');
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
      ExecuteSQLContOnExcept('insert into WCBNTYPEMOUV (WTM_TYPEMOUVEMENT,WTM_LIBELLE) values ("PRP","Proposition transfert ID- prévision")');
    end;
    ExecuteSQLContOnExcept('update WCBNTYPEMOUV set WTM_LIBELLE="Proposition transfert ID- ferme" where WTM_TYPEMOUVEMENT="PRT"');
    ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA="CC;CCR;" WHERE SOC_NOM="SO_WORDREVTE"');
    ExecuteSQLContOnExcept('update parpiece set gpp_typearticle = "CNS;" || gpp_typearticle where gpp_typearticle not like "%CNS%"');
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
     // MCD pour 651 if Vers_afcumul < 110 then    ExecuteSQLContOnExcept('UPDATE AFCUMUL SET ACU_TRAITECPTA = "X"');
     if Vers_activite < 130 then    ExecuteSQLContOnExcept('UPDATE ACTIVITE SET ACT_MONTANTTTC = 0');
     ExecuteSQLContOnExcept('UPDATE EACTIVITE SET EAC_MONTANTTTC = 0,EAC_MONTANTTVA=0');
     if  Vers_ArticleCompl  < 105 then    ExecuteSQLContOnExcept('UPDATE ARTICLECOMPL SET GA2_REMBOURSABLE = "X",GA2_RECUPERABLE="X",GA2_POURCRECUP=100');
  end;  // not IsDossierPCL
End;

Procedure MajVer636;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //GC
  ExecuteSQLContOnExcept('update CPPROFILUSERC set CPU_PCPACHAT = "-" where CPU_PCPACHAT is NULL ');
  ExecuteSQLContOnExcept('update CPPROFILUSERC set CPU_PCPVENTE = "-" where CPU_PCPVENTE is NULL');
  ExecuteSQLContOnExcept('update CPPROFILUSERC set CPU_FORCEDEPOT = "-" where CPU_FORCEDEPOT is NULL');
  ExecuteSQLContOnExcept('update CPPROFILUSERC set CPU_DEPOT = "" where CPU_DEPOT is NULL');
  ExecuteSQLContOnExcept('Update CPPROFILUSERC set CPU_GRPUTI = "GRP" where CPU_USER = "..."');
  ExecuteSQLContOnExcept('Update CPPROFILUSERC set CPU_GRPUTI = "UTI" where CPU_USER <> "..."');
  ExecuteSQLContOnExcept('Update CPPROFILUSERC set CPU_GRPUTI = "" where CPU_GRPUTI is NULL');
  //GPAO
  ExecuteSQLContOnExcept('Update TIERS set T_DELAIMOYEN = 0');

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
     if Vers_activite < 131 then ExecuteSQLContOnExcept('UPDATE ACTIVITE SET ACT_TRAITECPTA = 0');
  end;  // not IsDossierPCL
End;

Procedure MajVer639;
var i : integer;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  if not IsDossierPCL then
  begin
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT=""');
    for i := 26 to 34 do InsertChoixCode('RLZ', 'CL'+chr(i+55), 'Table libre ' + intTostr(i), '', '');

    //GPAO
    ExecuteSQLContOnExcept('UPDATE WNATURETRAVAIL SET WNA_WITHVALWGT1 = WNA_WITHVAL,WNA_WITHMODWGT1 = WNA_WITHMOD,WNA_WITHPERWGT1 = WNA_WITHPER,'
               +'WNA_WITHVALWNT2 = WNA_WITHVAL,WNA_WITHMODWNT2 = WNA_WITHMOD,WNA_WITHPERWNT2 = WNA_WITHPER,WNA_WITHVALWGT2 = WNA_WITHVAL,WNA_WITHMODWGT2 = WNA_WITHMOD,'
               +'WNA_WITHPERWGT2 = WNA_WITHPER,WNA_WITHVALWNT3 = WNA_WITHVAL,WNA_WITHMODWNT3 = WNA_WITHMOD,WNA_WITHPERWNT3 = WNA_WITHPER,WNA_WITHVALWGT3 = WNA_WITHVAL,'
               +'WNA_WITHMODWGT3 = WNA_WITHMOD,WNA_WITHPERWGT3 = WNA_WITHPER,wNA_MVTPRODUITFINI="", WNA_MVTCOMPOSANT="", WNA_MVTCOPRODUIT=""');
    ExecuteSQLContOnExcept('UPDATE WRECORDCMS SET WRD_RESSOURCE="" WHERE WRD_RESSOURCE IS NULL');
    ExecuteSQLContOnExcept('UPDATE DISPO SET GQ_AFFECTE=0, GQ_BLOCAGE=0 ,GQ_EMPLACEMENTACH="",GQ_EMPLACEMENTCON="",  GQ_EMPLACEMENTPRO=""'
               +', GQ_EMPLACEMENTVTE="",GQ_GEREEMPLACE="-",GQ_STOCKALERTE=0, GQ_STOCKRECOMPL=0, GQ_TRANSFERTRECU=0');
  //GRC
    ExecuteSQLContOnExcept('update prospects set rpr_rprlibtable26="",rpr_rprlibtable27="",rpr_rprlibtable28=""'
              +',rpr_rprlibtable29="",rpr_rprlibtable30="",rpr_rprlibtable31="",rpr_rprlibtable32=""'
              +',rpr_rprlibtable33="",rpr_rprlibtable34=""');
  end;  // not IsDossierPCL

  //DP
  ExecuteSQLContOnExcept('update DOSSIER set DOS_NETEXPERT="-", DOS_NECPSEQ=0');      // supprimé en 651
  If Not ExisteSQL('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="JFE" AND CC_CODE="NET"') then
     ExecuteSQLContOnExcept('INSERT INTO CHOIXCOD(CC_TYPE, CC_CODE, CC_LIBELLE, CC_ABREGE, CC_LIBRE) '
     + 'VALUES ("JFE", "NET", "Echanges NetExpert", "NetExpert", "63")' );
End;

Procedure MajVer640;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //DP
  ExecuteSQLContOnExcept('update jubibrubrique set jbr_histo = "-" ' );
  if not IsDossierPCL then
  begin
    // GPAO
    // ExecuteSQLContOnExcept('update wnaturetravail set WNA_MVTPRODUITFINI="APR", WNA_MVTCOMPOSANT="RPR" where WNA_NATURETRAVAIL="FAB"');
    ExecuteSQLContOnExcept('UPDATE WNATURETRAVAIL SET WNA_MVTPRODUITFINI="APR",WNA_MVTCOMPOSANT="RPR",WNA_MVTCOPRODUIT="ACP" WHERE WNA_NATURETRAVAIL="FAB"');
    ExecuteSQLContOnExcept('update article set ga_gereparc = "-" where ga_gereparc is null') ;
    ExecuteSQLContOnExcept('Update WINITCHAMPLIG set WIL_INITVALEUR = "UNI", WIL_DEFINITVALEUR = "UNI" where (WIL_INITVALEUR = "U") OR (WIL_DEFINITVALEUR = "U")');
    //--------- WCBNTYPEMOUV -----------------------
    if not ExisteSQL('select 1 from wcbntypemouv where wtm_typemouvement = "PRP"') then
    begin
      ExecuteSQLContOnExcept('insert into WCBNTYPEMOUV (WTM_TYPEMOUVEMENT,WTM_LIBELLE) values ("PRP","Proposition transfert ID- prévision")');
    end;
    ExecuteSQLContOnExcept('update WCBNTYPEMOUV set WTM_LIBELLE="Proposition transfert ID- ferme" where WTM_TYPEMOUVEMENT="PRT"');
    ExecuteSQLContOnExcept('UPDATE WNATURETRAVAIL SET WNA_TYPENATURE="PRD"');
    ExecuteSQLContOnExcept('UPDATE WORDRETET '
             + 'SET WOT_NOM="", WOT_RESSOURCE="", WOT_DATEINTER="'+UsDateTime(IDate1900)+'", WOT_REFEXTERNE="", WOT_DATEREFEXTERNE="' + USDATETIME(StrToDate('01/01/1900')) + '"'
             + 'WHERE WOT_NOM is null AND WOT_RESSOURCE is null AND WOT_DATEINTER is null AND WOT_REFEXTERNE is null AND WOT_DATEREFEXTERNE is null');
    ExecuteSQLContOnExcept('UPDATE WORDRELIG '
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
  ExecuteSQLContOnExcept('update DOSSIER set DOS_PASSWORD="",DOS_NECPDATEARRET="'+SS+'", DOS_NERECNBFIC=0, DOS_NERECDATE="'+SS+'"');
  ExecuteSQLContOnExcept('UPDATE YDATATYPELINKS SET YDL_LCOMMUNE="-"');
  ExecuteSQLContOnExcept('UPDATE ETABLISS SET ET_NODOSSIER=""');
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
    ExecuteSQLContOnExcept('UPDATE COLISAGE '
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
    ExecuteSQLContOnExcept('UPDATE ARTICLECOMPL SET GA2_MODELE="" WHERE GA2_MODELE is Null');
    ExecuteSQLContOnExcept('update stkmouvement set gsm_affaire = "" where gsm_affaire is null');
    ExecuteSQLContOnExcept('update wordrelig set wol_rang=0, wol_Affaire="" where wol_rang is null and wol_Affaire is null');
    ExecuteSQLContOnExcept('update dispo set GQ_UNITEMIN = "QTE", GQ_UNITEMAX = "QTE", GQ_UNITEALERTE = "QTE", GQ_UNITERECOMPL = "QTE" where GQ_UNITEMIN is null');
    ExecuteSQLContOnExcept('UPDATE QPHASE SET QPH_TRAITEMENT = "" WHERE QPH_TRAITEMENT IS NULL');
    ExecuteSQLContOnExcept('UPDATE QDETCIRC SET QDE_TRAITEMENT = "" WHERE QDE_TRAITEMENT IS NULL');
    ExecuteSQLContOnExcept('UPDATE WORDREPHASE SET WOP_TRAITEMENT = "" WHERE WOP_TRAITEMENT IS NULL');
    //ExecuteSQLContOnExcept('update WORDRETET set wot_affaire="",wot_affaire1="",wot_affaire2="",wot_affaire3="",wot_avenant=""  where wot_affaire is null');
    //ExecuteSQLContOnExcept('update wordrelig set wol_rang=0, wol_Affaire="",wol_Affaire1="",wol_Affaire2=""'
    //            +',wol_Affaire3="",wol_avenant="" '
    //            +' where wol_rang is null and wol_Affaire is null and wol_Affaire1 is null and wol_Affaire2 is null and wol_Affaire3 is null and wol_avenant is null');
    //ExecuteSQLContOnExcept('update WCBNEVOLUTION set wev_datedebut="' + SS  + '", wev_affaire="", wev_affaire1="", wev_affaire2="", wev_affaire3="", wev_avenant="" where wev_affaire1 is null');
    //ExecuteSQLContOnExcept('update stkmouvement set gsm_affaire = "",gsm_affaire1 = "",gsm_affaire2 = "",gsm_affaire3 = "",gsm_avenant = "" where gsm_affaire is null');
    //ExecuteSQLContOnExcept('UPDATE WORDRELIG SET WOL_TYPEDEFAIL="",WOL_NATUREDEFAIL="",WOL_GRAVITEDEFAIL="" WHERE WOL_TYPEDEFAIL IS NULL');
    // modif suite à suppression de champs en 644
    ExecuteSQLContOnExcept('UPDATE STKMOUVEMENT set gsm_affaire = "" where gsm_affaire is null');
    ExecuteSQLContOnExcept('UPDATE WCBNEVOLUTION set wev_datedebut="' + UsDateTime(iDate1900)  + '", wev_affaire="" where wev_affaire is null');
    ExecuteSQLContOnExcept('UPDATE WORDRETET set wot_affaire="" where wot_affaire is null');
    ExecuteSQLContOnExcept('UPDATE WORDRELIG SET WOL_TYPEDEFAIL="",WOL_NATUREDEFAIL="",WOL_GRAVITEDEFAIL="",wol_rang=0, wol_Affaire="" WHERE WOL_TYPEDEFAIL IS NULL');

    ExecuteSQLContOnExcept('UPDATE ARTICLE SET GA_GEREPARC="-" WHERE GA_GEREPARC IS NULL');


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
    ExecuteSQLContOnExcept('UPDATE LIGNECOMPL SET GLC_NATURETRAVAIL = "", GLC_LIGNEORDRE=0, GLC_OPECIRC="", GLC_TRAITEMENT="" WHERE GLC_NATURETRAVAIL IS NULL');
    ExecuteSQLContOnExcept('update STKMOUVEMENT set gsm_codelisteori="" where gsm_codelisteori is null');
    ExecuteSQLContOnExcept('update TRANSINVLIG set GIN_STATUTFLUX="",GIN_LOTINTERNE="",GIN_SERIEINTERNE="",GIN_TIERSPROP="" '
                       +',GIN_INDICEARTICLE="",GIN_MARQUE="", GIN_CHOIXQUALITE="",GIN_STATUTDISPO="",GIN_REFAFFECTATION="" '
                       +',GIN_LOT=(select GA_LOT from ARTICLE where GIN_ARTICLE=GA_ARTICLE)'
                       +',GIN_NUMEROSERIE=(select GA_NUMEROSERIE from ARTICLE where GIN_ARTICLE=GA_ARTICLE) '
                       +' where GIN_REFAFFECTATION is null');
     //GIGA
     GIGAApres644;
  end;  // not IsDossierPCL

  //PAIE
  ExecuteSQLContOnExcept('UPDATE ETABCOMPL SET ETB_EDITBULCP="NN1" WHERE ETB_CONGESPAYES="X"');
  ExecuteSQLContOnExcept('UPDATE ETABCOMPL SET ETB_NBRECPSUPP = 0, ETB_REGIMEALSACE ="-"');
  ExecuteSQLContOnExcept('UPDATE PAIEENCOURS SET PPU_CPACQUISMOD="-"');
  ExecuteSQLContOnExcept('UPDATE SALARIES SET PSA_TYPEDITBULCP="ETB"');
  ExecuteSQLContOnExcept('UPDATE SALARIES SET PSA_CONGESPAYES = ( SELECT ETB_CONGESPAYES FROM ETABCOMPL WHERE  PSA_ETABLISSEMENT = ETB_ETABLISSEMENT)');
  ExecuteSQLContOnExcept('update salaries set psa_cpacquissupp="PER" where psa_nbrecpsupp <> 0');
  ExecuteSQLContOnExcept('update salaries set psa_cpacquissupp="ETB" where psa_nbrecpsupp = 0');
  ExecuteSQLContOnExcept('UPDATE ABSENCESALARIE SET PCN_TYPEIMPUTE=""');
  ExecuteSQLContOnExcept('UPDATE ABSENCESALARIE SET PCN_TYPEIMPUTE="ACQ" WHERE PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="AJU"');
  ExecuteSQLContOnExcept('UPDATE ELTNATIONAUX SET PEL_REGIMEALSACE ="-" WHERE PEL_PREDEFINI <> "CEG"'); // modif P.Dumet 9/8/2004
  // P Dumet 31082004 ExecuteSQLContOnExcept('update variablepaie set PVA_MTARRONDI="7"');
  ExecuteSQLContOnExcept('update variablepaie set PVA_MTARRONDI="7" WHERE PVA_PREDEFINI <> "CEG"');
  ExecuteSQLContOnExcept('UPDATE ATTESTATIONS SET PAS_DATEPREAVISD2="'+SS+'",PAS_DATEPREAVISF2="'+SS+'"'
             +',PAS_MOTIFPREAVIS2="-",PAS_DATEPREAVISD3="'+SS+'",PAS_DATEPREAVISF3="'+SS+'"'
             +',PAS_MOTIFPREAVIS3="-",PAS_MOTIFPREAVIS1="-"');
  // ExecuteSQLContOnExcept('UPDATE DADSLEXIQUE SET PDL_EXERCICEDEB="", PDL_EXERCICEFIN=""');
  // remplacée à partir de 657 par les deux ligne suivantes (Vincent GALLIOT 31/08/2004)
  ExecuteSQLContOnExcept('UPDATE DADSLEXIQUE SET PDL_EXERCICEDEB="" WHERE PDL_EXERCICEDEB IS NULL ');
  ExecuteSQLContOnExcept('UPDATE DADSLEXIQUE SET PDL_EXERCICEFIN="" WHERE PDL_EXERCICEFIN IS NULL ');

  ExecuteSQLContOnExcept('UPDATE ORGANISMEPAIE SET POG_PREVOYANCE="-"');
  ExecuteSQLContOnExcept('UPDATE ETABCOMPL SET ETB_LIBELLE = (SELECT ET_LIBELLE FROM ETABLISS WHERE ETB_ETABLISSEMENT=ET_ETABLISSEMENT)');
  ExecuteSQLContOnExcept('UPDATE ETABCOMPl SET ETB_MEDTRAV=-1 WHERE ETB_MEDTRAV=0 AND NOT EXISTS (select * from annuaire where ANN_TYPEPER="MED" AND ANN_CODEPER=0)');
  // ExecuteSQLContOnExcept('UPDATE JEUECRPAIE SET PJP_TYPEPROVCP=""');
  ExecuteSQLContOnExcept('UPDATE JEUECRPAIE SET PJP_TYPEPROVCP="COD" WHERE PJP_PREDEFINI <> "CEG"'); //P.Dumet 08092004
  // ajout 1509204
  AglNettoieListes('PGMULMVTCP', 'PCN_TYPEMVT;PCN_ORDRE;PCN_ETABLISSEMENT;PCN_SALARIE;PCN_TYPECONGE;PCN_SENSABS;PCN_DATEVALIDITE;PCN_BASE;PCN_NBREMOIS;PCN_JOURS;PCN_APAYES',nil);
  AglNettoieListes('PGENVOIDUCS', 'PES_FICHIEREMIS',nil);

//PAIE
  ExecuteSQLContOnExcept('UPDATE MOTIFABSENCE SET PMA_OKSAISIESAL = "-", PMA_TYPEABS="" , PMA_GESTIONIJSS="-"');
  ExecuteSQLContOnExcept('UPDATE MOTIFABSENCE SET PMA_TYPEABS="RTT" WHERE PMA_TYPERTT="X"');
  ExecuteSQLContOnExcept('UPDATE ABSENCESALARIE SET PCN_NBJCARENCE=0, PCN_NBJCALEND=0, PCN_NBJIJSS=0, PCN_IJSSSOLDEE="-", PCN_MVTORIGINE=""');
  ExecuteSQLContOnExcept('UPDATE ATTESTATIONS SET PAS_TYPEABS=""');
  ExecuteSQLContOnExcept('UPDATE ETABCOMPL SET ETB_SUBROGATION="-"');
  ExecuteSQLContOnExcept('UPDATE SALARIES SET PSA_CATEGMAINTIEN=""');

  if not IsDossierPCL then
  begin
    //GPAO (TARIF)
    // init d'un champ ajouté en 699
    ExecuteSQLContOnExcept('UPDATE YTARIFSPARAMETRES  SET YFO_CODEPORT="" WHERE YFO_CODEPORT IS NULL ');
    // supprimé en 702 GPMajParamTarif;
    GPMajTypeTarif;
    // GPMajSTKNature; mis en fin de MajAvant systématique
    ExecuteSQLContOnExcept('UPDATE TARIF SET GF_TRANSFERE="-", GF_ATRANSFERER="", GF_TRFMESSAGE="" WHERE GF_TRANSFERE IS NULL');
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
    ExecuteSQLContOnExcept('UPDATE ARTICLE SET GA_FAMILLEVALO="" WHERE GA_FAMILLEVALO IS NULL') ;
  end;  // not IsDossierPCL
  //PAIE
  ExecuteSQLContOnExcept('UPDATE CHOIXCOD SET CC_ABREGE=CC_CODE WHERE CC_TYPE="PMS"');
  ExecuteSQLContOnExcept('UPDATE EXERFORMATION SET PFE_TAUXBUDGET=0');
  ExecuteSQLContOnExcept('UPDATE DEPORTSAL SET PSE_MSATECHNIQUE="X"');
  ExecuteSQLContOnExcept('UPDATE DEPORTSAL SET PSE_MSATECHNIQUE="X" WHERE PSE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES LEFT JOIN TAUXAT ON PAT_ORDREAT=PSA_ORDREAT AND PAT_ETABLISSEMENT=PSA_ETABLISSEMENT WHERE PAT_CODEBUREAU<>"BUR")');
  ExecuteSQLContOnExcept('UPDATE DEPORTSAL SET PSE_MSATECHNIQUE="-" WHERE PSE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES LEFT JOIN TAUXAT ON PAT_ORDREAT=PSA_ORDREAT AND PAT_ETABLISSEMENT=PSA_ETABLISSEMENT WHERE PAT_CODEBUREAU="BUR")');
  ExecuteSQLContOnExcept('UPDATE ABSENCESALARIE SET PCN_GESTIONIJSS="-" WHERE PCN_TYPECONGE NOT IN (SELECT PMA_MOTIFABSENCE FROM MOTIFABSENCE WHERE PMA_GESTIONIJSS="X")');
  ExecuteSQLContOnExcept('UPDATE ABSENCESALARIE SET PCN_GESTIONIJSS="X" WHERE PCN_TYPECONGE IN (SELECT PMA_MOTIFABSENCE FROM MOTIFABSENCE WHERE PMA_GESTIONIJSS="X")');
  ExecuteSQLContOnExcept('UPDATE MOTIFABSENCE SET PMA_OKSAISIESAL="-"');
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
    ExecuteSQLContOnExcept('update stkmouvement set gsm_qprepa = 0 where gsm_qprepa is null');
  end;  // not IsDossierPCL
  //PAIE
  ExecuteSQLContOnExcept('UPDATE CURSUSSTAGE SET PCC_MILLESIME="0000"');
  ExecuteSQLContOnExcept('UPDATE CURSUS SET PCU_INCLUSCAT="X"');
  ExecuteSQLContOnExcept('UPDATE COTISATION  SET PCT_ACTIVITE="" WHERE PCT_PREDEFINI<>"CEG"');
  ExecuteSQLContOnExcept('UPDATE REMUNERATION  SET PRM_ACTIVITE="" WHERE PRM_PREDEFINI<>"CEG"');
  //DP
  ExecuteSQLContOnExcept('UPDATE YGEDDICO SET YGD_CODEGED = SUBSTRING(YGD_CODEGED, 1, 1)||"0"||SUBSTRING(YGD_CODEGED, 2, 1) WHERE (LEN(YGD_CODEGED) = 2)');
  ExecuteSQLContOnExcept('UPDATE YDATATYPETREES SET YDT_SCODE = SUBSTRING(YDT_SCODE , 1, 1)||"0"||SUBSTRING(YDT_SCODE , 2, 1) WHERE (LEN(YDT_SCODE ) = 2) AND (YDT_CODEHDTLINK = "YYGEDNIV1GEDNIV2")');
  ExecuteSQLContOnExcept('UPDATE DPDOCUMENT SET DPD_CODEGED = SUBSTRING(DPD_CODEGED, 1, 1)||"0"||SUBSTRING(DPD_CODEGED, 2, 1) WHERE (LEN(DPD_CODEGED) = 2)');
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
  ExecuteSQLContOnExcept('UPDATE DADS2SALARIES SET PD2_NTIC="", PD2_EPARGNERETRAI=0');
  ExecuteSQLContOnExcept('UPDATE DADS2HONORAIRES SET PDH_NTIC=""');
  ExecuteSQLContOnExcept('UPDATE DEPORTSAL SET PSE_BTPPOSITION="", PSE_BTPECHELON="", PSE_BTPCATEGORIE="", PSE_BTPAFFILIRCIP=""');
  ExecuteSQLContOnExcept('UPDATE DADS2HONORAIRES SET PDH_NTIC=""');
  ExecuteSQLContOnExcept('UPDATE REMUNERATION SET PRM_SAISIEARRETAC="-" WHERE PRM_PREDEFINI<>"CEG"');
End;

Procedure MajVer648;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  ExecuteSQLContOnExcept('UPDATE DOSSIER SET DOS_PWDGLOBAL="-", DOS_NETEXPERTGED="-"');
  if not IsDossierPCL then
  begin
    //GRC
    ExecuteSQLContOnExcept('UPDATE ACTIONS SET RAC_AFFAIRE="",RAC_AFFAIRE0="",RAC_AFFAIRE1="",RAC_AFFAIRE2="",RAC_AFFAIRE3="",RAC_AVENANT=""');
    ExecuteSQLContOnExcept('UPDATE ACTIONSCHAINEES SET RCH_AFFAIRE="",RCH_AFFAIRE0="",RCH_AFFAIRE1="",RCH_AFFAIRE2="",RCH_AFFAIRE3="",RCH_AVENANT=""');
    ExecuteSQLContOnExcept('UPDATE ACTIONSGENERIQUES SET RAG_AFFAIRE="",RAG_AFFAIRE0="",RAG_AFFAIRE1="",RAG_AFFAIRE2="",RAG_AFFAIRE3="",RAG_AVENANT=""');
    //GPAO
    ExecuteSQLContOnExcept('UPDATE LIGNECOMPL SET GLC_PHASE = "" WHERE GLC_PHASE IS NULL');
    ExecuteSQLContOnExcept('UPDATE LISTEINVENT SET GIE_TYPEINVENTAIRE="INV" WHERE GIE_TYPEINVENTAIRE IS NULL');
    ExecuteSQLContOnExcept('UPDATE LISTEINVLIG SET GIL_STATUTDISPO="",GIL_STATUTFLUX="",GIL_LOTINTERNE=""'
       +',GIL_SERIEINTERNE="",GIL_TIERSPROP="",GIL_INDICEARTICLE="",GIL_MARQUE="", GIL_CHOIXQUALITE=""'
       +', GIL_REFAFFECTATION="",GIL_NUMEROSERIE=(SELECT GA_NUMEROSERIE FROM ARTICLE'
       +'  WHERE GIL_ARTICLE=GA_ARTICLE) WHERE GIL_REFAFFECTATION IS NULL');
    ExecuteSQLContOnExcept('DELETE FROM CHOIXCOD WHERE CC_TYPE="YTN" AND CC_CODE="P"');
    InsertChoixCode ('YTN','B','Des poids bruts','Poids bruts','');
    InsertChoixCode ('YTN','N','Des poids nets','Poids nets','');
  end;  // not IsDossierPCL
    // compta
  ExecuteSQLContOnExcept('UPDATE TIERSCOMPL SET YTC_PROFESSION="",YTC_REMUNERATION="",YTC_INDEMNITE="",YTC_AVANTAGE="",YTC_DAS2="",ytc_accelerateur="",ytc_schemagen=""');
End;


Procedure MajVer649;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  ExecuteSql ('Update ressource set ars_saisieact="-"');
  //GC
  ExecuteSql ('UPDATE PAYS SET PY_LIMITROPHE = "-"');
  //PAIE
  ExecuteSQLContOnExcept('UPDATE SALARIES SET PSA_NBREACQUISCP=0 WHERE PSA_NBREACQUISCP is null');
  ExecuteSQLContOnExcept('UPDATE SALARIES SET PSA_DADSNTIC="-"');
  ExecuteSQLContOnExcept('UPDATE COTISATION SET PCT_DADSEPARGNE="-"');
  ExecuteSQLContOnExcept('UPDATE INVESTFORMATION SET PIF_MTREALISEOPCA=0');
  // ajout 15092004
  AglNettoieListes('PGMULSALARIE', 'PSA_ETABLISSEMENT',nil);
  //compta
  ExecuteSQLContOnExcept('update journal set j_tresoimport="-",j_tresochainage="-",j_tresomontant="-",j_tresodate="-",j_tresolibelle="-"');
End;


Procedure MajVer650;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GPAO
    ExecuteSQLContOnExcept('UPDATE WORDREPHASE SET WOP_TRAITEMENT = "" WHERE WOP_TRAITEMENT IS NULL');
    ExecuteSQLContOnExcept('UPDATE WPDRTET SET WPE_VALEURPDR=0 WHERE WPE_VALEURPDR IS NULL');
    ExecuteSQLContOnExcept('UPDATE STKVALOPARAM SET GVP_NATURETRAVAIL="" WHERE GVP_NATURETRAVAIL IS NULL');
    ExecuteSQLContOnExcept('UPDATE STKVALOPARAM SET GVP_STKFLUX="" WHERE GVP_STKFLUX IS NULL');
    ExecuteSQLContOnExcept('UPDATE WCBNPREVTET SET wpt_affaire="" WHERE wpt_affaire IS NULL');
    ExecuteSQLContOnExcept('UPDATE STKMOUVEMENT SET GSM_NUMORDRE=0 WHERE GSM_NUMORDRE IS NULL');
    { Mise à jour du champ GL_ETATSOLDE }
    UpdateDecoupeLigne('GL_ETATSOLDE = iif(GL_ARTICLE = "", "", iif(GL_VIVANTE = "X", iif(GL_QTERESTE > 0, "ENC", "SOL"), "SOL"))');
    ExecuteSQLContOnExcept('UPDATE WORDRELIG SET WOL_MISEENPROD="" WHERE WOL_MISEENPROD is Null');
    ExecuteSQLContOnExcept('UPDATE WORDRELIG SET WOL_MISEENPROD=WOL_LIBREWOLA WHERE WOL_LIBREWOLA IN ("DEC", "ALL")');
    ExecuteSQLContOnExcept('UPDATE WORDRELIG SET WOL_LIBREWOLA="" WHERE WOL_MISEENPROD IN ("DEC", "ALL")');
    ExecuteSQLContOnExcept('UPDATE LIGNECOMPL '
             + 'SET GLC_MISEENPROD='
             + '(select isnull(WAN_MISEENPROD, "NON") FROM WARTNAT, LIGNE '
             + '  WHERE GL_ARTICLE=WAN_ARTICLE AND WAN_NATURETRAVAIL="FAB" AND GLC_NATUREPIECEG=GL_NATUREPIECEG '
             + '        AND GLC_SOUCHE=GL_SOUCHE AND GLC_NUMERO=GL_NUMERO AND GLC_INDICEG=GL_INDICEG '
             + '        AND GLC_NUMORDRE=GL_NUMORDRE)');       // modifié 17/08/2004 PCS

    ExecuteSQLContOnExcept('UPDATE WNOMELIG SET WNL_TYPELIEN="COM"');
    ExecuteSQLContOnExcept('UPDATE WORDREBES SET WOB_TYPELIEN="COM"');
    // déjà en 644 02/09/2004 GPMajSTKNature;
    //GRF
    ExecuteSQLContOnExcept('UPDATE PARACTIONS SET RPA_PRODUITPGI="GRC",RPA_TABLELIBREF1="",RPA_TABLELIBREF2="",RPA_TABLELIBREF3=""');
    ExecuteSQLContOnExcept('UPDATE ACTIONS SET RAC_PRODUITPGI="GRC",RAC_TABLELIBREF1="",RAC_TABLELIBREF2="",RAC_TABLELIBREF3=""');
    ExecuteSQLContOnExcept('UPDATE ACTIONSCHAINEES SET RCH_PRODUITPGI="GRC",RCH_TABLELIBRECHF1="",RCH_TABLELIBRECHF2="",RCH_TABLELIBRECHF3=""');
    ExecuteSQLContOnExcept('UPDATE PARCHAINAGES SET RPG_PRODUITPGI="GRC",RPG_TABLELIBRECHF1="",RPG_TABLELIBRECHF2="",RPG_TABLELIBRECHF3=""');
    ExecuteSQLContOnExcept('UPDATE ACTIONSGENERIQUES SET RAG_PRODUITPGI="GRC",RAG_TABLELIBREF1="",RAG_TABLELIBREF2="",RAG_TABLELIBREF3=""');
    ExecuteSQLContOnExcept('UPDATE CHAINAGEPIECES SET RLC_PRODUITPGI="GRC"');
    ExecuteSQLContOnExcept('UPDATE OPERATIONS SET ROP_PRODUITPGI="GRC",ROP_OBJETOPEF=""');
    ExecuteSQLContOnExcept('UPDATE PROSPECTCONF SET RTC_PRODUITPGI="GRC"');

    //GIGA
    ExecuteSql ('Update ACTIVITE set ACT_NUMORDREACH=0');
    // pour forcer les natures de piéces affaire paramsoc (pbm vielles bases)
    if not (ExisteSQL('SELECT AFF_AFFAIRE0 FROM AFFAIRE')) then
    begin
      ExecuteSQLContOnExcept('UPDATE paramsoc  SET SOC_DATA="AFF" where soc_nom = "so_afnataffaire"');
      ExecuteSQLContOnExcept('UPDATE paramsoc  SET SOC_DATA="PAF" where soc_nom = "so_afnatproposition"');
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
  ExecuteSQLContOnExcept('UPDATE DPFISCAL SET DFI_TAXEPROF = "X", DFI_CONTREVENUSLOC = "-"'
                     +', DFI_TAXEFONCIERE = "-", DFI_CONTSOCSOLDOC = "-"'
                     +',DFI_TAXEGRDSURF = "-", DFI_TAXEANNIMM = "-", DFI_TAXEVEHICSOC = "-"'
                     +', DFI_VIGNETTEAUTO = "-",DFI_IMPSOLFORTUNE = "-"');
  ExecuteSQLContOnExcept('UPDATE DPSOCIAL SET DSO_TAXESALARIE = "-", DSO_TXSALPERIODIC = "-", DSO_TAXEAPPRENT = "-"'
      +', DSO_PARTFORMCONT = "-",DSO_PARTCONSTRUC = "-", DSO_PAIECAB = "-", DSO_PAIEENT = "-"'
      +', DSO_PAIEENTSYS = "-",DSO_REGPERS = "-", DSO_DECUNEMB = "-", DSO_ELTVARIAENT = "-"'
      +', DSO_DATEDERPAIE = "'+UsDateTime(iDate1900)+'",DSO_GESTCONGES = "-", DSO_MUTSOCAGR = "-", DSO_INTERMSPEC = "-"'
      +', DSO_BTP = "-",DSO_TICKETREST = "-", DSO_GESTIONETS = "-", DSO_TELEDADS = "-", DSO_PLANPAIEACT = "-"'
      +', DSO_CDD = "-", DSO_CDI = "-", DSO_ABATFRAISPRO = "-", DSO_TPSPARTIEL = "-"'
      +', DSO_TPSPARTIEL30 = "-", DSO_CIE = "-", DSO_CEC = "-", DSO_CES = "-",DSO_CRE = "-", DSO_EMBSAL1 = "-"'
      +', DSO_EMBSAL23 = "-", DSO_CONTAPPRENT = "-",DSO_CONTQUAL = "-", DSO_CONTORIENT = "-"'
      +', DSO_EXOCHARGES = "-", DSO_DATEEXSOC = "'+UsDateTime(iDate1900)+'",DSO_COMITEENT = "-", DSO_DELEGUEPERS = "-"'
      +', DSO_DELEGUESYND = "-", DSO_EXISTTICKREST = "-",DSO_TAUXACCTRAV = 0, DSO_VERSETRANS = "-"');


  // DESGOUTTE 23092004  ExecuteSQLContOnExcept('update DPDOCUMENT set DPD_NEPUBLIABLE="X", DPD_NEAPUBLIER="-", DPD_CLAID=0, DPD_NEDATEPUBL="", DPD_NEPUBLIEUR=""');
  // DESGOUTTE 23092004   DOS_CPLIENGAMME="", DOS_CPRDDATERECEP="'+UsDateTime(iDate1900)+'", DOS_CPRDREPERTOIRE=""
  ExecuteSQLContOnExcept('update DOSSIER set DOS_NECDKEY=""');  // , DOS_CPPOINTAGESX="AUC" supprimé en 662
  // DESGOUTTE 23092004 ExecuteSQLContOnExcept('update YGEDDICO set YGD_CLAID=0');
  // remontée de paramsoc compta dans la base commune
 // DESGOUTTE 23092004
 { ExecuteSQLContOnExcept('update ##DP##.DOSSIER set DOS_CPLIENGAMME="' + GetParamSoc('SO_CPLIENGAMME') + '",'
   +' DOS_CPRDREPERTOIRE="' + GetParamSoc('SO_CPRDREPERTOIRE') + '",'
   +' DOS_CPRDDATERECEP="' +  UsDateTime(v)  + '"'
   +' WHERE DOS_NODOSSIER="'+V_PGI.NoDossier+'"');
 }
  ExecuteSQLContOnExcept('UPDATE DPTABCOMPTA SET DTC_NBENTREEIMMO=0, DTC_NBSORTIEIMMO=0, DTC_NBLIGNEIMMO=0');
  //GIGA
  ExecuteSQLContOnExcept('update paramsoc set soc_data=(select soc_data from paramsoc where soc_nom="so_afvaloactpr") where soc_nom = "so_afvalofraispr"');
  ExecuteSQLContOnExcept('update paramsoc set soc_data=(select soc_data from paramsoc where soc_nom="so_afvaloactpr") where soc_nom = "so_afvalofourpr"');
  ExecuteSQLContOnExcept('update paramsoc set soc_data=(select soc_data from paramsoc where soc_nom="so_afvalopv") where soc_nom = "so_afvalofourpv"');
  ExecuteSQLContOnExcept('update paramsoc set soc_data=(select soc_data from paramsoc where soc_nom="so_afvalopv") where soc_nom = "so_afvalofraispv"');

  if not IsDossierPCL then
  begin
    //GIGA
    ExecuteSQLContOnExcept('update afcumul set acu_errcpta=0, acu_familletaxe="",acu_regimetaxe="",acu_etablissement=""');

    if Vers_Piece < 164 then  UpdateDecoupePiece('GP_DATEDEBUTFAC =GP_DATEPIECE,GP_DATEFINFAC =GP_DATEPIECE');
    //GESCOM
    ExecuteSQLContOnExcept('UPDATE ACOMPTES SET GAC_MONTANTTRA=0, GAC_MONTANTTRADEV=0, GAC_PIECEPRECEDENTE="" ');
    ExecuteSQLContOnExcept('update article set ga_calculpa="-" ');
  end;  // not IsDossierPCL
  //compta
  ExecuteSQLContOnExcept('update IMMOLOG SET IL_DATEOPREELLE=IL_DATEOP ');
  ExecuteSQLContOnExcept('UPDATE GENERAUX SET G_IAS14 = "-" ');
  AglNettoieListes('CPMODIFENTPIE', 'E_NUMGROUPEECR', nil);
End;

Procedure MajVer652;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //Compta
  ExecuteSQLContOnExcept('UPDATE IMMO SET I_INFOUO=""');
  if not IsDossierPCL then
  begin
    //GC
    ExecuteSQLContOnExcept('UPDATE OPERCAISSE SET GOC_NUMZCAISSE=0, GOC_CAISSE="", GOC_ANTERIEUR="-"');
    //GP
    ExecuteSQLContOnExcept('update QSATTRIB set qsb_affaire="", qsb_ordfabpere="" where qsb_affaire is null');
    ExecuteSQLContOnExcept('update QSTKDISPO set qsd_affaire="", qsd_ordfabpere="" where qsd_affaire is null');
    ExecuteSQLContOnExcept('update QHISTOAFF set qha_affaire="", qha_ordfabpere="" where qha_affaire is null');
    ExecuteSQLContOnExcept('update WCBNEVOLUTION set wev_ordfabpere="" where wev_ordfabpere is null');


  //GC   (champs créés en 639)
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="SAC", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="AFS"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="EVE", GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG="AVS"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="SAC", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="BFA"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="SVE", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="BLC"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="EAC", GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG="BLF"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="EVE", GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG="BRC"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="EAC", GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG="BSA"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="RVE", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="CC"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="RVE", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="CCE"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="AAC", GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG="CF"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="AAC", GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG="CSA"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="SVE", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="FAC"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="AAC", GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG="FCF"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="EAC", GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG="FF"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="SVE", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="FFO"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="SVE", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="LCE"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="RPL", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="PRE"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="STR", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="TEM"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="ETR", GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG="TRE"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="TAV", GPP_SENSPIECE="MIX"  WHERE GPP_NATUREPIECEG="TRV"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="SAC", GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG="AFP"');
  end;  // not IsDossierPCL

  //DP
  // DESGOUTTE 23092004  ExecuteSQLContOnExcept('UPDATE DOSSIER SET DOS_NEDOSACREER="-", DOS_NEGEDACREER="-"');
End;

Procedure MajVer653;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //DP
  ExecuteSQLContOnExcept('UPDATE DOSSIER SET DOS_PWDGLOBAL="X" WHERE DOS_APPLISPROTEC=""');
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
    ExecuteSQLContOnExcept('DELETE FROM STOXEVENTS WHERE SEV_CODEEVENT = "PCP_DEMANDE"');
    ExecuteSQLContOnExcept('DELETE FROM STOXEVENTS WHERE SEV_CODEEVENT = "PCP_GENERIQUE"');
    ExecuteSQLContOnExcept('DELETE FROM STOXEVENTS WHERE SEV_CODEEVENT = "PCP_INIT_ACH"');
    ExecuteSQLContOnExcept('DELETE FROM STOXEVENTS WHERE SEV_CODEEVENT = "PCP_INIT_VTE"');
    ExecuteSQLContOnExcept('DELETE FROM STOXEVENTS WHERE SEV_CODEEVENT = "PCP_MOD_ACH"');
    ExecuteSQLContOnExcept('DELETE FROM STOXEVENTS WHERE SEV_CODEEVENT = "PCP_MOD_VTE"');
    ExecuteSQLContOnExcept('DELETE FROM STOXEVENTS WHERE SEV_CODEEVENT = "PCP_REMONTEE_ACH"');
    ExecuteSQLContOnExcept('DELETE FROM STOXEVENTS WHERE SEV_CODEEVENT = "PCP_REMONTEE_VTE"');
    MajEvtPCP_V654;
    // Fin JTR
    (*
    // GC initialisation des nouveaux stocks
    if (V_PGI.driver <> dbORACLE7) then
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
    *)
  end;  // not IsDossierPCL

End;

Procedure MajVer655 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSQLContOnExcept('UPDATE ARTICLE SET GA_UNITEQTEVTE = IIF(GA_UNITEQTEVTE = "", GA_QUALIFUNITESTO, GA_UNITEQTEVTE) '
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
  ExecuteSQLContOnExcept('delete from CHOIXEXT WHERE YX_TYPE like "LB%" and len(yx_code) > 3');
  //GI
  AglNettoieListes('AFMULLIGNE1', 'GL_NUMORDRE', nil);
End;

Procedure MajVer657 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //GRC
  ExecuteSQLContOnExcept('UPDATE CONTACT SET C_CLETELEPHONE="" where C_CLETELEPHONE is null');
  if not IsDossierPCL then
  begin
    //GP
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_REFINTEXT="INT", GPP_MAJINFOTIERS="X" WHERE GPP_NATUREPIECEG="BSP"');
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
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_ACTIONFINI="TRA",GPP_TAXE="X",GPP_NATURETIERS="FOU;",GPP_APPELPRIX="PAS",GPP_TYPEARTICLE="MAR;"'
       +  ',GPP_MAJPRIXVALO="DPA;DPR;PPA;PPR;",GPP_PRIORECHART1="ART",GPP_DIMSAISIE="TOU",GPP_IFL1="004",GPP_IFL2="002",'
       +  'GPP_IFL3="011", GPP_GEREECHEANCE="DEM", GPP_COMPANALLIGNE="SAN",GPP_COMPANALPIED="SAN" WHERE GPP_NATUREPIECEG = "BSP"');

    InsertChoixCode('WMA','AFF', 'A l''affaire',' A l''affaire','');
    //ExecuteSQLContOnExcept(‘UPDATE RESSOURCE SET ARS_DEPOT= “‘+VH_GC.GCDepotDefaut+’“ WHERE ARS_DEPOT=  “” ‘);
    ExecuteSQLContOnExcept('UPDATE RESSOURCE SET ARS_DEPOT=(Select SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_GCDEPOTDEFAUT") WHERE ARS_DEPOT=""');
    W_SAV_COMPETENCE_UPDATE;
    W_SAV_COMPETENCE;
    W_SAV_INTER_COMPL;
    //--------- WCBNTYPEMOUV -----------------------
    if not ExisteSQL('select 1 from wcbntypemouv where wtm_typemouvement = "PLC"') then
    begin
      ExecuteSQLContOnExcept('insert into WCBNTYPEMOUV (WTM_TYPEMOUVEMENT,WTM_LIBELLE) values ("PLC","Plan commercial")');
    end;

    if not ExisteSQL('select 1 from wcbntypemouv where wtm_typemouvement = "DEV"') then
    begin
      ExecuteSQLContOnExcept('insert into WCBNTYPEMOUV (WTM_TYPEMOUVEMENT,WTM_LIBELLE) values ("DEV","Devis")');
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
    ExecuteSQLContOnExcept('UPDATE DISPO SET GQ_UNITEMIN="1",GQ_UNITEMAX="1",GQ_UNITEALERTE="1",GQ_UNITERECOMPL="1" WHERE (GQ_UNITEMIN IS NULL) OR (GQ_UNITEMIN ="")');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_SOUCHE="BSP",GPP_LISTESAISIE="GCSAISIEBSP",GPP_RELIQUAT="X" WHERE GPP_NATUREPIECEG = "BSP"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_SOUCHE="CSP",GPP_LISTESAISIE="GCSAISIECSP",GPP_RELIQUAT="X" WHERE GPP_NATUREPIECEG = "CSP"');
    UpdateDecoupeLigne('gl_typenomenc="ASC"',
         ' AND gl_typearticle="NOM" and gl_typenomenc="ASS" and gl_tenuestock="-" and gl_naturepieceg not in ("FF", "BLF", "CF", "AF", "AFP", "AFS", "BFA", "BSA", "BSP", "CFR", "CSA", "CSP", "DEF", "EEX", "SEX", "FCF", "FRF", "LFR", "REA", "TEM", "TRE")');
    //GC
    AglNettoieListes('LSTOXQUERYS', 'SQE_TYPETRT',nil);
    AglNettoieListes('LSTOXCHRONO', 'SXC_CODESITE',nil);
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="" WHERE GPP_NATUREPIECEG = "BSP"');
  end;  // not IsDossierPCL
End;

Procedure MajVer660 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    // GPAO
    ExecuteSQLContOnExcept('Update stknature set GSN_SIGNEMVT = "SOR" where gsn_StkTypeMvt = "RES"');
    ExecuteSQLContOnExcept('update WCBNTYPEMOUV set WTM_LIBELLE="Besoin / vente"  where WTM_TYPEMOUVEMENT="VEN"');
    ExecuteSQLContOnExcept('update DISPO set GQ_UNITEMIN="1",GQ_UNITEMAX="1",GQ_UNITEALERTE="1",GQ_UNITERECOMPL="1" where (GQ_UNITEMIN is null) or (GQ_UNITEMIN ="") or (GQ_UNITEMIN ="QTE")');
    ExecuteSQLContOnExcept('update PARPIECE set GPP_PILOTEORDRE="X", GPP_IMPBESOIN="X", GPP_IMPAUTOETATCBN="-"  where GPP_NATUREPIECEG="CSA"');
    ExecuteSQLContOnExcept('update PARPIECE set GPP_PILOTEORDRE="X", GPP_IMPBESOIN="-", GPP_IMPAUTOETATCBN="-"  where GPP_NATUREPIECEG="BSA"');

    { Initialisation de la table des types de prix de revient utilisée notamment dans STKVALOTYPE }
    GPMajPdrType;
    { Initialisation des deux tables STKVALOTYPE et STKVALOPARAM pour un fonctionnement standard }
    GPMajSTKValoParam;
    { Réinitialisation des N° de jetons notamment pour les tables STKVALOTYPE et STKVALOPARAM }
(*    wResetWJT; *)
  end;  // not IsDossierPCL
  // PAie
  AglNettoieListes('PGEMULMVTABS', 'PCN_SALARIE;PCN_ORDRE;PCN_TYPEMVT',nil);
  AglNettoieListes('PGEMULMVTABSR', 'PCN_SALARIE;PCN_ORDRE;PCN_TYPEMVT',nil);
  //DP
  ExecuteSQLContOnExcept('update YGEDDICO set YGD_CDPUBLI="", YGD_EWSID=""');
  ExecuteSQLContOnExcept('update DPDOCUMENT set DPD_EWSID="", DPD_EWSPUBLIABLE="X", DPD_EWSAPUBLIER="-", DPD_EWSDATEPUBL="'+UsDateTime(iDate1900)+'", DPD_EWSPUBLIEUR=""');
  ExecuteSQLContOnExcept('update DOSSIER set DOS_USRS1="", DOS_PWDS1="", DOS_EWSCREE="-"');
	// supprimé le 30092004 ExecuteSQLContOnExcept('update JQ6_fonction set jq6_formeq = ""');
	ExecuteSQLContOnExcept('update histoannulien set hnl_typemodif = ""');

  // MB : Table USERCONF renommée en OLDUSERCONF et supprimée ensuite mais remplacée par une VIEW.
  if IsMonoOuCommune Then
    If TableExiste('OLDUSERCONF') then  //Dans le cas d'un PGIMajVer V8
    Begin
      ExecuteSQLContOnExcept('insert into OLDUSERCONF(uco_groupeconf, uco_user) select cc_code, us_utilisateur '
                            +' from choixcod, utilisat where cc_type="UCO" and us_superviseur="X" '
                            +' and not exists(select 1 from OLDUSERCONF where uco_user=us_utilisateur '
                            +' and uco_groupeconf=cc_code) group by cc_code, us_utilisateur');
    End
    Else                                //Dans le cas d'un PGIMajVer V7
    Begin
    	If TableExiste('USERCONF') then  //Dans le cas d'un PGIMajVer V8
      begin

      ExecuteSQLContOnExcept('insert into USERCONF(uco_groupeconf, uco_user) select cc_code, us_utilisateur '
                            +' from choixcod, utilisat where cc_type="UCO" and us_superviseur="X" '
                            +' and not exists(select 1 from USERCONF where uco_user=us_utilisateur '
                            +' and uco_groupeconf=cc_code) group by cc_code, us_utilisateur');
    	end;
    End;
End;

Procedure MajVer661 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GPAO
     ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_QTEPLUS = "GQ_PHYSIQUE;" WHERE GSN_QUALIFMVT = "SEX"') ;
  end;  // not IsDossierPCL
  //TRESO
   // AglNettoieListes('TRECRITURERAPPRO', 'TE_DATECOMPTABLE;TE_NUMTRANSAC;TE_NUMEROPIECE;TE_CODEFLUX;TE_CODEBANQUE;TE_EXERCICE;BQ_LIBELLE;TE_JOURNAL;TE_MONTANT;TE_DEVISE;TE_MONTANTDEV;TE_DATERAPPRO;TE_DATEVALEUR;TE_NUMLIGNE;TE_CLEVALEUR;TE_NATURE;TE_GENERAL;TE_CPNUMLIGNE;TE_NUMECHE;');
   ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA = "-" WHERE SOC_NOM = "SO_PREMIERESYNCHRO" ');
   //COMPTA
   AglNettoieListes('MULMIMMOS', 'I_GROUPEIMMO',nil);
   ExecuteSQLContOnExcept('update journal set j_accelerateur="X" where j_naturejal="ACH" or j_naturejal="VTE" ');
   //DP
   ExecuteSQLContOnExcept('UPDATE DOSSIER SET DOS_SERIAS1="X"');
End;

Procedure MajVer662 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GC
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_ARTSTOCK="-" WHERE GPP_NATUREPIECEG="BFA"');
    //GPAO
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_LIBELLE="Commande SST d''achat" WHERE GPP_NATUREPIECEG="CSA"') ;
    if not ExisteSQL('SELECT 1 FROM WCBNTYPEMOUV WHERE WTM_TYPEMOUVEMENT = "BOS"') then
      ExecuteSQLContOnExcept('INSERT INTO WCBNTYPEMOUV (WTM_TYPEMOUVEMENT,WTM_LIBELLE) VALUES ("BOS","Besoin / ordre de sst d''achat")');
    if not ExisteSQL('SELECT 1 FROM WCBNTYPEMOUV WHERE WTM_TYPEMOUVEMENT = "BPS"') then
      ExecuteSQLContOnExcept('INSERT INTO WCBNTYPEMOUV (WTM_TYPEMOUVEMENT,WTM_LIBELLE) VALUES ("BPS","Besoin / proposition de sst d''achat")');
    ExecuteSQLContOnExcept('UPDATE WCBNTYPEMOUV SET WTM_LIBELLE="Proposition de sst d''achat" WHERE WTM_TYPEMOUVEMENT="PRS"');
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
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_LISTESAISIE="AFSAISIEFAC" WHERE GPP_NATUREPIECEG="AFF"  AND  GPP_LISTESAISIE="GCSAISIECC"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_LISTESAISIE="AFSAISIEFAC" WHERE GPP_NATUREPIECEG="PAF"  AND  GPP_LISTESAISIE="GCSAISIEDEC"');
    AglNettoieListes('AFMULRECHAFFAIRE', 'AFF_TIERS',nil);
    //  Ajout pour 6.01
    ExecuteSQLContOnExcept('Update stknature set gsn_signemvt="MIX" where gsn_qualifmvt in ("EVE", "SAC")');
  end;
    // nouveaux concepts pour présentations et filtres initialisés aux mêmes droits que les concepts existants
  ExecuteSQLContOnExcept('update menu set mn_ACCESGRP=(select MAX(mn_accesgrp) from menu MN where MN.mn_tag=26000 and MN.mn_1 =26) '
    +' where (mn_1 = 26) and mn_tag = 26004 ');
  ExecuteSQLContOnExcept('update menu set mn_ACCESGRP=(select MAX(mn_accesgrp) from menu MN where MN.mn_tag=26001 and MN.mn_1 =26) '
    +' where (mn_1 = 26) and mn_tag = 26003 ');
   //COMPTA
   ExecuteSQLContOnExcept('UPDATE BANQUES SET PQ_ETABBQ="" where PQ_ETABBQ is null');
   ExecuteSQLContOnExcept('UPDATE BANQUES SET PQ_ABREGE="" where PQ_ABREGE is null');
   AglNettoieListes('CPEEXBQ','EE_ORIGINERELEVE;EE_NUMERO');
   AglNettoieListes('CPEEXBQ2','EE_ORIGINERELEVE;EE_NUMERO');
   AglNettoieListes('CPRELANCECLIENT','E_BLOCNOTE');
End;

Procedure MajVer664 ;
Begin
  ExecuteSQL('UPDATE COMMERCIAL SET GCL_HRRANG="" WHERE (GCL_HRRANG IS NULL)');
  ExecuteSQL('UPDATE NOMENLIG SET GNL_QTEMINMENU=0, GNL_QTEMAXMENU=0 , GNL_POIDSMENU=0 WHERE (GNL_QTEMINMENU IS NULL)');
  ExecuteSQL('UPDATE PARCAISSE SET GPK_FAMRES="" WHERE (GPK_FAMRES IS NULL)');
  ExecuteSQL('UPDATE ARTICLECOMPL SET GA2_DEPOT="",GA2_HRQTESTAT="-",GA2_INTERDITACHAT="-" WHERE (GA2_DEPOT IS NULL)');
  ExecuteSQL('UPDATE HRGROUPESSTAT SET HGS_LFAMRES1="",HGS_LFAMRES2="",HGS_LFAMRES3="",HGS_LFAMRES4="",HGS_LFAMRES5="",'+
             'HGS_LFAMRES6="", HGS_LIBELLECHAMP1="",HGS_LIBELLECHAMP2="",HGS_LIBELLECHAMP3="",HGS_LIBELLECHAMP4="",HGS_LIBELLECHAMP5="",HGS_LIBELLECHAMP6=""  WHERE HGS_LFAMRES1 IS NULL');
  ExecuteSQL('UPDATE HRDOSRES SET HDR_CODEFACT="", HDR_NOMBREPERSOFF=0 WHERE HDR_CODEFACT IS NULL');
  ExecuteSQL('UPDATE HRDOSRES SET HDR_LIBELLERESA1="",HDR_LIBELLERESA2="" WHERE HDR_LIBELLERESA1 IS NULL');
  ExecuteSQL('UPDATE HRDOSRES SET HDR_TYPRESASSOC="" WHERE HDR_TYPRESASSOC IS NULL');
  ExecuteSQL('UPDATE HRALLOTEMENT SET HAL_TYPRESASSOC="" WHERE HAL_TYPRESASSOC IS NULL');
  ExecuteSQL('UPDATE HRLIGNE SET HRL_ORDREFACTURE=0 WHERE HRL_ORDREFACTURE IS NULL');
  ExecuteSQL('UPDATE HRPARAMPLANNING SET HPP_VISULIGNETRF="-" WHERE HPP_VISULIGNETRF IS NULL');
  ExecuteSQL('UPDATE HRDOSSIER SET HDC_DOSSIMPLIFIE="-",HDC_HORSCONTRAT="-",HDC_HORSALLOTEMENT="-"  WHERE HDC_DOSSIMPLIFIE IS NULL');
  ExecuteSQL('UPDATE HRDOSSIER SET HDC_FAMILLETAXE1="" WHERE HDC_FAMILLETAXE1 IS NULL');
end;

Procedure MajVer665 ;
Begin
  ExecuteSQL('UPDATE PARPIECECOMPL SET GPC_MAQUETTEDOC=""');
  ExecuteSQL('UPDATE AFFAIRE SET AFF_PRIOCONTRAT=""');
  ExecuteSQL('UPDATE LIENDEVCHA SET BDA_QUANTITE=0');
  SetParamSoc('SO_AFPROFILGENER', 'PAF');
end;

Procedure MajVer667 ;
var iInd : integer;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GC
    for iInd := 1 to 8 do
      ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_IFL' + intToStr(iInd) + '=IIF(GPP_IFL'
        + intToStr(iInd) + '="411","015",IIF(GPP_IFL' + intToStr(iInd) + '="400","001",IIF(GPP_IFL'
        + intToStr(iInd) + '="409","002",IIF(GPP_IFL'
        + intToStr(iInd) + '="404","010",GPP_IFL' + intToStr(iInd) + '))))');
    //GPAO
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_IMPMODELE = "", GPP_IMPETAT= "BSP" WHERE GPP_NATUREPIECEG="BSP"');
    ExecuteSQLContOnExcept('UPDATE EDILIGNE SET ELI_PUTTCDEV=0,ELI_REMISELIBRE=0,ELI_REMISELIGNE=0,'
             + ' ELI_REMISEPIED=0,ELI_VALEURFIXEDEV=0,ELI_VALEURREMDEV=0'
             + ' WHERE ELI_PUTTCDEV is Null AND ELI_REMISELIBRE is Null'
             + ' AND ELI_REMISELIGNE is Null AND ELI_REMISEPIED is Null'
             + ' AND ELI_VALEURFIXEDEV is Null AND ELI_VALEURREMDEV is Null');
    if Not ExisteSQL('SELECT WDB_QUALIFIANT FROM WCUQUALIFIANT WHERE WDB_QUALIFIANT="QAF"') then
        ExecuteSQLContOnExcept('INSERT INTO WCUQUALIFIANT VALUES ("QAF", "Quantité affectée", "*")');

    if Not ExisteSQL('SELECT WDX_CONTEXTE FROM WFORMULECHAMP WHERE WDX_CONTEXTE="WOB" AND WDX_QUALIFIANT="QAF"') then
        ExecuteSQLContOnExcept('INSERT INTO WFORMULECHAMP'
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
    {BTP}
    ExecuteSQL('UPDATE CODECPTA SET GCP_FAMILLETAXE="" WHERE GCP_FAMILLETAXE IS NULL');
    ExecuteSQL('UPDATE LIGNEPHASES SET BLP_SOUCHE = (SELECT GPP_SOUCHE FROM PARPIECE WHERE GPP_NATUREPIECEG=LIGNEPHASES.BLP_NATUREPIECEG) WHERE BLP_SOUCHE IS NULL');
    // maj pour bordereaux de prix et tarification fournisseurs
    ExecuteSQL('UPDATE LIGNECOMPL SET GLC_FROMBORDEREAU="-", GLC_GETCEDETAIL="-" WHERE GLC_FROMBORDEREAU IS NULL AND GLC_GETCEDETAIL IS NULL');
    ExecuteSQL('UPDATE BDETETUDE SET BDE_DATEDEPART="'+ UsDateTime(iDate1900)+'",BDE_DATEFIN="'+ UsDateTime(iDate2099)+'",BDE_NATUREAUXI="",BDE_CLIENT="",BDE_NATUREPIECEG="",BDE_SOUCHE="",BDE_NUMERO=0,BDE_INDICEG=0 WHERE BDE_DATEDEPART IS NULL');
    ExecuteSQL('UPDATE ARTICLECOMPL SET GA2_SOUSFAMTARART="" WHERE GA2_SOUSFAMTARART IS NULL');
    ExecuteSQL('UPDATE PROFILART SET GPF_SOUSFAMTARART="" WHERE GPF_SOUSFAMTARART IS NULL');
    ExecuteSQL('UPDATE TARIF SET GF_SOUSFAMTARART="" WHERE GF_SOUSFAMTARART IS NULL');
    ExecuteSql('UPDATE PARPIECE SET GPP_LISTESAISIE="BTSAISIEDOC" WHERE GPP_NATUREPIECEG="LBT"');
  End;

  //Treso
  // J.PASTERIS 17012005 ExecuteSQLContOnExcept('UPDATE TRECRITURE SET TE_COMMISSION = "S"');
  // J.PASTERIS 17012005 ExecuteSQLContOnExcept('UPDATE FLUXTRESO SET TFT_CLASSEFLUX = "FIN" WHERE TFT_PREVISIONNEL = "-" OR TFT_PREVISIONNEL = "" OR TFT_PREVISIONNEL IS NULL');
  // J.PASTERIS 17012005 ExecuteSQLContOnExcept('UPDATE FLUXTRESO SET TFT_CLASSEFLUX = "PRE" WHERE TFT_PREVISIONNEL = "X"');
  // J.PASTERIS 17012005 ExecuteSQLContOnExcept('UPDATE FLUXTRESO SET TFT_CLASSEFLUX = "REF" WHERE TFT_FLUX IN ("EQR", "EQD", "REI")');
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
    ExecuteSQLContOnExcept('UPDATE STKMOUVEMENT SET GSM_QRUPTURE = 0.0');
  end;
End;

Procedure MajVer671 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GPAO
    //SAUZET Jean-Luc Demande N° 1238 (Reportée dans la futur version V7 édition 7) 
    //ExecuteSQLContOnExcept('update stkmouvement set GSM_DPA=0.0, GSM_DPR=0.0, GSM_PMAP=0.0, GSM_PMRP=0.0, GSM_PRIXSAISIS="X" where GSM_PRIXSAISIS is null' );
    ExecuteSQLContOnExcept('UPDATE STKMOUVEMENT SET GSM_DPA=0.0, GSM_DPR=0.0, GSM_PMAP=0.0, GSM_PMRP=0.0, GSM_PRIXSAISIS="-" WHERE GSM_PRIXSAISIS IS NULL' );

    AglNettoieListes('GCSTKDISPO', 'GA_QUALIFUNITEVTE', nil);
    ExecuteSQLContOnExcept('UPDATE STKMOUVEMENT SET GSM_ETATTRANSFERT="" WHERE GSM_ETATTRANSFERT IS NULL');
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
    // js1 200606 vient du majavant 673 (pose pbe en pcl)
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_NATURESUIVANTE=GPP_NATURESUIVANTE||"PRF;" WHERE GPP_NATUREPIECEG="CF" AND GPP_NATURESUIVANTE not like "%PRF%"');
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
    ExecuteSQLContOnExcept('update TRANSINVENT set GIT_CODELISTE="" where GIT_CODELISTE is null');
    ExecuteSQLContOnExcept('UPDATE DEPOTS SET GDE_STKTRFEMPLACE="", GDE_STKTRFNOMINV="", GDE_STKTRFNOMENT="", GDE_STKTRFNOMSOR="" WHERE GDE_STKTRFEMPLACE IS NULL');
    if not ExisteSQL('select 1 from STKNATURE where GSN_QUALIFMVT="EIT"') then
    begin
      ExecuteSQLContOnExcept('INSERT INTO STKNATURE (GSN_QUALIFMVT,GSN_LIBELLE,GSN_STKTYPEMVT,GSN_QTEPLUS,GSN_QUALIFMVTSUIV,GSN_SIGNEMVT,GSN_STKFLUX,GSN_GERECOMPTA) VALUES ("EIT","Entrée d''inventaire : transtockeur","PHY","GQ_PHYSIQUE;","","ENT","STO","-")');
    end;
    if not ExisteSQL('select 1 from STKNATURE where GSN_QUALIFMVT="SIT"') then
    begin
      ExecuteSQLContOnExcept('INSERT INTO STKNATURE (GSN_QUALIFMVT,GSN_LIBELLE,GSN_STKTYPEMVT,GSN_QTEPLUS,GSN_QUALIFMVTSUIV,GSN_SIGNEMVT,GSN_STKFLUX,GSN_GERECOMPTA) VALUES ("SIT","Sortie d''inventaire : transtockeur","PHY","GQ_PHYSIQUE;","","SOR","STO","-")');
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
    ExecuteSQLContOnExcept('UPDATE WPDRTET SET WPE_NATUREPDR= "BTH", WPE_LIGNEORDRE=0 WHERE WPE_TYPEPDR="ACT" AND WPE_NATUREPDR IS NULL');
    ExecuteSQLContOnExcept('UPDATE WPDRTYPE SET WRT_NATUREPDR= "BTH" WHERE WRT_NATUREPDR IS NULL');
    ExecuteSQLContOnExcept('UPDATE WPDRTYPE SET WRT_PARTYPEPARTN= "-", WRT_PARTYPECOUTN="-", WRT_PARTYPEENTITEN="-", WRT_PARENTITEN="-", WRT_PARSECTIONN="-"'
                  +', WRT_PARRUBRIQUEN="-", WRT_PARORIGINEN="-", WRT_PAROPECIRCN="-", WRT_PARTYPEPARTG= "-"'
                  +', WRT_PARTYPECOUTG="-", WRT_PARTYPEENTITEG="-", WRT_PARENTITEG="-", WRT_PARSECTIONG="-"'
                  +', WRT_PARRUBRIQUEG="-", WRT_PARORIGINEG="-",  WRT_PAROPECIRCG="-", WRT_DEFAUTN= "X"'
                  +', WRT_DEFAUTG = "X"  WHERE WRT_PARTYPEPARTN IS NULL');
    // supprimé en 691 ExecuteSQLContOnExcept('UPDATE WORDREAUTO SET WOA_BOOLLIBRE9= "-" WHERE WOA_BOOLLIBRE9 IS NULL');
    ExecuteSQLContOnExcept('UPDATE WORDREBES SET WOB_TYPECOMPOSANT= "" WHERE WOB_TYPECOMPOSANT IS NULL');
    ExecuteSQLContOnExcept('UPDATE WORDRERES SET WOR_SECTIONPDR="", WOR_RUBRIQUEPDR="" WHERE WOR_SECTIONPDR IS NULL');
    ExecuteSQLContOnExcept('UPDATE WGAMMERES SET WGR_SECTIONPDR="", WGR_RUBRIQUEPDR="" WHERE WGR_SECTIONPDR IS NULL');
    //
    if not ExisteSQL('SELECT QDE_CIRCUIT FROM QDETCIRC WHERE QDE_CIRCUIT="NUL"') then
    begin
     ExecuteSQLContOnExcept('INSERT INTO QDETCIRC (QDE_CTX,QDE_CIRCUIT,QDE_OPECIRC,QDE_JALON,QDE_POLE,QDE_SITE,QDE_GRP,QDE_TYPTRANS,QDE_JALONMAX,QDE_JALONCOURT,QDE_TYPTRANSCOURT,QDE_DATEMODIF)'
           + ' VALUES ("0","NUL","110","","","","","","","","","'+ UsDateTime(iDate1900)  +'")');
    end
    else
    begin
     ExecuteSQLContOnExcept('update QDETCIRC set QDE_JALON="",QDE_JALONMAX="",QDE_JALONCOURT=""  where QDE_CIRCUIT="NUL"');
    end;
    ExecuteSQLContOnExcept('UPDATE TRECRITURE SET TE_COMMISSION = "S"');
    ExecuteSQLContOnExcept('UPDATE FLUXTRESO SET TFT_CLASSEFLUX = "REF" WHERE TFT_FLUX IN ("EQD", "EQR", "REI")');
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
    ExecuteSQLContOnExcept('UPDATE WPARAM SET WPA_VARCHAR08="BTH" WHERE WPA_CODEPARAM="PRIXDEREVIENT" AND WPA_VARCHAR08 IS NULL');
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
    // supprimée en 691 ExecuteSQLContOnExcept('UPDATE WORDREAUTO SET WOA_BOOLLIBREA= "-" WHERE WOA_BOOLLIBREA IS NULL');
    ExecuteSQLContOnExcept('UPDATE WPDRLIG SET WPL_TIERS="", WPL_CIRCUIT="", WPL_DEPOT="", WPL_SITE="" WHERE WPL_TIERS IS NULL');
    ExecuteSQLContOnExcept('UPDATE DEVISE SET D_PARITEEUROFIXING=D_PARITEEURO,D_ARRONDIPRIXACHAT=D_DECIMALE,D_ARRONDIPRIXVENTE=D_DECIMALE WHERE D_ARRONDIPRIXVENTE<1');
    ExecuteSQLContOnExcept('UPDATE WORDREPHASE SET WOP_VALEURFIXE=0, WOP_PRIXUNIT=0 WHERE WOP_VALEURFIXE IS NULL');
  end;
End;

Procedure MajVer686 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSQLContOnExcept('UPDATE WTRAITEMENT SET WTR_TYPETRAITSTR="", WTR_SECTIONPDR="", WTR_RUBRIQUEPDR="" WHERE WTR_TYPETRAITSTR IS NULL');

    if not ExisteSQL('SELECT WRT_TYPEPDR FROM WPDRTYPE WHERE WRT_TYPEPDR="ORR"') then
    begin
      ExecuteSQLContOnExcept('INSERT INTO WPDRTYPE '
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
      ExecuteSQLContOnExcept('INSERT INTO WPDRTYPE '
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
    ExecuteSQLContOnExcept('UPDATE ARTICLE SET GA_CLASSE1="",GA_CLASSE2="",GA_CLASSE3="",GA_CLASSE4="",GA_CLASSE5=""');


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
    ExecuteSQLContOnExcept('UPDATE EMPLACEMENT SET GEM_DATECREATION="' + UsdateTime(iDate1900) + '", GEM_DATEMODIF="' + UsdateTime(iDate1900) + '", GEM_CREATEUR="", GEM_UTILISATEUR="", GEM_STATUTDISPO="", GEM_STATUTFLUX=""');
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
    ExecuteSQLContOnExcept('UPDATE LIGNECOMPL SET GLC_INDICEARTICLE=""');
    ExecuteSQLContOnExcept('UPDATE WNOMEDEC SET WND_INDICECOMPOSAN=""');
    ExecuteSQLContOnExcept('UPDATE WNOMELIG SET WNL_INDICECOMPOSAN=""');
    ExecuteSQLContOnExcept('UPDATE WORDREBES SET WOB_INDICECOMPOSAN=""');
    ExecuteSQLContOnExcept('UPDATE WORDRELIG SET WOL_INDICEARTICLE=""');
    ExecuteSQLContOnExcept('UPDATE WGAMMELIG SET WGL_TYPEEMPLOI=""');
    ExecuteSQLContOnExcept('UPDATE WORDREBES SET WOB_TYPEEMPLOI="", WOB_INDICECOMPOSAN=""');
    ExecuteSQLContOnExcept('UPDATE WORDREGAMME SET WOG_TYPEEMPLOI=""');
    //GC
    ExecuteSQLContOnExcept('UPDATE PARAMCLASSEABC SET GPM_NATUREPIECEG=ISNULL(GPM_NATUREPIECEG,"") '
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
 ExecuteSQLContOnExcept('update generaux set g_cutoff="-",g_cutoffperiode="",g_cutoffechue="-",g_visarevision="-", g_totdebn2=0,g_totcren2=0,g_cyclerevision=""');
 ExecuteSQLContOnExcept('update generauxref set ger_cutoff="-",ger_cutoffperiode="",ger_cutoffechue="-"');
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
    UpdateDecoupeLigne('GL_TYPECADENCE = IIF( ((ISNULL(GL_TYPECADENCE, "") = "") AND (ISNULL(GL_ARTICLE, "") <> "")), "FER", "")'
                          + ', GL_REFTRANSPORT=""'
                          + ', GL_DATEENLEVEMENT = "' + UsDateTime(IDate1900) + '"');

    ExecuteSQLContOnExcept('UPDATE WNOMELIG SET WNL_INDICECOMPOSAN=""');
    ExecuteSQLContOnExcept('UPDATE WORDRELIG SET WOL_ORDREPERE=0, WOL_OPECIRCWOB="", WOL_LIENNOMEWOB=0');
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
    ExecuteSQLContOnExcept('UPDATE COMPTADIFFEREE SET GCD_TIERSPAYEUR=""');
    //GRC
    ExecuteSQLContOnExcept('UPDATE PARCHAINAGES SET RPG_NONTERMINE="-"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_CHAINAGE=""');
    ExecuteSQLContOnExcept('UPDATE DOMAINEPIECE SET GDP_CHAINAGE=""');
    // GPAO
    // supprimé 691 ExecuteSQLContOnExcept('UPDATE TIERS SET T_ESTTRANSPORTEUR="-", T_SECTEURGEO ="", T_TIMBRE=0,  T_SURTAXE=0  WHERE T_ESTTRANSPORTEUR IS NULL');
    ExecuteSQLContOnExcept('UPDATE ADRESSES SET ADR_SECTEURGEO="" WHERE ADR_SECTEURGEO IS NULL');
    ExecuteSQLContOnExcept('UPDATE PIECEADRESSE SET GPA_SECTEURGEO="" WHERE GPA_SECTEURGEO IS NULL');
    ExecuteSQLContOnExcept('UPDATE YTARIFSPARAMETRES SET YFO_OKSECTEURGEO="",YFO_OKADRESSE="",YFO_OKMODEEXP="" WHERE YFO_OKSECTEURGEO IS NULL');
    ExecuteSQLContOnExcept('UPDATE YTARIFS SET YTS_SECTEURGEO="",YTS_CASCSECTEURGEO="-", YTS_CASCTOUSSECT="-",YTS_CODEPOSTAL="",YTS_PAYS="",YTS_REGION="",YTS_MODEEXP=""  WHERE YTS_SECTEURGEO IS NULL');
    If (V_PGI.driver <> dbORACLE7) then
      ExecuteSQLContOnExcept('UPDATE YTARIFS SET YTS_POIDSRECHERCHE= '
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
    ExecuteSQLContOnExcept('UPDATE TIERSCOMPL SET  YTC_SECTEURGEO ="",YTC_TIMBRE=0,YTC_SURTAXE=0  WHERE YTC_SECTEURGEO IS NULL');
    ExecuteSQLContOnExcept('UPDATE WPDRLIG SET WPL_CONSOLIDE="-" WHERE WPL_CONSOLIDE IS NULL');
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
    ExecuteSQLContOnExcept('UPDATE WPARAMFONCTION SET WPF_PREFIXE="", WPF_NATUREPIECEG=""');
    // GC
    ExecuteSQLContOnExcept('UPDATE MODEPAIE SET MP_AVECINFOCOMPL = "-" WHERE MP_AVECINFOCOMPL IS NULL');
    ExecuteSQLContOnExcept('UPDATE MODEPAIE SET MP_AVECNUMAUTOR = "-" WHERE MP_AVECNUMAUTOR IS NULL ');
    ExecuteSQLContOnExcept('UPDATE MODEPAIE SET MP_COPIECBDANSCTRL = "-" WHERE MP_COPIECBDANSCTRL IS NULL ');
    ExecuteSQLContOnExcept('UPDATE MODEPAIE SET MP_AFFICHNUMCBUS = "-" WHERE MP_AFFICHNUMCBUS IS NULL ');
    ExecuteSQLContOnExcept('UPDATE MODEPAIE SET MP_CPTECAISSE = "", MP_JALREMBQ = "", MP_CPTEREMBQ = "" ');
    ExecuteSQLContOnExcept('UPDATE MODEPAIECOMPL SET MPC_JALREMBQ = "", MPC_CPTEREMBQ = ""  ');
  end;
End;

Procedure MajVer692 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    // GPAO
    ExecuteSQLContOnExcept('UPDATE LIGNECOMPL SET GLC_ARTICLECFX="" WHERE GLC_ARTICLECFX IS NULL');
    ExecuteSQLContOnExcept('UPDATE YTARIFSPARAMETRES SET YFO_OKSECTEURGEO="---" WHERE YFO_OKSECTEURGEO IS NULL OR YFO_OKSECTEURGEO=""');
    ExecuteSQLContOnExcept('UPDATE YTARIFSPARAMETRES SET YFO_OKADRESSE="---" WHERE YFO_OKADRESSE IS NULL OR YFO_OKADRESSE=""');
    ExecuteSQLContOnExcept('UPDATE YTARIFSPARAMETRES SET YFO_OKMODEEXP="---" WHERE YFO_OKMODEEXP IS NULL OR YFO_OKMODEEXP=""');
    //GC
    ExecuteSQLContOnExcept('UPDATE COMPTADIFFEREE SET GCD_TIERSPAYEUR = "" WHERE GCD_TIERSPAYEUR IS NULL');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_SOLDETRANSFO="-" WHERE GPP_SOLDETRANSFO IS NULL');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_REGROUPE="X" WHERE GPP_REGROUPE IS NULL');
    //GI/GA
    AglNettoieListes('AFMULAPPRECON', 'AFA_TRANSFENCOURS',nil);
    AglNettoieListes('AFMULAPPRECONANU', 'AFA_AFFAIRETRAENC;AFA_TRANSFENCOURS',nil);
  end;
  ExecuteSQLContOnExcept('update generaux set g_cutoffcompte="" ');
  ExecuteSQLContOnExcept('update generauxref set ger_cutoffcompte=""');
End;

Procedure MajVer694 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GPAO  reprise initialisation fait sur 677 pour 6.20
    ExecuteSQLContOnExcept('UPDATE ARTICLE SET GA_LOWLEVELCODE=-1 WHERE GA_LOWLEVELCODE IS NULL');
    // GPAO
    ExecuteSQLContOnExcept('DELETE FROM YTARIFSPARAMETRES WHERE YFO_FONCTIONNALITE="501" AND (YFO_ORIENTATION="TIE" OR YFO_ORIENTATION="ART")');
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
    {
    Laurent ABélard le 2 Octobre 2007 Le champ JTP_NATUREAUXI a disparu dans la SOCREF 850
    ExecuteSql ('update  JUTYPEPER set JTP_NATUREAUXI=""');
    }
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
ExecuteSQLContOnExcept('DELETE FROM CHOIXCOD WHERE CC_TYPE="PSC" AND (CC_LIBRE="" OR CC_LIBRE IS NULL)');
AglNettoieListes('PGACOMPTE', 'PSA_ETABLISSEMENT;PSD_SALARIE;PSA_LIBELLE;PSA_PRENOM;PSD_MONTANT;PSD_DATEDEBUT;PSD_DATEPAIEMENT;PSD_DATEFIN;PSD_ORIGINEMVT;PSD_ORDRE;PSD_RUBRIQUE',nil);
AglNettoieListes('PGVIRSALAIRES', 'PVI_SALARIE;PVI_LIBELLE;PVI_PRENOM;PVI_ETABLISSEMENT;PVI_DATEDEBUT;PVI_DATEFIN;PVI_MONTANT;PVI_PAYELE;PVI_BANQUE;PVI_BANQUEEMIS;PVI_TOPREGLE;PVI_AUXILIAIRE;PVI_RIBSALAIRE;PVI_RIBSALAIREEMIS',nil);
AglNettoieListes('PGMULMVTCP', 'PCN_ORDRE;PCN_ETABLISSEMENT;PCN_SALARIE;PCN_DATEVALIDITE;PCN_TYPECONGE;PCN_SENSABS;PCN_JOURS;PCN_HEURES;PCN_LIBELLE;PCN_CODETAPE;PCN_TYPEMVT;PCN_APAYES;PCN_BASE;PCN_NBREMOIS',nil);
AglNettoieListes('PGMULCONTRAT', 'PSA_ETABLISSEMENT;PCI_ETABLISSEMENT;PCI_SALARIE;PSA_LIBELLE;PSA_PRENOM;PCI_TYPECONTRAT;PCI_DEBUTCONTRAT;PCI_FINCONTRAT',nil);
ExecuteSQLContOnExcept('UPDATE PAIEENCOURS SET PPU_ECHEANCE = PPU_PAYELE');
//PDumet 12102005
v_pgi.enableDEShare := True;
V_PGI.StandardSurDP := True;
ExecuteSQLContOnExcept('update histoanalpaie set pha_ordreetat = (SELECT prm_ordreetat from remuneration WHERE ##prm_predefini## pha_rubrique=prm_rubrique) WHERE pha_naturerub="AAA" ');
//PDumet 12102005
v_pgi.enableDEShare := false;
V_PGI.StandardSurDP := False;
ExecuteSQLContOnExcept('UPDATE histoanalpaie set pha_ordreetat = 3 where pha_naturerub<>"AAA"');
If ExisteSQL('SELECT * FROM ABSENCESALARIE WHERE PCN_TYPEMVT="ABS"') then
   ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA = "X" WHERE SOC_NOM="SO_PGABSENCE"');
ExecuteSQLContOnExcept('UPDATE RHCOMPETENCES SET PCO_LIBELLECOMPL= ""');
// ExecuteSQLContOnExcept('UPDATE histoanalpaie SET PHA_COTREGUL= "..."');
// ExecuteSQLContOnExcept('UPDATE histobulletin SET PHB_COTREGUL= "..."');
// COMPTA
// AYEL 25102005 ExecuteSQLContOnExcept('UPDATE STDCPTA SET STC_DATEMODIF="' + UsdateTime(iDate1900) + '"');
End;


Procedure MajVer696 ;
var ss,SQL : string;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
       //GC
       ExecuteSQLContOnExcept('UPDATE STKMOUVEMENT SET GSM_CONTREMARQUE="-", GSM_FOURNISSEUR="", GSM_REFERENCE=""');
       ExecuteSQLContOnExcept('UPDATE DISPODETAIL SET GQD_CONTREMARQUE="-", GQD_FOURNISSEUR="", GQD_REFERENCE=""');
       //GPAO
       ExecuteSQLContOnExcept('update ARTICLECOMPL set GA2_FAMILLEMAT="" where GA2_FAMILLEMAT is null');
       ExecuteSQLContOnExcept('update PROFILART  set GPF_FAMILLEMAT="" where GPF_FAMILLEMAT is null');
       ExecuteSQLContOnExcept('update WORDRELIG  set WOL_SCMORDO="" where WOL_SCMORDO is null');
       ExecuteSQLContOnExcept('update WPARAMFONCTION  set WPF_FAMILLEMAT="", WPF_TYPCOMP="" where WPF_FAMILLEMAT is null');
       ExecuteSQLContOnExcept('update UCOMPFOURNI  set UCF_FAMILLEMAT="", UCF_TYPCOMP="" where UCF_FAMILLEMAT is null');
       //DP
       ExecuteSQLContOnExcept('UPDATE DPFISCAL SET DFI_DOMHORSFRANCE="", DFI_BAFORFAIT="", DFI_DISTRIBDIVID="-"');
       ExecuteSQLContOnExcept('UPDATE DPORGA SET DOR_UTILRESPCOMPTA="", DOR_UTILRESPSOCIAL="", DOR_UTILRESPJURID="", DOR_UTILCHEFGROUPE="", DOR_CABINETASSOC=""');
        //GIGA
       ExecuteSql ('UPDATE ANNUAIRE SET ANN_NATUREAUXI="CLI" where ANN_TIERS <>""');
       ExecuteSql ('UPDATE AFFAIRE set AFF_ISECOLE="-",AFF_DOMAINE=""');
       SS := UsDateTime(IDate1900);
       Sql := 'update afftiers set AFT_TYPEORIG="AFF",AFT_NUMORIG=0,AFT_TYPECONTACT="",AFT_AUXILIAIRE="",'
         +'AFT_NUMEROCONTACT=0,AFT_INTERVENTION="",AFT_DATECREATION="'+SS+'",AFT_DATEMODIF="'+SS
         +'",AFT_CREATEUR ="",AFT_UTILISATEUR=""';
       ExecuteSQLContOnExcept(SQL);
       ExecuteSQLContOnExcept('update afftiers  set AFT_TYPEINTERV="CLI" where  AFT_TIERS <>""');
       ExecuteSQLContOnExcept('update afftiers  set AFT_TYPEINTERV="RES" where  AFT_TIERS =""' );
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
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_MODELEWORD="" ');
    ExecuteSQLContOnExcept('UPDATE PARPIECECOMPL SET GPC_MODELEWORD=""  ');
  // GPAO
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="", GSN_SFLUXPICKING="" where GSN_QUALIFMVT="000"'
            + ' or GSN_QUALIFMVT="AAC" or gsn_qualifmvt="ACP" or gsn_qualifmvt="ADM" or gsn_qualifmvt="APF" or gsn_qualifmvt="APR" or gsn_qualifmvt="ATD"'
            + ' or gsn_qualifmvt="EAC" or gsn_qualifmvt="ECP" or gsn_qualifmvt="EDM" or gsn_qualifmvt="EEX" or gsn_qualifmvt="EIN" or gsn_qualifmvt="EIT"'
            + ' or gsn_qualifmvt="EPR" or gsn_qualifmvt="ETR" or gsn_qualifmvt="EVE" or gsn_qualifmvt="MPR"');

    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;AFF;", GSN_SFLUXPICKING="STD;PCO;PVE" where GSN_QUALIFMVT="CAF"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="BLQ", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;", GSN_SFLUXPICKING="STD;PCO;PVE" where GSN_QUALIFMVT="CBS"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="LBR", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="BLQ;", GSN_SFLUXPICKING="STD;PCO;PVE" where GSN_QUALIFMVT="CDS"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="<<Tous>>", GSN_SFLUXPICKING="<<Tous>>" where GSN_QUALIFMVT="CEM"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;", GSN_SFLUXPICKING="STD;" where GSN_QUALIFMVT="CIA"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="<<Tous>>", GSN_SFLUXPICKING="<<Tous>>" where GSN_QUALIFMVT="CLO"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;AFF;", GSN_SFLUXPICKING="STD;PCO;" where GSN_QUALIFMVT="RDM"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;AFF;", GSN_SFLUXPICKING="STD;PVE;" where GSN_QUALIFMVT="RPL"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;AFF;", GSN_SFLUXPICKING="STD;PCO;" where GSN_QUALIFMVT="RPR"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;AFF;", GSN_SFLUXPICKING="STD;PVE;" where GSN_QUALIFMVT="RTD"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;AFF;", GSN_SFLUXPICKING="STD;PVE;" where GSN_QUALIFMVT="RVE"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;", GSN_SFLUXPICKING="STD;CTR;" where GSN_QUALIFMVT="SAC"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;", GSN_SFLUXPICKING="STD;CTR;" where GSN_QUALIFMVT="SDM"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;BLQ;", GSN_SFLUXPICKING="<<Tous>>" where GSN_QUALIFMVT="SEX"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="<<Tous>>", GSN_SFLUXPICKING="<<Tous>>" where GSN_QUALIFMVT="SIN"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="<<Tous>>", GSN_SFLUXPICKING="<<Tous>>" where GSN_QUALIFMVT="SIT"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;AFF;", GSN_SFLUXPICKING="STD;PCO;" where GSN_QUALIFMVT="SPR"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;", GSN_SFLUXPICKING="STD;PVE" where GSN_QUALIFMVT="STR"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPODISPATCH="", GSN_SFLUXDISPATCH="", GSN_SDISPOPICKING="LBR;AFF;", GSN_SFLUXPICKING="STD;PVE;" where GSN_QUALIFMVT="SVE"');

    ExecuteSQLContOnExcept('UPDATE QTYPTRANS SET QTY_PORTDEPART="", QTY_PORTARRIVEE="" WHERE QTY_PORTDEPART IS NULL');

    InsertChoixCode('YTR', 'X', 'pas de recherche', '', '');
    InsertChoixCode('YTU', 'X', 'ne somme pas', '', '');

    //ExecuteSQLContOnExcept('update Dispodetail set gqd_dateperemption="' + UsDateTime(iDate2099) + '" where gqd_dateperemption="' + UsdateTime(iDate1900) + '"');
    //ExecuteSQLContOnExcept('update stkmouvement set gsm_dateperemption="' + UsDateTime(iDate2099) + '" where gsm_dateperemption="' + UsdateTime(iDate1900) + '"');
    // mofif KB 31082005
    ExecuteSQLContOnExcept('UPDATE DISPODETAIL SET GQD_DATEPEREMPTION="' + UsDateTime(iDate2099) + '" WHERE GQD_DATEPEREMPTION IS NULL OR GQD_DATEPEREMPTION="' + UsdateTime(iDate1900) + '"');
    ExecuteSQLContOnExcept('UPDATE STKMOUVEMENT SET GSM_DATEPEREMPTION="' + UsDateTime(iDate2099) + '" WHERE GSM_DATEPEREMPTION IS NULL OR GSM_DATEPEREMPTION="' + UsdateTime(iDate1900) + '"');
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
     ExecuteSQLContOnExcept('UPDATE DPORGA SET DOR_NONTRAITE="-", DOR_MOTIFNONTRAITE="", DOR_WETABENTITE="", DOR_WCATREVENUS="", DOR_WNATUREATTEST="", DOR_WREGIMEIMPO="", DOR_WREGIMETVA="", DOR_WTYPEIMPO=""');
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
    AglNettoieListes('YYMULSELDOSS', '',nil);
         //fin bureau
	//GIGA
    ExecuteSql ('update afmodeletache set afm_commentaire=""');
    AglNettoieListes('AFMULCOMPARTICLE', 'AAC_COMPETENCE;AAC_CODEARTICLE;AAC_ARTICLE',nil);
    if isOracle then
        begin
        ExecuteSQLContOnExcept('UPDATE AFPLANNING set APL_HEUREDEB_PLA = APL_DATEDEBPLA + 8/24');
        ExecuteSQLContOnExcept('UPDATE AFPLANNING set APL_HEUREFIN_PLA = APL_DATEFINPLA + 18/24');
        ExecuteSQLContOnExcept('UPDATE AFPLANNING set APL_HEUREDEB_REAL = APL_DATEDEBREAL + 8/24'	);
        ExecuteSQLContOnExcept('UPDATE AFPLANNING set APL_HEUREFIN_REAL = APL_DATEFINREAL + 18/24');
        end
    else if isDB2 then      //rien à faire
    else begin  // cas SQl ou access ou ...
      ExecuteSQLContOnExcept('UPDATE AFPLANNING set APL_HEUREDEB_PLA = APL_DATEDEBPLA + "8:00:00"' );
      ExecuteSQLContOnExcept('UPDATE AFPLANNING set APL_HEUREFIN_PLA = APL_DATEFINPLA + "18:00:00"' );
      ExecuteSQLContOnExcept('UPDATE AFPLANNING set APL_HEUREDEB_REAL = APL_DATEDEBREAL + "8:00:00"');
      ExecuteSQLContOnExcept('UPDATE AFPLANNING set APL_HEUREFIN_REAL = APL_DATEFINREAL + "18:00:00"');
    end;
	//fin GIGA
  //GC
    ExecuteSql ('UPDATE LIGNECOMPL SET GLC_REPRESENTANT2="", GLC_REPRESENTANT3="", GLC_COMMISSIONR2=0, GLC_COMMISSIONR3=0, GLC_TYPECOM2="-", GLC_TYPECOM3="-", GLC_VALIDECOM2="-", GLC_VALIDECOM3="-", GLC_MONTANTTPF=0, GLC_COMMISSIONNABL="-"');
    ExecuteSql ('UPDATE ARTICLECOMPL SET GA2_MONTANTTPF=0 ');
  // GPAO
    ExecuteSQLContOnExcept('update STKNATURE SET GSN_CALLGSS="-" WHERE GSN_CALLGSS IS NULL');
    UpdateDecoupeLigne('GL_COEFCONVQTE=1', ' AND (ISNULL(GL_ARTICLE, " ") <> " ")');
  end;

  ExecuteSQLContOnExcept('update remuneration set prm_ordreetat=9 where prm_ordreetat=6');
  ExecuteSQLContOnExcept('update remuneration set prm_ordreetat=7 where prm_ordreetat=5');
  ExecuteSQLContOnExcept('update remuneration set prm_ordreetat=6 where prm_ordreetat=4');

  //PDumet 12102005
  v_pgi.enableDEShare := True;
  V_PGI.StandardSurDP := True;
  ExecuteSQLContOnExcept('UPDATE histobulletin set phb_ordreetat=(Select pct_ordreetat from cotisation '
     +' where ##PCT_predefini## pct_rubrique=phb_rubrique) where phb_naturerub<>"AAA" and phb_rubrique in (select pct_rubrique from cotisation cot where ##cot.pct_predefini## cot.pct_themecot="RDC")');
  ExecuteSQLContOnExcept('UPDATE histoanalpaie set pha_ordreetat=(Select pct_ordreetat from cotisation '
     +' where ##PCT_predefini## pct_rubrique=pha_rubrique) where pha_naturerub<>"AAA" and pha_rubrique in (select pct_rubrique from cotisation cot where ##cot.pct_predefini## cot.pct_themecot="RDC")');
  ExecuteSQLContOnExcept('UPDATE histobulletin set phb_ordreetat=(Select prm_ordreetat from remuneration '
     +' where ##Prm_predefini## prm_rubrique=phb_rubrique) where phb_naturerub="AAA" and phb_ordreetat>3');
  ExecuteSQLContOnExcept('UPDATE histoanalpaie set pha_ordreetat=(Select prm_ordreetat from remuneration '
    +' where ##Prm_predefini## prm_rubrique=pha_rubrique) where pha_naturerub="AAA" and pha_ordreetat>3');

  v_pgi.enableDEShare := False;

  V_PGI.StandardSurDP := FALSE;


  If ExisteSQL('SELECT PFO_SALARIE FROM FORMATIONS WHERE PFO_TYPEPLANPREV IS NULL OR PFO_TYPEPLANPREV=""') then

  begin

    ExecuteSQLContOnExcept('UPDATE FORMATIONS SET PFO_TYPEPLANPREV="PLF", PFO_HTPSTRAV=PFO_NBREHEURE,PFO_HTPSNONTRAV=0,PFO_PAYER="-",PFO_DATEPAIE="'+SS+'",PFO_TYPOFORMATION="001",PFO_ALLOCFORM=0');

    ExecuteSQLContOnExcept('UPDATE INSCFORMATION SET PFI_TYPEPLANPREV="PLF",PFI_HTPSTRAV=PFI_DUREESTAGE,PFI_HTPSNONTRAV=0,'

      + 'PFI_TYPOFORMATION="001",PFI_DATEDIF=PFI_DATECREATION,'
      + 'PFI_NATUREFORM=(SELECT PST_NATUREFORM FROM STAGE WHERE PST_CODESTAGE=PFI_CODESTAGE AND PST_MILLESIME=PFI_MILLESIME),PFI_ALLOCFORM=0');

  end;

  ExecuteSQLContOnExcept('update ducsentete set pdu_typbordereau="913" where '+
    'PDU_ABREGEPERIODE not like "%%00" AND PDU_DUCSDOSSIER = "-" AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE '+
        'ON PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE  POG_NATUREORG ="100")');
  ExecuteSQLContOnExcept('update ducsentete set pdu_typbordereau="914" where '+
    'PDU_ABREGEPERIODE not like "%%00" AND PDU_DUCSDOSSIER = "X" AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE '+
        'ON PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE  POG_NATUREORG ="100")');
  ExecuteSQLContOnExcept('update ducsentete set pdu_typbordereau="915" where '+
    'PDU_ABREGEPERIODE  like "%%00" AND PDU_DUCSDOSSIER = "-" AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE '+
        'ON PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE POG_NATUREORG ="100")');
  ExecuteSQLContOnExcept('update ducsentete set pdu_typbordereau="916" where '+
    'PDU_ABREGEPERIODE  like "%%00" AND PDU_DUCSDOSSIER = "X" AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE '+
        'ON PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE POG_NATUREORG ="100")');
  ExecuteSQLContOnExcept('update ducsentete set pdu_typbordereau="920" where '+
    'PDU_ABREGEPERIODE  not like "%%00" AND PDU_DUCSDOSSIER = "-" AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE ON '+
        'PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE POG_NATUREORG ="200")');
  ExecuteSQLContOnExcept('update ducsentete set pdu_typbordereau="921" where '+
    'PDU_ABREGEPERIODE  not like "%%00" AND PDU_DUCSDOSSIER = "X" AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE '+
       'ON PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE POG_NATUREORG ="200")');
  ExecuteSQLContOnExcept('update ducsentete set pdu_typbordereau="922" where '+
    'PDU_ABREGEPERIODE   like "%%00" AND PDU_DUCSDOSSIER = "-" AND PDU_ORGANISME '+
      'IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE ON '+
        'PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE POG_NATUREORG ="200")');
  ExecuteSQLContOnExcept('update ducsentete set pdu_typbordereau="923" where '+
    'PDU_ABREGEPERIODE   like "%%00" AND PDU_DUCSDOSSIER = "X" AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE '+
        'ON PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE POG_NATUREORG ="200")');
  ExecuteSQLContOnExcept('update ducsentete set pdu_typbordereau="930" where '+
    'PDU_ABREGEPERIODE  like "%%0%" AND PDU_ABREGEPERIODE  not  like "%%00"  AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE '+
        'ON PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE POG_NATUREORG ="300")');
  ExecuteSQLContOnExcept('update ducsentete set pdu_typbordereau="931" where '+
    'PDU_ABREGEPERIODE   not like "%%0%" AND PDU_ABREGEPERIODE  not  like "%%00"  AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE '+
        'ON PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE POG_NATUREORG ="300")');
  ExecuteSQLContOnExcept('update ducsentete set pdu_typbordereau="932" where '+
    'PDU_ABREGEPERIODE  like "%%00"  AND '+
      'PDU_ORGANISME IN (SELECT PDU_ORGANISME FROM DUCSENTETE LEFT JOIN ORGANISMEPAIE '+
        'ON PDU_ORGANISME=POG_ORGANISME AND PDU_ETABLISSEMENT=POG_ETABLISSEMENT WHERE POG_NATUREORG ="300")');
  ExecuteSQLContOnExcept('update ducsentete set PDU_NBSALQ944=0, '+
    'PDU_NBSALQ945=0, PDU_NBSALQ960=0, PDU_NBSALQ961=0,PDU_NBSALQ962=0, '+
      'PDU_NBSALQ963=0,PDU_NBSALQ964=0, PDU_NBSALQ965=0,PDU_NBSALQ966=0, '+
        'PDU_NBSALQ967=0,PDU_ECARTZE1="-",PDU_ECARTZE2="-",PDU_ECARTZE3="-",'+
          'PDU_ECARTZE4="-",PDU_ECARTZE5="-",PDU_ECARTZE6="-",PDU_ECARTZE7="-", PDU_ECARTZE8 ="-"');
  ExecuteSQLContOnExcept('update emetteursocial set PET_EMAILDUCS=""');
  ExecuteSQLContOnExcept('update ducsdetail set PDD_CODECOMMUNE=""');
  ExecuteSQLContOnExcept('update organismepaie set POG_TITULAIRECPT=""');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEPERIOD="913" where '+
    'POG_NATUREORG="100" AND POG_DUCSDOSSIER="-" AND (POG_PERIODICITDUCS="M" OR POG_PERIODICITDUCS="T")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEPERIOD="914" where '+
    'POG_NATUREORG="100" AND POG_DUCSDOSSIER="X" AND (POG_PERIODICITDUCS="M" OR POG_PERIODICITDUCS="T")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEPERIOD="915" where '+
    'POG_NATUREORG="100" AND POG_DUCSDOSSIER="-" AND (POG_PERIODICITDUCS="A")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEPERIOD="916" where '+
    'POG_NATUREORG="100" AND POG_DUCSDOSSIER="X" AND (POG_PERIODICITDUCS="A")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEPERIOD="920" where '+
    'POG_NATUREORG="200" AND POG_DUCSDOSSIER="-" AND (POG_PERIODICITDUCS="M" OR POG_PERIODICITDUCS="T")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEPERIOD="921" where '+
    'POG_NATUREORG="200" AND POG_DUCSDOSSIER="X" AND ( POG_PERIODICITDUCS="M" OR POG_PERIODICITDUCS="T")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEPERIOD="922" where '+
    'POG_NATUREORG="200" AND POG_DUCSDOSSIER="-" AND ( POG_PERIODICITDUCS="A")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEPERIOD="923" where '+
    'POG_NATUREORG="200" AND POG_DUCSDOSSIER="X" AND (POG_PERIODICITDUCS="A")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEPERIOD="930" where '+
    'POG_NATUREORG="300" AND (POG_PERIODICITDUCS="T")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEPERIOD="931" where '+
    'POG_NATUREORG="300" AND (POG_PERIODICITDUCS="M")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEPERIOD="932" where '+
    'POG_NATUREORG="300" AND (POG_PERIODICITDUCS="A")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="913" where '+
    'POG_NATUREORG="100" AND POG_DUCSDOSSIER="-" AND (POG_AUTREPERIODUCS="M" OR POG_AUTREPERIODUCS="T")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="914" where '+
    'POG_NATUREORG="100" AND POG_DUCSDOSSIER="X" AND (POG_AUTREPERIODUCS="M" OR POG_AUTREPERIODUCS="T")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="915" where '+
    'POG_NATUREORG="100" AND POG_DUCSDOSSIER="-" AND (POG_AUTREPERIODUCS="A")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="916" where '+
    'POG_NATUREORG="100" AND POG_DUCSDOSSIER="X" AND (POG_AUTREPERIODUCS="A")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="920" where '+
    'POG_NATUREORG="200" AND POG_DUCSDOSSIER="-" AND (POG_AUTREPERIODUCS="M" OR POG_AUTREPERIODUCS="T")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="921" where '+
    'POG_NATUREORG="200" AND POG_DUCSDOSSIER="X" AND (POG_AUTREPERIODUCS="M" OR POG_AUTREPERIODUCS="T")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="922" where '+
    'POG_NATUREORG="200" AND POG_DUCSDOSSIER="-" AND (POG_AUTREPERIODUCS="A")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="923" where '+
    'POG_NATUREORG="200" AND POG_DUCSDOSSIER="X" AND (POG_AUTREPERIODUCS="A")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="930" where '+
    'POG_NATUREORG="300" AND (POG_AUTREPERIODUCS="T")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="931" where '+
    'POG_NATUREORG="300" AND (POG_AUTREPERIODUCS="M")');
  ExecuteSQLContOnExcept('update ORGANISMEPAIE set POG_TYPEAUTPERIOD="932" where '+
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
    ExecuteSQLContOnExcept('UPDATE COMMERCIAL SET GCL_HRRANG="" WHERE (GCL_HRRANG IS NULL)');
    ExecuteSQLContOnExcept('UPDATE NOMENLIG SET GNL_QTEMINMENU=0, GNL_QTEMAXMENU=0 , GNL_POIDSMENU=0 WHERE (GNL_QTEMINMENU IS NULL)');
    ExecuteSQLContOnExcept('UPDATE ARTICLECOMPL SET GA2_DEPOT="",GA2_HRQTESTAT="-",GA2_INTERDITACHAT="-" WHERE (GA2_DEPOT IS NULL)');
    ExecuteSQLContOnExcept('UPDATE HRGROUPESSTAT SET HGS_LFAMRES1="",HGS_LFAMRES2="",HGS_LFAMRES3="",HGS_LFAMRES4="",HGS_LFAMRES5="",'+
               'HGS_LFAMRES6="", HGS_LIBELLECHAMP1="",HGS_LIBELLECHAMP2="",HGS_LIBELLECHAMP3="",HGS_LIBELLECHAMP4="",HGS_LIBELLECHAMP5="",HGS_LIBELLECHAMP6=""  WHERE HGS_LFAMRES1 IS NULL');
    ExecuteSQLContOnExcept('UPDATE HRDOSRES SET HDR_CODEFACT="", HDR_NOMBREPERSOFF=0 WHERE HDR_CODEFACT IS NULL');
    ExecuteSQLContOnExcept('UPDATE HRDOSRES SET HDR_LIBELLERESA1="",HDR_LIBELLERESA2="" WHERE HDR_LIBELLERESA1 IS NULL');
    ExecuteSQLContOnExcept('UPDATE HRDOSRES SET HDR_TYPRESASSOC="" WHERE HDR_TYPRESASSOC IS NULL');
    ExecuteSQLContOnExcept('UPDATE HRALLOTEMENT SET HAL_TYPRESASSOC="" WHERE HAL_TYPRESASSOC IS NULL');
    ExecuteSQLContOnExcept('UPDATE HRLIGNE SET HRL_ORDREFACTURE=0 WHERE HRL_ORDREFACTURE IS NULL');
    ExecuteSQLContOnExcept('UPDATE HRDOSSIER SET HDC_FAMILLETAXE1="" WHERE HDC_FAMILLETAXE1 IS NULL');
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
    ExecuteSQLContOnExcept('UPDATE PARCAISSE SET GPK_LFAMRES="" WHERE (GPK_LFAMRES IS NULL)');
    ExecuteSQLContOnExcept('UPDATE ARTICLECOMPL SET GA2_CUISINE1="",GA2_CUISINE2="-",GA2_CUISINE3="-",GA2_CUISINE4="-",GA2_CUISINE5="-" WHERE (GA2_CUISINE1 IS NULL)');
    //DP
    // Requête de réparation pour les liasses (cf YJ/SL)
    ExecuteSQLContOnExcept('update DPFISCAL set DFI_OPTIONAUTID=isnull(DFI_OPTIONAUTID,"-"),  DFI_OPTIONRDSUP=isnull(DFI_OPTIONRDSUP,"-"), DFI_OPTIONRSS=isnull(DFI_OPTIONRSS,"-"), DFI_OPTIONREPORT=isnull(DFI_OPTIONREPORT,"-"), '
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
     ExecuteSQLContOnExcept('UPDATE AFFAIRE set AFF_SEUILINFO = 0');
     // GPAO
    InsertChoixCode('YTN', 'H', 'des heures', 'Heures', '') ;
    ExecuteSQLContOnExcept('UPDATE PORT SET GPO_TIERS="",GPO_VENTPIECE="",GPO_VENTCOMPTA="",GPO_REPARTITION="",GPO_FACTURABLE="X",GPO_TYPEPDR="", GPO_TYPEFOURNI=""  WHERE GPO_TIERS IS NULL');
    ExecuteSQLContOnExcept('UPDATE YTARIFS SET YTS_CODEPORT="",YTS_NATURETRAVAIL="",YTS_RESSOURCE="",YTS_TARIFRESSOURCE="",YTS_ATTRIBUTION="", YTS_CASCRESSOURCE="-",yts_casctarifresso="-",YTS_CASCTOUSRESSO="-" WHERE YTS_CODEPORT IS NULL');
    ExecuteSQLContOnExcept('UPDATE PIEDPORT SET GPT_MONTANTREALISE=0 WHERE GPT_MONTANTREALISE IS NULL');
    // ExecuteSQLContOnExcept('UPDATE LIGNECOMPL SET GLC_GUIDGSM="" WHERE GLC_GUIDGSM IS NULL');
    ExecuteSQLContOnExcept('UPDATE YTARIFSPARAMETRES SET YFO_CODEPORT="",YFO_OKCODEPORT="----",YFO_OKNATURETRA="----", YFO_OKRESSOURCE="----" WHERE YFO_CODEPORT IS NULL');
    ExecuteSQLContOnExcept('UPDATE PORT SET GPO_TYPEFRAIS="FAN" WHERE GPO_TYPEPORT="TRA"');
    ExecuteSQLContOnExcept('UPDATE PORT SET GPO_TYPEFRAIS="CIN" WHERE GPO_TYPEFRAIS=""');
    ExecuteSQLContOnExcept('UPDATE PORT SET GPO_TYPEPORT="TAR" WHERE GPO_TYPEPORT="TRA"');
    ExecuteSQLContOnExcept('UPDATE PORT SET GPO_VENTPIECE="P", GPO_REPARTITION="Q" WHERE GPO_TYPEPORT="TAR" AND GPO_VENTPIECE IS NULL');
    ExecuteSQLContOnExcept('UPDATE TIERSCOMPL SET YTC_TYPEFOURNI="" WHERE YTC_TYPEFOURNI IS NULL');
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
    ExecuteSQLContOnExcept('UPDATE LISTEINVLIG SET GIL_REFERENCE = "", GIL_FOURNISSEUR = "", GIL_CONTREMARQUE = "-"');
    //GPAO
    //ExecuteSQLContOnExcept('UPDATE paramsoc SET soc_data="001" WHERE soc_nom = "SO_GCGESTUNITEMODE"');
    //ExecuteSQLContOnExcept('UPDATE paramSoc SET soc_data="PIE" WHERE soc_nom = "SO_GCDEFQUALIFUNITEGA"');
    //ExecuteSQLContOnExcept('UPDATE paramSoc SET soc_data="UNI" WHERE soc_nom = "SO_GCDEFUNITEGA"');
    ExecuteSQLContOnExcept('UPDATE winitchamplig SET wil_famchamp="DIV" WHERE wil_nomchamp = "GA_UNITEQTEVTE"');
    ExecuteSQLContOnExcept('UPDATE PORT SET GPO_TIERS="",GPO_VENTPIECE="A",GPO_VENTCOMPTA="P",GPO_REPARTITION="Q",GPO_FACTURABLE="X",GPO_TYPEPDR="", GPO_TYPEFOURNI=""  WHERE GPO_TIERS IS NULL');
    // La requète ci dessous remplace la requète d'init. suivante :
    ExecuteSQLContOnExcept('UPDATE PORT SET GPO_TIERS="",GPO_VENTPIECE="",GPO_VENTCOMPTA="",GPO_REPARTITION="",GPO_FACTURABLE="X",GPO_TYPEPDR="", GPO_TYPEFOURNI=""  WHERE GPO_TIERS IS NULL');
    //GIGA
    ExecuteSQLContOnExcept('UPDATE FACTAFF SET AFA_DATEPIECE="' + SS + '"');
  end;
  ExecuteSQLContOnExcept('update journal set j_tresovalid="-",j_tresolibre="-"');
  ExecuteSQLContOnExcept('update jalref set jr_tresoimport="",jr_tresochainage="-",jr_tresomontant="-",jr_tresodate="-",jr_tresolibelle="-",jr_tresovalid="-",jr_tresolibre="-"');
End;

Procedure MajVer700 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //DP
   // correction 27092005 MD ExecuteSQLContOnExcept('update DPSOCIAL set DSO_PERURSSAF="", DSO_PERASSEDIC="", DSO_PERIRC="", DSO_REMBULLPAIE="", DSO_DECLHANDICAP="", DSO_DECLMO=""');
    ExecuteSQLContOnExcept('update DPSOCIAL set DSO_PERURSSAF="", DSO_PERASSEDIC="", DSO_PERIRC="", DSO_REMBULLPAIE=0, DSO_DECLHANDICAP="-", DSO_DECLMO="-"');
    ExecuteSQLContOnExcept('update DPORGA set DOR_WMODETRAITCPT="", DOR_WTYPEBNC=""');
    ExecuteSQLContOnExcept('UPDATE CONTACT SET C_CLEFAX="",C_CLETELEX=""');
    //GIGA
    executeSQL ('UPDATE AFVALINDICE set AFV_BASE100="-"');
    //GPAO
    (*
    if (V_PGI.Driver <> dbORACLE7 ) then // pas supporté en DB2
        MajPoidsTarifs;
    *)
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
    ExecuteSQLContOnExcept('UPDATE PAYS SET PY_CODEINSEE="" WHERE (PY_CODEINSEE IS NULL)');
    //DP
    ExecuteSQLContOnExcept('UPDATE YMESSAGES SET YMS_INBOX="-"');
    ExecuteSQLContOnExcept('UPDATE DOSSIER SET DOS_WINSTALL="", DOS_TYPEMAJ="", DOS_DETAILMAJ=""');
    ExecuteSQLContOnExcept('UPDATE YMESSAGES SET YMS_INBOX="X" WHERE YMS_BROUILLON="X"');

    // GPAO
    ExecuteSQLContOnExcept('UPDATE PIEDPORT SET GPT_FACTURABLE="X" WHERE GPT_FACTURABLE IS NULL');
    ExecuteSQLContOnExcept('UPDATE WPARAMFONCTION SET WPF_USER="", WPF_CONTEXTE="" WHERE WPF_USER IS NULL');
    ExecuteSQLContOnExcept('UPDATE STKMOUVEMENT SET GSM_NUMPREPAORI=0 WHERE GSM_NUMPREPAORI IS NULL');
    ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA="NON" WHERE SOC_NOM="SO_FRAISINDPR" AND SOC_DATA="-"');
    ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA="THE" WHERE SOC_NOM="SO_FRAISINDPR" AND SOC_DATA="X"');
    ExecuteSQLContOnExcept('UPDATE WORDRERES SET WOR_RESSOURCESCM="-" WHERE WOR_RESSOURCESCM IS NULL');
    ExecuteSQLContOnExcept('UPDATE WGAMMERES SET WGR_RESSOURCESCM="-" WHERE WGR_RESSOURCESCM IS NULL');
    ExecuteSQLContOnExcept('UPDATE WGAMMELIG SET WGL_RESSOURCESCM="-",WGL_TYPEOPE="" WHERE WGL_RESSOURCESCM IS NULL');
    ExecuteSQLContOnExcept('UPDATE WORDREGAMME SET WOG_RESSOURCESCM="-",WOG_TYPEOPE="" WHERE WOG_RESSOURCESCM IS NULL');
    ExecuteSQLContOnExcept('UPDATE ARTICLE SET GA_CODEFORME="PIE" WHERE ISNULL(GA_CODEFORME,"") = ""');
    //GC
    ExecuteSQLContOnExcept('UPDATE COMMERCIAL SET GCL_COMMENCAISS ="-" ');
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
ExecuteSQLContOnExcept('UPDATE MOTIFABSENCE SET PMA_PGCOLORVAL = "",PMA_PGCOLORATT = "",PMA_PGCOLORREF = "",PMA_PGCOLORANN = ""');

End;

Procedure MajVer702 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //giGA
    ExecuteSQLContOnExcept('update factaff set afa_etatvisa="VIS" where afa_typeche="NOR" and afa_etatvisa<>"VIS"');
    //GRC
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_INFOSCPLPIECE="-"');
    ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA = "<<aucun>>" WHERE SOC_NOM = "SO_QUALIFMVTWPC"');
    InsertChoixCode('RGD', 'RD1', '.-Table libre 1', 'Table libre 1', '');
    InsertChoixCode('RGD', 'RD2', '.-Table libre 2', 'Table libre 2', '');
    InsertChoixCode('RGD', 'RD3', '.-Table libre 3', 'Table libre 3', '');
    //GPAO
//    ExecuteSQLContOnExcept('UPDATE CATALOGU SET GCA_QUALIFUNITEACH="PIE" WHERE ISNULL(GCA_QUALIFUNITEACH, "") = ""');
    if ExisteSQL('SELECT 1 FROM CATALOGU WHERE ISNULL(GCA_QUALIFUNITEACH, "") = ""') then
    begin
      if ExisteSQL('SELECT 1 FROM MEA WHERE GME_MESURE = "UNI"') then
        ExecuteSQLContOnExcept('UPDATE CATALOGU SET GCA_QUALIFUNITEACH="UNI" WHERE ISNULL(GCA_QUALIFUNITEACH, "") = ""')
      else if ExisteSQL('SELECT 1 FROM MEA WHERE GME_MESURE = "U"') then
        ExecuteSQLContOnExcept('UPDATE CATALOGU SET GCA_QUALIFUNITEACH="U" WHERE ISNULL(GCA_QUALIFUNITEACH, "") = ""')
      else
      begin
        if not ExisteSQL('SELECT 1 FROM MEA WHERE GME_MESURE = "UNI"') then
          ExecuteSQLContOnExcept('INSERT INTO MEA (GME_QUALIFMESURE,GME_MESURE,GME_LIBELLE,GME_QUOTITE) VALUES ("PIE", "UNI", "Unité", 1)');
        ExecuteSQLContOnExcept('UPDATE CATALOGU SET GCA_QUALIFUNITEACH="UNI" WHERE ISNULL(GCA_QUALIFUNITEACH, "") = ""')
      end;
    end;
    MaJdesPrixGSMFromBlocNote;
    (*
    AjouteNewTarifs;
    *)
    //GC
    ExecuteSQLContOnExcept('update parpiece set gpp_qualifmvt="RCC" where gpp_naturepieceg="CC" ');
    ExecuteSQLContOnExcept('update parpiece set gpp_qualifmvt="SCC" where gpp_naturepieceg="BLC" ');
    ExecuteSQLContOnExcept('update parpiece set gpp_qualifmvt="SCC" where gpp_naturepieceg="FAC" ');
    ExecuteSQLContOnExcept('update parpiece set gpp_qualifmvt="ECC" where gpp_naturepieceg="BLF" ');
    ExecuteSQLContOnExcept('update parpiece set gpp_qualifmvt="ECC" where gpp_naturepieceg="FF"  ');
    ExecuteSQLContOnExcept('update parpiece set gpp_qualifmvt="ACC" where gpp_naturepieceg="CF"  ');
    ExecuteSQLContOnExcept('update parpiece set gpp_qualifmvt="RCC" where gpp_naturepieceg="PRE" ');
    ExecuteSQLContOnExcept('update parpiece set gpp_qualifmvt="ACC" where gpp_naturepieceg="PRF" ');
    //PAIE
    ExecuteSQLContOnExcept('UPDATE ABSENCESALARIE SET PCN_REPRISEARRET="'+UsDateTime (IDate1900)+'", PCN_DATEATTEST="'+UsDateTime (IDate1900)+'"');
    AglNettoieListes('PGBILANSOCIAL','PBC_DATEDEBUT;PBC_DATEFIN;PBC_BSPRESENTATION;PBC_ETABLISSEMENT;PBC_BSINDTHEME;(PBC_INDICATEURBS);PBC_INDICATEURBS;PBC_CATBILAN;PBC_VALCAT;PBC_PGINDICATION;PBC_NBELEMENT;PBC_NODOSSIER;PBC_INCREM',nil);
    AglNettoieListes('PGBSINDSELECT','PID_PREDEFINI;PID_NODOSSIER;PID_BSPRESENTATION;PID_LIBELLE;PID_PGPERIODICITE',nil);
    AglNettoieListes('PBSINDICATEURS','PBI_TYPINDICATBS;PBI_INDICATEURBS;PBI_NODOSSIER;PBI_LIBELLE;PBI_PREDEFINI;PBI_BSINDTHEME;PBI_FOURCHETTEMINI;PBI_FOURCHETTEMAXI',nil);
    ExecuteSQLContOnExcept('UPDATE DUCSENTETE SET PDU_MASSEANNUEL = 0, PDU_REMUNDADS = 0');
    ExecuteSQLContOnExcept('UPDATE CONVENTIONCOLL SET PCV_IDCC="",PCV_NUMERODIV="",PCV_LIENCCN=""');
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
        ExecuteSQLContOnExcept('UPDATE YDATATYPETREES SET YDT_CODEHDTLINK ="'+'GC'+stCle+IntToStr(i)+'" WHERE YDT_CODEHDTLINK="'+stCle+IntToStr(i)+'"');
        ExecuteSQLContOnExcept('UPDATE YDATATYPELINKS SET YDL_CODEHDTLINK ="'+'GC'+stCle+IntToStr(i)+'",YDL_PREDEFINI="CEG" WHERE YDL_CODEHDTLINK="'+stCle+IntToStr(i)+'"');
      end;
    end;
    if GetParamSoc ('SO_RTACTTABHIE') then
    begin
      for i := 1 to 3 do
      begin
        stCle := 'RTRPRLIBACTION';
        ExecuteSQLContOnExcept('UPDATE YDATATYPETREES SET YDT_CODEHDTLINK ="'+'GC'+stCle+IntToStr(i)+'" WHERE YDT_CODEHDTLINK="'+stCle+IntToStr(i)+'"');
        ExecuteSQLContOnExcept('UPDATE YDATATYPELINKS SET YDL_CODEHDTLINK ="'+'GC'+stCle+IntToStr(i)+'",YDL_PREDEFINI="CEG" WHERE YDL_CODEHDTLINK="'+stCle+IntToStr(i)+'"');
      end;
    end;
    //GPAO
    ExecuteSQLContOnExcept('UPDATE ARTICLECOMPL SET GA2_DATEMODIFBT="'+UsDateTime(iDate1900)+'" WHERE GA2_DATEMODIFBT IS NULL');
    ExecuteSQLContOnExcept('UPDATE USERGRP SET UG_ENVIRONNEMENT="" WHERE UG_ENVIRONNEMENT IS NULL');
    ExecuteSQLContOnExcept('UPDATE DISPO SET GQ_DELAIMOYEN=0 WHERE GQ_DELAIMOYEN IS NULL');
    ExecuteSQLContOnExcept('UPDATE QSIMULATION SET QSM_MODESIMU="" WHERE QSM_MODESIMU IS NULL');
    ExecuteSQLContOnExcept('UPDATE QARTTECH SET QAR_INACTIF="-" WHERE QAR_INACTIF IS NULL');
    ExecuteSQLContOnExcept('UPDATE QBPSESSIONBP SET QBS_OKOBJPREV="-", QBS_OBJPREV="X" WHERE QBS_OBJPREV IS NULL');
    ExecuteSQLContOnExcept('UPDATE WPDRLIG SET WPL_CONSOLIDE="-" WHERE WPL_CONSOLIDE IS NULL');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_MAJPRIXVALO="" WHERE GSN_MAJPRIXVALO IS NULL');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_MAJPRIXVALO="DPA;DPR;PPA;PPR" WHERE GSN_QUALIFMVT IN ("CLO","EAC","ECP","EDM","EEX","EIN","EIT","EPR","ETR","000")');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_MAJPRIXVALO="DPA;DPR" WHERE GSN_QUALIFMVT IN ("AAC","APR","ATD")');
    ExecuteSQLContOnExcept(' UPDATE RESSOURCE SET ARS_RESSOURCESCM="-",ARS_GRP="" '
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
ExecuteSQLContOnExcept('UPDATE HISTOBULLETIN SET PHB_ORDREETAT = 3 WHERE PHB_NATURERUB <> "AAA"');
ExecuteSQLContOnExcept('UPDATE HISTOANALPAIE SET PHA_ORDREETAT = 3 WHERE PHA_NATURERUB <> "AAA"');
ExecuteSQLContOnExcept('UPDATE histobulletin set phb_omtsalarial = phb_ordreetat');
ExecuteSQLContOnExcept('UPDATE histoanalpaie set pha_omtsalarial = pha_ordreetat');

//PDumet 12102005
v_pgi.enableDEShare := True;
V_PGI.StandardSurDP := True;

ExecuteSQLContOnExcept('UPDATE HISTOANALPAIE SET PHA_OMTSALARIAL=(SELECT PCT_ORDREETAT FROM COTISATION WHERE ##PCT_PREDEFINI## PCT_RUBRIQUE=PHA_RUBRIQUE) '
       +' WHERE PHA_NATURERUB<>"AAA" AND PHA_RUBRIQUE IN (SELECT PCT_RUBRIQUE FROM COTISATION'
                            +' COT WHERE ##COT.PCT_PREDEFINI## COT.PCT_THEMECOT="RDC")');
ExecuteSQLContOnExcept('UPDATE HISTOBULLETIN SET PHB_OMTSALARIAL=(SELECT PCT_ORDREETAT FROM COTISATION '
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
    ExecuteSQLContOnExcept('UPDATE STKREGLEGESTION SET GSR_ORDRESELECT="-GQD_REFAFFECTATION;+GQD_DATEENTREELOT;+GQD_DATEPEREMPTION;+GQD_INDICEARTICLE;+GQD_LOTINTERNE;+GQD_SERIEINTERNE;+GQD_EMPLACEMENT" WHERE GSR_REGLEGESTION="FIFO"');
    ExecuteSQLContOnExcept('UPDATE STKREGLEGESTION SET GSR_ORDRESELECT="-GQD_REFAFFECTATION;-GQD_DATEENTREELOT;+GQD_DATEPEREMPTION;+GQD_INDICEARTICLE;+GQD_LOTINTERNE;+GQD_SERIEINTERNE;+GQD_EMPLACEMENT" WHERE GSR_REGLEGESTION="LIFO"');
    ExecuteSQLContOnExcept('UPDATE STKREGLEGESTION SET GSR_ORDRESELECT="-GQD_REFAFFECTATION;+GQD_DATEPEREMPTION;+GQD_DATEENTREELOT;+GQD_INDICEARTICLE;+GQD_LOTINTERNE;+GQD_SERIEINTERNE;+GQD_EMPLACEMENT" WHERE GSR_REGLEGESTION="FEFO"');

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
    ExecuteSQLContOnExcept('UPDATE CHOIXCOD SET CC_TYPE="ZLA" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "A%"' );
    // changé DBR le 29082005
    ExecuteSQLContOnExcept('UPDATE CHOIXCOD SET CC_TYPE="ZLT" WHERE CC_TYPE="ZLI" AND ((CC_CODE LIKE "C%" AND CC_CODE NOT LIKE "CP%") OR CC_CODE LIKE "F%")');

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
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPOPICKING="LBR;BLQ;PER;" WHERE GSN_QUALIFMVT="SEX"');
    //GRC
    RT_InsertTablettesSAV;
    ExecuteSQLContOnExcept('DELETE FROM CHOIXCOD WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "W%"');
    W_SAV_ArticlesParc;
  end;

  //SYSTEM
  ExecuteSQLContOnExcept('UPDATE UTILISAT SET US_GRPSDELEGUES="" where US_GRPSDELEGUES IS NULL');

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
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SDISPOPICKING="AFF;", GSN_SFLUXPICKING="STD;PVE;" WHERE GSN_QUALIFMVT="SCC"');
    //GC
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_LIBELLE="Sortie exceptionnelle" where gsn_qualifmvt="SEX"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_LIBELLE="Entrée exceptionnelle" where gsn_qualifmvt="EEX"');
  end;
End;

Procedure MajVer709 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    AGLNettoieListes('WCBNPREVTET','WPT_AFFAIRE',nil);
    ExecuteSQLContOnExcept('UPDATE CHOIXCOD SET CC_TYPE="GED" WHERE CC_TYPE="GCE"');
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
      ExecuteSQLContOnExcept('DELETE FROM PARPIECE WHERE GPP_NATUREPIECEG LIKE "W%"');
      ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_INSERTLIG="X",GPP_ACTIONFINI="ENR",GPP_STKQUALIFMVT="STR",GPP_SENSPIECE="SOR" WHERE GPP_NATUREPIECEG = "TEM"');
      ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_INSERTLIG="-",GPP_ACTIONFINI="ENR",GPP_STKQUALIFMVT="ETR",GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG = "TRE"');
      ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_INSERTLIG="-",GPP_ACTIONFINI="TRA",GPP_STKQUALIFMVT="ETR",GPP_SENSPIECE="ENT" WHERE GPP_NATUREPIECEG = "TRV"');
      ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_CONTEXTES="GP" WHERE GPP_NATUREPIECEG IN ("BSA","BSP","CSA","CSP")');
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
      (* NON BTP
      CallSTKMoulinetteContreMarque(True);
      *)
      if V_PGI.SAV then LogAGL('Fin Init Stock ContreMarque' + DateTimeToStr(Now));
    end;

    AGLNettoieListesPlus('GCMULMODIFDOCACH','GP_CDETYPE',nil, True);
    AGLNettoieListesPlus('GCMULMODIFDOCVEN','GP_CDETYPE',nil, True);
    AGLNettoieListesPlus('GCMULLIGNEACH','GL_TYPECADENCE',nil, True);
    AGLNettoieListesPlus('GCDUPLICPIECE','GP_CDETYPE',nil, True);

    {-- MISE A JOUR DE STKNATURE --}
    ExecuteSQLContOnExcept('DELETE FROM STKNATURE WHERE GSN_QUALIFMVT="TAV"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_SIGNEMVT="SOR" WHERE GSN_QUALIFMVT="STR"');
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
      ExecuteSQLContOnExcept('UPDATE parpiece set gpp_contextes="AFF;GCA;" where gpp_naturepieceg in ("AFF","PAF","FPR","APR")');
      ExecuteSQLContOnExcept('UPDATE parpiece set gpp_contextes="AFF;GC;" where gpp_naturepieceg in ("FF","AF","FAC","CF","BLF","AVC")');
      ExecuteSQLContOnExcept('UPDATE parpiece set gpp_contextes="GC;" where gpp_naturepieceg in ("AFP","AFS","AVS", "BFA","BLC","BRC","CC","CCE","DE","FFA","FFO","FRA","LCE","PRE","PRF","PRO","REA","TEM","TRE","TRV")');

      If IsMonoOuCommune and (V_PGI.ModePCL='1') then
      begin // mise à jour table tiers si base 00 et PCL
        ExecuteSQLContOnExcept('update tiers set t_tiers=t_auxiliaire where t_tiers=""');
        ExecuteSQLContOnExcept('update tierscompl set ytc_tiers=ytc_auxiliaire where ytc_tiers=""');
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
{$IFNDEF BTP}
    Q := OpenSQL('SELECT COUNT(*)  FROM LIGNE WHERE GL_ARTICLE<>""', True);
    try
      if (not Q.Eof) and (Q.Fields[0].AsInteger > 0) and (Q.Fields[0].AsInteger < 2000000) then
        { On lance le reformatage des ligne uniquement si on a moins de 2millions de lignes}
        MoulinetteFormatageChampsLienPiece(True);
    finally
      Ferme(Q);
    end;
{$ENDIF}
    ExecuteSQLContOnExcept('UPDATE ANNULIEN SET ANL_RACINE="PFC" WHERE ANL_FONCTION="PFC"');
  end;
//PAIE
AglNettoieListes('PGMULSALARIE', 'PSA_ETABLISSEMENT;PSA_SALARIE;PSA_CONFIDENTIEL;PSA_LIBELLE;PSA_PRENOM;PSA_DATEENTREE;PSA_DATESORTIE',Nil);
ExecuteSQLContOnExcept('UPDATE SALARIES SET PSA_CATBILAN="000" WHERE (PSA_CATBILAN="" OR PSA_CATBILAN IS NULL)');
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
ExecuteSQLContOnExcept('UPDATE STDCPTA SET STC_DATEMODIF="' + UsdateTime(iDate1900) + '" WHERE STC_DATEMODIF IS NULL')
End;



Procedure MajVer714 ;
var S : String;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSQLContOnExcept('update parpiece set gpp_qualifmvt="ERC" where gpp_naturepieceg="AVS"');
    ExecuteSQLContOnExcept('update parpiece set gpp_qualifmvt="ERC" where gpp_naturepieceg="BRC"');
    ExecuteSQLContOnExcept('update parpiece set gpp_qualifmvt="" where gpp_naturepieceg="FRA"');
    ExecuteSQLContOnExcept('update parpiece set gpp_qualifmvt="" where gpp_naturepieceg="LCE"');
    ExecuteSQLContOnExcept('update parpiece set gpp_qualifmvt="SRC" where gpp_naturepieceg="BFA"');
    //
    //ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE=DH_CONTROLE||"C" WHERE DH_PREFIXE = "GA" AND DH_CONTROLE LIKE "LD%"');
    //ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE=DH_CONTROLE||"C" WHERE DH_PREFIXE = "ARS" AND DH_CONTROLE LIKE "LD%"');
    //ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE=DH_CONTROLE||"C" WHERE DH_PREFIXE = "GL" AND DH_CONTROLE LIKE "LD%"');
  end;
//Compta
S := 'E_AUXILIAIRE;E_AFFAIRE;T_LIBELLE;T_SCORERELANCE;T_RELANCETRAITE;T_RELANCEREGLEMENT;E_NUMEROPIECE;E_MODEPAIE;E_DATEECHEANCE;E_DATERELANCE;E_QUALIFORIGINE;E_COUVERTUREDEV;E_DEBITDEV;E_CREDITDEV;';
S := S + 'E_DEBIT;E_CREDIT;E_COUVERTURE;E_DATECOMPTABLE;E_JOURNAL;E_EXERCICE;E_QUALIFPIECE;E_NATUREPIECE;';
S := S + 'E_NUMLIGNE;E_NUMECHE;E_DATEMODIF;E_ETABLISSEMENT;E_DEVISE;E_DATETAUXDEV;E_TAUXDEV;E_VALIDE;E_MODESAISIE;E_NIVEAURELANCE;E_BLOCNOTE;';

AglNettoieListes('CPRELANCECLIENT', S);
ExecuteSQL('update menu set mn_versiondev="-" where mn_tag=26095');    // activation lien OLE forcé à autorisé 28102005
// BTP remise à jour tablette AFETATAFFAIRE
ExecuteSQL('DELETE FROM CHOIXCOD WHERE CC_TYPE = "AET"');
InsertChoixCode('AET', 'ACA', 'Accepté', 'Accepté', '3');
InsertChoixCode('AET', 'ACD', 'Attente acceptation devis', 'Attente accept', '63');
InsertChoixCode('AET', 'ACP', 'Accepté', 'Accepté', 'BTP');
InsertChoixCode('AET', 'AFF', 'Affecté', 'Affecté', '95');
InsertChoixCode('AET', 'ANN', 'Annulé', 'Annulé', '76');
InsertChoixCode('AET', 'ECO', 'En cours', 'En cours', '92');
InsertChoixCode('AET', 'ENC', 'En cours', 'En cours', 'BTP');
InsertChoixCode('AET', 'FAC', 'Facturé', 'Facturé', '15');
InsertChoixCode('AET', 'FIN', 'Terminé', 'Terminé', '74');
InsertChoixCode('AET', 'REA', 'Réalisé', 'Réalisé', '79');
InsertChoixCode('AET', 'TER', 'Terminé', 'Terminé', 'BTP');
// -
ExecuteSQL('UPDATE AFFAIRE SET AFF_CHANTIER="" WHERE AFF_CHANTIER IS NULL');
ExecuteSQL('UPDATE CONSOMMATIONS SET BCO_AFFAIRESAISIE="",BCO_QTEFACTUREE=BCO_QUANTITE WHERE BCO_AFFAIRESAISIE IS NULL');
ExecuteSQL('UPDATE TIERS SET T_CLETELEPHONE="",T_CLETELEPHONE2="",T_CLETELEPHONE3="",T_CLEFAX="" WHERE T_CLETELEPHONE2 IS NULL');
ExecuteSQL('UPDATE ADRESSES SET ADR_CLETELEPHONE="",ADR_INT="-" WHERE ADR_CLETELEPHONE IS NULL');
ExecuteSQL('UPDATE CONTACT SET C_CLETELEPHONE="",C_CLEFAX="",C_CLETELEX="" WHERE C_CLEFAX IS NULL');
// MAJ V6.53
ExecuteSQL('UPDATE CONSOMMATIONS SET BCO_INTEGRNUMFIC=0 WHERE BCO_INTEGRNUMFIC IS NULL');
End;

Procedure MajVer715 ;
var S : String;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  ExecuteSql ('UPDATE PARPIECE SET GPP_FORCEENPA="-"');
  ExecuteSql ('UPDATE PARPIECE SET GPP_FORCEENPA="X" WHERE GPP_NATUREPIECEG IN ("CBT","PBT","LBT","ETU")');
  ExecuteSql ('UPDATE PARPIECE SET GPP_SOUCHE="PBT" WHERE GPP_NATUREPIECEG="PBT"');
  ExecuteSql ('UPDATE PARPIECE SET GPP_SOUCHE="FRC" WHERE GPP_NATUREPIECEG="FRC"');
  ExecuteSql ('UPDATE LIGNECOMPL SET GLC_NONAPPLICFRAIS="-"');
  ExecuteSql ('UPDATE LIGNEOUV SET BLO_NONAPPLICFRAIS="-"');
  ExecuteSql ('UPDATE PIECE SET GP_PIECEFRAIS="", GP_APPLICFGST="-", GP_PIECEENPA="-"');
  ExecuteSql ('UPDATE PIECE SET GP_PIECEENPA="X" WHERE GP_NATUREPIECEG IN ("PBT","CBT","LBT","ETU")');
  // mise a jour des compteurs
  ExecuteSql ('UPDATE SOUCHE SET SH_NUMDEPART=(SELECT MAX(GP_NUMERO) FROM PIECE WHERE GP_NATUREPIECEG="FRC") WHERE SH_TYPE="GES" AND SH_SOUCHE="FRC"');
  ExecuteSql ('UPDATE SOUCHE SET SH_NUMDEPART=(SELECT MAX(GP_NUMERO) FROM PIECE WHERE GP_NATUREPIECEG="PBT") WHERE SH_TYPE="GES" AND SH_SOUCHE="PBT"');
  // ILS ONT RAJOUTE CA PENDANT LA SORTIE DE LA 714 ...mffff
  if not (isOracle) then
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
  ExecuteSQLContOnExcept('UPDATE QMATIERE SET QMT_FAMILLENIV1="",QMT_FAMILLENIV2="",QMT_FAMILLENIV3="" WHERE QMT_FAMILLENIV1 IS NULL');
  ExecuteSQLContOnExcept('UPDATE QARTTECH SET QAR_FAMILLENIV1="",QAR_FAMILLENIV2="",QAR_FAMILLENIV3="",QAR_FAMILLEMAT="" WHERE QAR_FAMILLENIV1 IS NULL');
  ExecuteSQLContOnExcept('UPDATE QMAGAPP  SET QMP_MATIERE="", QMP_FAMILLENIV1="",QMP_FAMILLENIV2="",QMP_FAMILLENIV3="" WHERE QMP_MATIERE IS NULL');
  ExecuteSQLContOnExcept('UPDATE QSTKCOLT SET QSK_STOCKRECOMPL=0,QSK_UNITEMIN="",QSK_UNITEALERTE="",QSK_UNITEMAX="",QSK_UNITERECOMPL="",QSK_DELAIMOYAPPRO=0,QSK_TAILLELOT=0,QSK_QTEMINFAB=0 WHERE QSK_STOCKRECOMPL IS NULL');
  ExecuteSQLContOnExcept('UPDATE QAPPROS SET QA_REFCOMMANDE="" WHERE QA_REFCOMMANDE IS NULL');
  ExecuteSQLContOnExcept('UPDATE QSTKDISPO SET QSD_REFCOMMANDE="" WHERE QSD_REFCOMMANDE IS NULL');
  ExecuteSQLContOnExcept('UPDATE WORDRELIG SET WOL_LIGNEORDREGEN=0 WHERE WOL_LIGNEORDREGEN IS NULL');
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
    ExecuteSQLContOnExcept('UPDATE WORDREBES SET WOB_SURCONSO="X",WOB_CONSOTOT="-" WHERE WOB_SURCONSO IS NULL');

  end;
End;

Procedure MajVer723 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //Maj_GSM_GUIDORI;
    ExecuteSQLContOnExcept('UPDATE LIGNECOMPL SET GLC_NBINTERVENANT=0 WHERE GLC_NBINTERVENANT IS NULL');
  end;
End;

Procedure MajVer724 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //SAV
    ExecuteSQLContOnExcept('UPDATE WPARCVERSION SET WPV_ARTICLEPARC=(SELECT WPN_ARTICLEPARC FROM WPARCNOME WHERE WPN_IDENTIFIANT=WPV_IDENTIFIANT)');
    //GPAO
    ExecuteSQLContOnExcept('UPDATE WNATURETRAVAIL SET WNA_IMPMODELEWOL="OR1", WNA_APERCUWOL="X" where WNA_IMPMODELEWOL IS NULL');
    { Mise à jour des familles de messages EDI }
    { Bons de réception }
    if not ExisteSQL('SELECT 1 FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="BLF"') then
    ExecuteSQLContOnExcept('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE       , EFM_NATURESPIECE, EFM_SERIALISE)'
            +                  ' VALUES ("BLF"          , "Bon de réception", "PRF;BLF"       , "X"          )');

    { Commandes d'achat }
    if not ExisteSQL('SELECT 1 FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="CDA"') then
    ExecuteSQLContOnExcept('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE     , EFM_NATURESPIECE, EFM_SERIALISE)'
           +                   ' VALUES ("CDA"          , "Commande achat", "CF"            , "X"          )');

    ExecuteSQLContOnExcept('update TIERSCOMPL  set YTC_MODELEBON="" where YTC_MODELEBON is null');
    ExecuteSQLContOnExcept('update TIERSCOMPL  set YTC_EDITRA="" where YTC_EDITRA is null');
    ExecuteSQLContOnExcept('update TIERSCOMPL  set YTC_MODELETXT="" where YTC_MODELETXT is null');
    ExecuteSQLContOnExcept('update TIERSCOMPL  set YTC_TIERSEXPE="" where YTC_TIERSEXPE is null');
    ExecuteSQLContOnExcept('update TIERSCOMPL  set YTC_QUALIFPOIDS="" where YTC_QUALIFPOIDS is null');
  end;
End;


Procedure MajVer725 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSQLContOnExcept('UPDATE PAYS SET PY_LIBELLE = "HAITI", PY_ABREGE = "HAITI" WHERE PY_PAYS = "HAI"');
    ExecuteSQLContOnExcept('UPDATE PAYS SET PY_LIBELLE = "GUADELOUPE", PY_ABREGE = "GUADELOUPE" WHERE PY_PAYS = "GOU"');
    //GP
    ExecuteSQLContOnExcept('UPDATE TIERSCOMPL SET YTC_DEPOT="", YTC_ETABLISSEMENT="" WHERE YTC_DEPOT IS NULL');
  end;
End;


Procedure MajVer726 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    // GCAIX
    ExecuteSQLContOnExcept('update TIERSCOMPL  set YTC_MODELEBON="",YTC_EDITRA="-",YTC_MODELETXT="",YTC_TIERSEXPE="",YTC_QUALIFPOIDS="" where YTC_MODELEBON is null and  YTC_EDITRA is null and YTC_MODELETXT is null and YTC_TIERSEXPE is null and YTC_QUALIFPOIDS is null');
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
    if not ExisteSQL('SELECT WDY_NOMCHAMP FROM WFORMULECHAMPDEF WHERE WDY_CONTEXTE="WOR" AND WDY_NOMCHAMP="WOR_COEFPROD"') then
      ExecuteSQLContOnExcept('INSERT INTO WFORMULECHAMPDEF '
        + ' (WDY_CONTEXTE, WDY_NOMCHAMP, WDY_TYPECHAMP, WDY_DATECREATION, WDY_DATEMODIF, WDY_CREATEUR, WDY_UTILISATEUR) '
        + ' VALUES ("WOR","WOR_COEFPROD","COE","'+ USTime(Nowh)+'","'+UsTime(Nowh)+'","'+V_PGI.User+'","'+V_PGI.User+'")' );

    if not ExisteSQL('SELECT WDY_NOMCHAMP FROM WFORMULECHAMPDEF WHERE WDY_CONTEXTE="WOR" AND WDY_NOMCHAMP="WOR_QPRODSAIS"') then
      ExecuteSQLContOnExcept('INSERT INTO WFORMULECHAMPDEF (WDY_CONTEXTE, WDY_NOMCHAMP, WDY_TYPECHAMP, '
        + ' WDY_DATECREATION, WDY_DATEMODIF, WDY_CREATEUR, WDY_UTILISATEUR) '
        + ' VALUES ("WOR","WOR_QPRODSAIS","QTE","'+ USTime(Nowh)+'","'+UsTime(Nowh)+'","'+V_PGI.User+'","'+V_PGI.User+'")' );

    if not ExisteSQL('SELECT WDY_NOMCHAMP FROM WFORMULECHAMPDEF WHERE WDY_CONTEXTE="WOR" AND WDY_NOMCHAMP="WOR_QPRODSAIS"') then
      ExecuteSQLContOnExcept('INSERT INTO WFORMULECHAMPDEF '
        + ' (WDY_CONTEXTE, WDY_NOMCHAMP, WDY_TYPECHAMP, WDY_DATECREATION, WDY_DATEMODIF, WDY_CREATEUR, WDY_UTILISATEUR) '
        + ' Values ("WOR","WOR_QPRODSTOC","QTE","'+ USTime(Nowh)+'","'+UsTime(Nowh)+'","'+V_PGI.User+'","'+V_PGI.User+'")' );

    if not ExisteSQL('SELECT WDY_NOMCHAMP FROM WFORMULECHAMPDEF WHERE WDY_CONTEXTE="WOR" AND WDY_NOMCHAMP="WOR_QUALIFUNITESTO"') then
      ExecuteSQLContOnExcept('INSERT INTO WFORMULECHAMPDEF '
        + ' (WDY_CONTEXTE, WDY_NOMCHAMP, WDY_TYPECHAMP, WDY_DATECREATION, WDY_DATEMODIF, WDY_CREATEUR, WDY_UTILISATEUR) '
        + ' VALUES ("WOR","WOR_QUALIFUNITESTO","UNI","'+ USTime(Nowh)+'","'+UsTime(Nowh)+'","'+V_PGI.User+'","'+V_PGI.User+'")' );

    if not ExisteSQL('SELECT WDY_NOMCHAMP FROM WFORMULECHAMPDEF WHERE WDY_CONTEXTE="WOR" AND WDY_NOMCHAMP="WOR_UNITEPROD"') then
      ExecuteSQLContOnExcept('INSERT INTO WFORMULECHAMPDEF '
        + ' (WDY_CONTEXTE, WDY_NOMCHAMP, WDY_TYPECHAMP, WDY_DATECREATION, WDY_DATEMODIF, WDY_CREATEUR, WDY_UTILISATEUR) '
        + ' VALUES ("WOR","WOR_UNITEPROD","UNI","'+ USTime(Nowh)+'","'+UsTime(Nowh)+'","'+V_PGI.User+'","'+V_PGI.User+'")' );

    if not ExisteSQL('SELECT WDX_CONTEXTE FROM WFORMULECHAMP WHERE WDX_CONTEXTE="WOR" AND WDX_QUALIFIANT="QPR"') then
      ExecuteSQLContOnExcept('INSERT INTO WFORMULECHAMP '
        + ' (WDX_CONTEXTE, WDX_QUALIFIANT, WDX_CHAMPSAIS, WDX_UNITECHAMPSAIS,'
        + 'WDX_CHAMPCOEF, WDX_CHAMPCALC, WDX_UNITECHAMPCALC, WDX_DATECREATION, WDX_DATEMODIF, WDX_CREATEUR, WDX_UTILISATEUR) '
        + ' VALUES ("WOR","QPR","WOR_QPRODSTOC","WOR_QUALIFUNITESTO","WOR_COEFPROD",'
        + '"WOR_QPRODSAIS","WOR_UNITEPROD","'
        + USTime(Nowh)+'","'+UsTime(Nowh)+'","'+V_PGI.User+'","'+V_PGI.User+'")' );

    ExecuteSQLContOnExcept('UPDATE WORDRERES'
      + ' SET WOR_RESORIGINEMES="", WOR_NBRESSOURCE=1,WOR_UNITEPROD= (SELECT WOG_UNITEACC FROM WORDREGAMME'
      + ' WHERE WOG_NATURETRAVAIL=WOR_NATURETRAVAIL AND WOG_LIGNEORDRE=WOR_LIGNEORDRE'
      + ' AND WOG_OPECIRC=WOR_OPECIRC AND WOG_NUMOPERGAMME=WOR_NUMOPERGAMME),'
      + ' WOR_QUALIFUNITESTO= (SELECT WOG_QUALIFUNITESTO FROM WORDREGAMME'
      + ' WHERE WOG_NATURETRAVAIL=WOR_NATURETRAVAIL AND WOG_LIGNEORDRE=WOR_LIGNEORDRE'
      + ' AND WOG_OPECIRC=WOR_OPECIRC AND WOG_NUMOPERGAMME=WOR_NUMOPERGAMME),WOR_COEFPROD=1,WOR_QPRODSAIS=0,WOR_QPRODSTOC=0');

    ExecuteSql ('UPDATE ARTICLE SET GA_GMARQUE="-", GA_GCHOIXQUALITE="-" WHERE GA_GMARQUE IS NULL');

    ExecuteSQLContOnExcept('UPDATE CATALOGU SET GCA_LIBREGCA1="",GCA_LIBREGCA2="",GCA_LIBREGCA3="",GCA_LIBREGCA4="",GCA_LIBREGCA5="",GCA_LIBREGCA6="",GCA_LIBREGCA7="",GCA_LIBREGCA8="",GCA_LIBREGCA9="",GCA_LIBREGCAA="",'
                             + 'GCA_CHARLIBRE1="",GCA_CHARLIBRE2="",GCA_CHARLIBRE3="",'
                             + 'GCA_DATELIBRE1="' + UsDateTime(iDate1900) + '",GCA_DATELIBRE2="' + UsDateTime(iDate1900) + '",GCA_DATELIBRE3="' + UsDateTime(iDate1900) + '",'
                             + 'GCA_VALLIBRE1=0,GCA_VALLIBRE2=0,GCA_VALLIBRE3=0,GCA_BOOLLIBRE1="-",GCA_BOOLLIBRE2="-",GCA_BOOLLIBRE3="-" WHERE GCA_LIBREGCA1 IS NULL');


    ExecuteSQLContOnExcept('UPDATE LIGNECOMPL SET GLC_MARQUE="", GLC_CHOIXQUALITE="" WHERE GLC_MARQUE IS NULL');

    MajMarkMESApres;
    //TRESO
    AglNettoieListes('TRLISTEVIR', 'TEQ_PAYSDEST;TEQ_PAYSSOURCE;TEQ_FICVIR;TEQ_DEVISE;', nil);
    //GRC
    AglNettoieListes('WARTICLEPARC_GENE','WAP_DUREEGARANTIE',nil);
    AglNettoieListes('WPARCSELECT','WPC_IDENTIFIANT;WPC_TIERS',nil);
    //GPAO
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_NATURESUIVANTE="BSA;", GPP_CONTEXTES="GC;"  WHERE GPP_NATUREPIECEG="CSA"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_CONTEXTES="GC;"  WHERE GPP_NATUREPIECEG="BSA"');

    ExecuteSQLContOnExcept('UPDATE STKFICHETRACE SET GST_MARQUE = "" WHERE GST_MARQUE IS NULL');
    ExecuteSQLContOnExcept('UPDATE STKFICHETRACE SET GST_CHOIXQUALITE = "" WHERE GST_CHOIXQUALITE IS NULL');

    ExecuteSQLContOnExcept('UPDATE LIGNECOMPL SET GLC_MARQUE = "" WHERE GLC_MARQUE IS NULL');
    ExecuteSQLContOnExcept('UPDATE LIGNECOMPL SET GLC_CHOIXQUALITE = "" WHERE GLC_CHOIXQUALITE IS NULL');

    if not existeSql('SELECT 1 FROM STKNATURE WHERE GSN_QUALIFMVT="CMQ"') then
    begin
        ExecuteSQLContOnExcept('INSERT INTO STKNATURE (GSN_QUALIFMVT, GSN_LIBELLE, GSN_STKTYPEMVT, GSN_QTEPLUS, GSN_QUALIFMVTSUIV, GSN_SIGNEMVT,'
                 + ' GSN_STKFLUX, GSN_GERECOMPTA, GSN_CALLGSL, GSN_CALLGSS, GSN_CONTREMARQUE, GSN_MAJPRIXVALO, GSN_SDISPODISPATCH, GSN_SFLUXDISPATCH, GSN_SDISPOPICKING, GSN_SFLUXPICKING)'
                 + ' VALUES ("CMQ", "Changement de marque", "CET", "", "", "MIX",'
                 + ' "STO", "-", "-", "-", "-", "", "", "", "LBR;", "STD;")');
    end;

    if not existeSql('SELECT 1 FROM STKNATURE WHERE GSN_QUALIFMVT="CCQ"') then
    begin
        ExecuteSQLContOnExcept('INSERT INTO STKNATURE (GSN_QUALIFMVT, GSN_LIBELLE, GSN_STKTYPEMVT, GSN_QTEPLUS, GSN_QUALIFMVTSUIV, GSN_SIGNEMVT,'
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
    //  ExecuteSQLContOnExcept('UPDATE LIGNECOMPL SET GLC_DATEDEB="'+UsDateTime(iDate1900)+'",GLC_DATEFIN="'+UsDateTime(iDate1900)+'",GLC_NUMEROMARCHE=0,GLC_CODEMARCHE="" WHERE GLC_CODEMARCHE IS NULL');
    AglNettoieListes('GCLISTINVLIGNE'  , 'GM_LIBLLE;GCQ_LIBELLE;GIL_REFAFFECATION',nil);

    {
    Mise en commentaire suite à la Demande N° 1881
    //GIGA
    AglNettoieListes('AFPLALIGNE_MUL', 'GLC_AFFAIRELIE',nil);
    }
  end;

  ExecuteSQLContOnExcept('UPDATE UTILISAT SET US_EMAILSMTPLOGIN="", US_EMAILSMTPPWD ="", US_POPPORT="", US_SMTPPORT=""');

  //ajout 01/02/2006
  ExecuteSQLContOnExcept('UPDATE IMMO SET I_OPEDEPREC="-",I_REPRISEDEP=0.00,I_ECCLEECR="",I_DOCGUID="",I_DATEFINCB="'+UsDateTime(iDate1900)+'"');

End;


Procedure MajVer729 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    InsertChoixCode('CVU', '003', 'Conversion d''unité', 'AdvConvU', '');
    ExecuteSQLContOnExcept('UPDATE ARTICLE SET GA_UNITEPRIXACH=ISNULL(GA_QUALIFUNITEVTE,""),GA_COEFCONVQTEACH=1,GA_COEFCONVQTEVTE=1,GA_COEFCONVCONSO=1 WHERE GA_COEFCONVQTEACH IS NULL');
    ExecuteSQLContOnExcept('UPDATE CATALOGU SET GCA_UNITEQTEACH=GCA_QUALIFUNITEACH,GCA_COEFCONVQTEACH=1 WHERE GCA_UNITEQTEACH IS NULL');
    UpdateDecoupeLigne('GL_UNITEPRIX=GL_QUALIFQTESTO', 'AND (ISNULL(GL_ARTICLE, " ") <> " ")');
    ExecuteSQLContOnExcept('UPDATE WCBNEVOLUTION SET WEV_MARQUE="",WEV_CODEMARCHE="",WEV_NUMEROMARCHE=0,WEV_QUOTAPCTDEM=0,WEV_QUOTAPCTAPP=0 WHERE WEV_MARQUE IS NULL');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_TARIFGENSAISIE="010", GPP_TARIFGENTRANSF="010", GPP_TARIFGENDATE="010", GPP_TARIFGENDEPOT="-", GPP_TARIFGENSPECIA="010" WHERE GPP_TARIFGENSAISIE IS NULL');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_TARIFGENSAISIE="020", GPP_TARIFGENTRANSF="010", GPP_TARIFGENDATE="020", GPP_TARIFGENDEPOT="-", GPP_TARIFGENSPECIA="040" WHERE GPP_NATUREPIECEG="CMF"');
    ExecuteSQLContOnExcept('UPDATE YTARIFSPARAMETRES SET YFO_OKTARIFMARCHE = "----" WHERE YFO_OKTARIFMARCHE IS NULL');
    ExecuteSQLContOnExcept('UPDATE YTARIFS SET YTS_TARIFMARCHE="" WHERE YTS_TARIFMARCHE IS NULL');
    ExecuteSQLContOnExcept('UPDATE YTARIFS SET YTS_FORMULEMT="" WHERE YTS_FORMULEMT IS NULL');
    ExecuteSQLContOnExcept('UPDATE WPARAMFONCTION SET WPF_FORMULEMT="" WHERE  WPF_FORMULEMT IS NULL');
    { Nouvelle nature de stock permettant de gérer les unités logistiques }
    if not ExisteSQL('SELECT 1 FROM STKNATURE WHERE GSN_QUALIFMVT="SCO"') then
    begin
      ExecuteSQLContOnExcept('INSERT INTO STKNATURE (GSN_CALLGSL, GSN_CALLGSS, GSN_CONTREMARQUE, GSN_GERECOMPTA, GSN_LIBELLE, '
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
    ExecuteSQLContOnExcept('UPDATE LIGNECOMPL SET GLC_MODIFPRIXACHAT="-" WHERE GLC_MODIFPRIXACHAT IS NULL');
  end;
End;

Procedure MajVer730 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GPAO
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_NATURESUIVANTE="" WHERE GPP_NATUREPIECEG="CMF"');
    ExecuteSQLContOnExcept('UPDATE WORDRELAS SET WLS_PHASE="" WHERE WLS_PHASE IS NULL');
    ExecuteSQLContOnExcept('UPDATE QCIRCUIT SET QCI_COMAGENT=0 WHERE QCI_COMAGENT IS NULL');
    ExecuteSQLContOnExcept('UPDATE STKVALOTYPE SET GVT_VALOPRIXCOMP = "" WHERE  GVT_VALOPRIXCOMP IS NULL');
    ExecuteSQLContOnExcept('UPDATE CATALOGU SET GCA_QECOACH=0,GCA_QPCBACH=0 WHERE GCA_QECOACH IS NULL');
    //AIX
    { Init Nouvelles Zones }
    ExecuteSQLContOnExcept('update TIERSCOMPL  set YTC_MODELE="",YTC_AUTOEDITETIQ="-"'
             +',YTC_AUTOEDITOT ="-",YTC_TYPEECHANGE="",YTC_PATHFICPRIVE=""'
             +',YTC_AUTOGENEPRIVE="-",YTC_SOUCHEOT="",YTC_CODEPRODTRA=""'
             +',YTC_QUALIFLINEAIRE="",YTC_QUALIFVOLUME="" '
             +' where YTC_MODELE is null and YTC_AUTOEDITETIQ is null '
             +'       and YTC_AUTOEDITOT is null and YTC_TYPEECHANGE is null '
             +'       and YTC_PATHFICPRIVE is null and YTC_AUTOGENEPRIVE is null '
             +'       and YTC_SOUCHEOT is null and YTC_CODEPRODTRA is null '
             +'       and YTC_QUALIFLINEAIRE is Null and YTC_QUALIFVOLUME is Null');
    {Maj YTC_TYPEECHANGE Si YTC_EDITRA est True}
    ExecuteSQLContOnExcept('update TIERSCOMPL  set YTC_TYPEECHANGE="PRI",YTC_AUTOGENEPRIVE="X" where YTC_EDITRA = "X" ');
    { Init Nouvelle Zone Régime transport }
    ExecuteSQLContOnExcept('update gcproduittra set gtc_regimetra="" ');

    // MB : PCL : On ne devrais plus passer par là, on le fera plus tard ...
    {if IsMonoOuCommune and (Not ExisteSQL('SELECT 1 FROM DOSSIERGRP')) THEN
        ExecuteSQL ('INSERT INTO DOSSIERGRP (DOG_NODOSSIER,DOG_GROUPECONF) SELECT DOS_NODOSSIER AS DOG_NODOSSIER,DOS_GROUPECONF AS DOG_GROUPECONF FROM DOSSIER')}

end;

ExecuteSQL ('UPDATE SECTION SET S_LIBELLE = "FRANCE METRO. - TAUX 19,6 %" WHERE S_SECTION = "0206" AND S_CREERPAR ="TVA"');
if (GetParamSocSecur('SO_CPPCLSAISIETVA', '')) and (GetParamSocSecur('SO_CPPCLAXETVA', '') = '') then
    begin
          SetParamSoc('SO_CPPCLAXETVA', GetColonneSQL('AXE', 'X_AXE', 'X_LIBELLE = "TVA"'));
          InsertChoixCode('ATU', '038', 'des tiers', '', 'des clients');
    end;
ExecuteSQLContOnExcept('update dechamps set dh_controle="L" where dh_nomchamp="GU_REFRELEVE" or dh_nomchamp="GU_TRESORERIE"');

End;

Procedure MajVer734 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    // MES
    if not ExisteSQL('SELECT 1 from QOFAMILLEALEA where QOF_FAMILLEALEA="OP"') then
      ExecuteSQLContOnExcept('INSERT INTO QOFAMILLEALEA '
            + ' (QOF_FAMILLEALEA, QOF_LIBELLE, QOF_ALEADEDUCTIBLE, QOF_ALEAPRIME, QOF_ALEARATIO) '
            + ' VALUES ("OP","Opérations","X","X","-")' );
    if not ExisteSQL('SELECT 1 from QOALEA where QOA_ALEA="ALEAOP"') then
      ExecuteSQLContOnExcept('INSERT INTO QOALEA '
            + ' (QOA_ALEA,QOA_FAMILLEALEA,QOA_LIBELLE,QOA_ACTIF,QOA_ALEADEDUCTIBLE,QOA_DATEDEBCONST,QOA_DATEFINCONST, '
            +   'QOA_DATECREATION,QOA_DATEMODIF,QOA_CREATEUR,QOA_UTILISATEUR) '
            + ' VALUES ("ALEAOP","OP","Aléas sur Opérations","X","X","'+UsTime(iDate1900)+'","'+UsTime(iDate1900)+'",'
            + ' "'+UsTime(Nowh)+'","'+UsTime(Nowh)+'","'+V_PGI.User+'","'+V_PGI.User+'")' );
    ExecuteSQLContOnExcept('Update QWHISTORES SET QWH_ORDREWOR=0');
    //GPAO
    ExecuteSQLContOnExcept('UPDATE WORDRELIG SET WOL_DATEEDITFIC = "' + UsDateTime(iDate1900) + '" WHERE WOL_DATEEDITFIC IS NULL');
    ExecuteSQLContOnExcept('UPDATE WORDRELIG SET WOL_DATEEDITBON = "' + UsDateTime(iDate1900) + '" WHERE WOL_DATEEDITBON IS NULL');
    ExecuteSQLContOnExcept('UPDATE WORDRELIG SET WOL_USEREDITFIC = "" WHERE WOL_USEREDITFIC IS NULL');
    ExecuteSQLContOnExcept('UPDATE WORDRELIG SET WOL_USEREDITBON = "" WHERE WOL_USEREDITBON IS NULL');
    //DP
    AGLNettoieListes ('DPDOCUMENTS_MUL','DPD_UTILISATEUR',nil);
    //GEP
  end;

ExecuteSQLContOnExcept('update scriptsp set YSC_QUALIFIANT="SG" where ysc_script not in (select YFM_FORME from formesp) and ysc_typescript="YY"');

End;

Procedure MajVer735 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GPAO
    ExecuteSQLContOnExcept('UPDATE LIGNEFRAIS SET LF_POURCENT1=LF_POURCENT, LF_POURCENT2=0, LF_POURCENT3=0 WHERE LF_POURCENT1 IS NULL');
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
    ExecuteSQLContOnExcept('update YMSGFILES set YMG_ATTACHED="X"');
    // Transfert du contenu Html des mails reçus
    ExecuteSQLContOnExcept('insert into YMSGFILES(YMG_MSGGUID, YMG_MSGID, YMG_FILEGUID, YMG_FILEID, YMG_ATTACHED)'
     + ' select YMS_MSGGUID, YMS_MSGID, YMF_FILEGUID, YMF_FILEID, "-"'
     + ' from YMAILFILES, YMESSAGES'
     + ' where YMS_USERMAIL=YMF_UTILISATEUR AND YMS_MAILID=YMF_MAILID and YMF_ATTACHED="-"');

    //GPAO
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_MAJPRIXVALO="" WHERE GPP_NATUREPIECEG IN ("CSP","BSP") AND GPP_MAJPRIXVALO LIKE "DPA;DPR;PPA;PPR;"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_MAJPRIXVALO="DPA;DPR;" WHERE GPP_NATUREPIECEG = "CSA" AND GPP_MAJPRIXVALO LIKE "DPA;DPR;PPA;PPR;"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="ASA" WHERE GPP_NATUREPIECEG IN ("CSA")');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_STKQUALIFMVT="ESA" WHERE GPP_NATUREPIECEG IN ("BSA")');
    //
    AGLNettoieListes('GCCATALOGUE','GCA_TIERS',nil);

    ExecuteSQLContOnExcept('UPDATE ETABLISS SET ET_EAN="" WHERE ET_EAN IS NULL');

    if not ExisteSQL('SELECT 1 FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="DEL"') then
      ExecuteSQLContOnExcept('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE, EFM_NATURESPIECE, EFM_SERIALISE) VALUES ("DEL", "Plan d''approvisionnement DELFOR", "", "X")');
{$IFNDEF BTP}
    if not ExisteSQL('SELECT 1 FROM STKVALOTYPE where GVT_VALOTYPE="SAA"') then
    begin
      ExecuteSQLContOnExcept('INSERT INTO STKVALOTYPE (GVT_VALOTYPE, GVT_LIBELLE, GVT_VALOPRIX, GVT_VALOPRIXCOMP, GVT_TYPEPDR, GVT_TEMPSREEL, GVT_RECALCUL, GVT_ECARTVALO, GVT_AVECINDIRECT, GVT_CREATEUR, GVT_DATECREATION, GVT_UTILISATEUR, GVT_DATEMODIF)'
                                +' VALUES ("SAA", "Valorisation au dernier prix de sous traitance d''achat", "04", "41", "ACT", "-", "-", "-", "-", "CEG", "'+UsTime(iDate1900)+'", "'+V_PGI.User+'", "'+UsTime(iDate1900)+'")');
    end;
    if not ExisteSQL('SELECT 1 FROM STKNATURE WHERE GSN_QUALIFMVT="ASA"') then
    begin
      ExecuteSQLContOnExcept('INSERT INTO STKNATURE (GSN_CALLGSL, GSN_CALLGSS, GSN_CONTREMARQUE, GSN_GERECOMPTA, GSN_LIBELLE, '
                +                       'GSN_MAJPRIXVALO, GSN_QTEPLUS, GSN_QUALIFMVT, GSN_QUALIFMVTSUIV, '
                +                       'GSN_SDISPODISPATCH, GSN_SDISPOPICKING, GSN_SFLUXDISPATCH, GSN_SFLUXPICKING, '
                +                       'GSN_SIGNEMVT, GSN_STKFLUX, GSN_STKTYPEMVT)'
                + ' VALUES ("-", "-", "-", "-", "Attendu de sous traitance d''achat", "DPA;DPR,", "GQ_RESERVEFOU;", "ASA", "", "", "", "", "", "ENT", "ACH", "ATT")');
    end;

    if not ExisteSQL('SELECT 1 FROM STKNATURE WHERE GSN_QUALIFMVT="ESA"') then
    begin
      ExecuteSQLContOnExcept('INSERT INTO STKNATURE (GSN_CALLGSL, GSN_CALLGSS, GSN_CONTREMARQUE, GSN_GERECOMPTA, GSN_LIBELLE, '
                +                       'GSN_MAJPRIXVALO, GSN_QTEPLUS, GSN_QUALIFMVT, GSN_QUALIFMVTSUIV, '
                +                       'GSN_SDISPODISPATCH, GSN_SDISPOPICKING, GSN_SFLUXDISPATCH, GSN_SFLUXPICKING, '
                +                       'GSN_SIGNEMVT, GSN_STKFLUX, GSN_STKTYPEMVT)'
                + ' VALUES ("-", "-", "-", "-", "Entréee de sous traitance d''achat", "DPA;DPR;PPA;PPR;", "GQ_PHYSIQUE;", "ESA", "", "", "", "", "", "MIX", "ACH", "PHY")');
    end;

//  ExecuteSQLContOnExcept('UPDATE STKMOUVEMENT SET GSM_QUALIFMVT="ASA" WHERE GSM_NATUREORI="CSA"');
//  ExecuteSQLContOnExcept('UPDATE STKMOUVEMENT SET GSM_QUALIFMVT="ESA" WHERE GSM_NATUREORI="BSA"');
    ExecuteSQLContOnExcept('UPDATE STKMOUVEMENT SET GSM_QUALIFMVT="ASA" WHERE GSM_STKTYPEMVT="ATT" AND GSM_QUALIFMVT="AAC" AND GSM_NATUREORI="CSA"');
    ExecuteSQLContOnExcept('UPDATE STKMOUVEMENT SET GSM_QUALIFMVT="ESA" WHERE GSM_STKTYPEMVT="PHY" AND GSM_QUALIFMVT="EAC" AND GSM_NATUREORI="BSA"');

    if not ExisteSQL('SELECT 1 FROM STKVALOPARAM WHERE GVP_QUALIFMVT="ASA"') then
    begin
      ExecuteSQLContOnExcept('INSERT INTO STKVALOPARAM (GVP_IDENTIFIANT, GVP_QUALIFMVT, GVP_NATURETRAVAIL, GVP_STKFLUX, GVP_FAMILLENIV1, GVP_FAMILLENIV2, GVP_FAMILLENIV3, GVP_FAMILLEVALO, GVP_VALOTYPE, GVP_CREATEUR, GVP_DATECREATION, GVP_UTILISATEUR, GVP_DATEMODIF)'
                                + ' VALUES ('+inttostr(wSetJeton('GVP'))+', "ASA", "", "", "", "", "", "", "SAA", "'+V_PGI.User+'", "'+UsTime(iDate1900)+'", "'+V_PGI.User+'", "'+UsTime(iDate1900)+'")');
    end;
    if not ExisteSQL('SELECT 1 FROM STKVALOPARAM WHERE GVP_QUALIFMVT="ESA"') then
    begin
      ExecuteSQLContOnExcept('INSERT INTO STKVALOPARAM (GVP_IDENTIFIANT, GVP_QUALIFMVT, GVP_NATURETRAVAIL, GVP_STKFLUX, GVP_FAMILLENIV1, GVP_FAMILLENIV2, GVP_FAMILLENIV3, GVP_FAMILLEVALO, GVP_VALOTYPE, GVP_CREATEUR, GVP_DATECREATION, GVP_UTILISATEUR, GVP_DATEMODIF)'
                                + ' VALUES ('+inttostr(wSetJeton('GVP'))+', "ESA", "", "", "", "", "", "", "SAA", "'+V_PGI.User+'", "'+UsTime(iDate1900)+'", "'+V_PGI.User+'", "'+UsTime(iDate1900)+'")');
    end;

{$ENDIF}
    ExecuteSQLContOnExcept('UPDATE WORDREBES SET WOB_DPA=0, WOB_DPR=0 WHERE WOB_DPA IS NULL');
{$IFDEF BTP}
    ExecuteSQLContOnExcept('UPDATE STKFORMULE SET GSF_COMPOSANT="",GSF_LIBREART1="",GSF_LIBREART2="",GSF_LIBREART3="",GSF_LIBREART4="",GSF_LIBREART5="",GSF_LIBREART6="",GSF_LIBREART7="",GSF_LIBREART8="",GSF_LIBREART9="",GSF_LIBREARTA="" WHERE GSF_COMPOSANT IS NULL');

    ExecuteSQLContOnExcept('UPDATE STKFICHETRACE SET GST_VALLIBRE1=0,GST_VALLIBRE2=0,GST_VALLIBRE3=0 WHERE GST_VALLIBRE1 IS NULL');

    AglNettoieListes('GCSTKDISPATCH'   , 'GQD_VALLIBRE1;GQD_VALLIBRE2,GQD_VALLIBRE3',nil);
    AglNettoieListes('GCSTKPICKING'    , 'GQD_VALLIBRE1;GQD_VALLIBRE2,GQD_VALLIBRE3',nil);
    AglNettoieListes('GCSTKDISPODETAIL', 'GQD_VALLIBRE1;GQD_VALLIBRE2,GQD_VALLIBRE3',nil);
    AglNettoieListes('GCLISTINVLIGNECTM', 'GIL_REFAFFECTATION',nil);
{$ENDIF}
    //giga
    ExecuteSQLContOnExcept('Update CONTACT set C_GUIDPER='' where C_GUIDPER is null');
    If IsMonoOuCommune then SuppAnnuinterloc; //pasage table annuinterloc dans contact dans base commune
    //fin giga
    //DP
    If IsMonoOuCommune then
    begin
      ExecuteSQLContOnExcept('update YMESSAGES set YMS_STRDATE=""');
     //  ExecuteSQLContOnExcept('update YMESSAGES set YMS_STRDATE=YMA_STRDATE from YMAILS where YMS_USERMAIL=YMA_UTILISATEUR and YMS_MAILID=YMA_MAILID');
      ExecuteSQLContOnExcept('update YMESSAGES set YMS_STRDATE=(select YMA_STRDATE from YMAILS where YMS_USERMAIL=YMA_UTILISATEUR and YMS_MAILID=YMA_MAILID)');
    end;
    //GC
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_GEREARTICLELIE = "AUT" WHERE GPP_GEREARTICLELIE IS NULL');
    ExecuteSQLContOnExcept('update articlelie set gal_articlelie=isnull((select ga_article from article where ga_codearticle=gal_articlelie and ga_statutart in ("UNI","GEN")),gal_articlelie) where len(gal_articlelie)<34');

    ExecuteSQLContOnExcept('UPDATE PIECE SET GP_OPERATION = ""');

    // transfert elements de GCZONELIBRE vers GCZONELIBRECON
    ExecuteSQLContOnExcept('UPDATE CHOIXCOD SET CC_TYPE="ZLC" WHERE CC_TYPE="ZLI" AND (CC_CODE LIKE "BT%" OR CC_CODE LIKE "BB%" or CC_CODE LIKE "BC%" or CC_CODE LIKE "BD%" or CC_CODE LIKE "BM%")');


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
    ExecuteSQLContOnExcept('update JUSYSCLE set jsc_guidx = jsc_codex WHERE JSC_GUIDX = "" OR JSC_GUIDX IS NULL');
    //gp
    AglNettoieListes('GCSTKDISPODETAIL', 'GQD_DATEDISPO',nil);
    AglNettoieListes('GCSTKPICKING', 'GQD_DATEDISPO',nil);
    AglNettoieListes('GCSTKDISPATCH', 'GQD_DATEDISPO',nil);

    ExecuteSQLContOnExcept('update stknature set gsn_ctrldispo="000" where gsn_ctrldispo is null');
    ExecuteSQLContOnExcept('update stknature set gsn_ctrlperemption="000" where gsn_ctrlperemption is null');
  end;
  //PAIE
  ExecuteSQLContOnExcept('UPDATE DECLARATIONS SET PDT_AUTREVICT="-"');
  ExecuteSQLContOnExcept('UPDATE SALARIES SET PSA_ETATBULLETIN=(SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_PGBULDEFAUT")');
  ExecuteSQLContOnExcept('UPDATE SALARIES SET PSA_NUMEROBL=""');
  ExecuteSQLContOnExcept('UPDATE ORGANISMEPAIE SET  POG_CONDSPEC="-"');
  ExecuteSQLContOnExcept('UPDATE DUCSPARAM SET PDP_CONDITION="" WHERE PDP_PREDEFINI <> "CEG"');
  ExecuteSQLContOnExcept('UPDATE MAINTIEN SET PMT_HISTOMAINT ="-"  ');
  ExecuteSQLContOnExcept('UPDATE MOTIFABSENCE SET PMA_CARENCEIJSS=0');
  ExecuteSQLContOnExcept('UPDATE MOTIFABSENCE SET PMA_CARENCEIJSS=3 WHERE PMA_TYPEABS = "MAN"');
  ExecuteSQLContOnExcept('UPDATE MOTIFABSENCE SET PMA_CARENCEIJSS=1 WHERE PMA_TYPEABS = "MAP" OR PMA_TYPEABS = "ATJ" OR PMA_TYPEABS = "ATR"');
  ExecuteSQLContOnExcept('UPDATE MAINTIEN SET PMT_TYPEMAINTIEN="MAI" WHERE PMT_CARENCE <> 9999');
  ExecuteSQLContOnExcept('UPDATE MAINTIEN SET PMT_TYPEMAINTIEN="GAR" WHERE PMT_CARENCE = 9999');
  ExecuteSQLContOnExcept('UPDATE MAINTIEN SET PMT_HISTOMAINT ="-"');
  ExecuteSQLContOnExcept('UPDATE CRITMAINTIEN SET PCM_RUBGARANTIE=""');
  ExecuteSQLContOnExcept('UPDATE REMUNERATION SET PRM_EXCLMAINT="-", PRM_EXCLENVERS="-" WHERE PRM_PREDEFINI <> "CEG"');
  AglNettoieListes('PGMAINTIEN', 'PMT_HISTOMAINT', nil);
  // V Galliot 24042006 AglNettoieListes('PGMULMVTIJSS', 'PSA_LIBELLE', nil);
  AglNettoieListes('PGDUCSINIT', '', nil);
  st := 'update absencesalarie set pcn_dateannulation="' + UsDateTime(Idate1900) +
    '",pcn_etatpostpaie="VAL",pcn_motifannul="",pcn_annulepar=""';
  ExecuteSQLContOnExcept(st);
  st := 'update motifabsence set pma_annulable="-", pma_controlmotif="X",pma_createur="",pma_utilisateur="",' +
    'pma_pgcoloractif="10944422",pma_pgcolorpay="16053248",pma_pgcolorval="10944422",  pma_pgcoloratt="10210815"' +
    ', pma_pgcolorann="8421631",pma_pgcolorref="8421631",PMA_EDITPLANPAIE="X",PMA_EDITPLANABS="-",PMA_SSJOURFERIE="-"';
    ExecuteSQLContOnExcept(St);
  ExecuteSQLContOnExcept('update motifabsence set PMA_EDITPLANABS="X" WHERE PMA_MOTIFEAGL="X"');
  ExecuteSQLContOnExcept('update RECAPSALARIES set PRS_ACQUISANC=0, PRS_ACQUISCPSUIV=0, PRS_ACQUISRTTSUIV=0');
  ExecuteSQLContOnExcept('Update CONTRATTRAVAIL set PCI_FINTRAVAIL = PCI_FINCONTRAT');
  ExecuteSQLContOnExcept('UPDATE ENVOISOCIAL SET PES_CTRLENVOI="X" WHERE PES_CTRLENVOI<>"X" AND PES_CTRLENVOI<>"-"');
  ExecuteSQLContOnExcept('UPDATE DADSPERIODES SET PDE_EXERCICEDADS="2005", PDE_DECALEE="01"');
  ExecuteSQLContOnExcept('UPDATE DADSDETAIL SET PDS_EXERCICEDADS="2005"');
  ExecuteSQLContOnExcept('UPDATE DADS2SALARIES SET PD2_CHQEMPLOI="-", PD2_INDEMIMPATRI=0, PD2_AUTREREVENUS=0');
  ExecuteSQLContOnExcept('UPDATE ATTESTATIONS SET PAS_RECLASS="-", PAS_INDCNEB="-", PAS_INDCNEM=0, PAS_INDSINISTREB="-", PAS_INDSINISTREM=0, PAS_INDSPCFIQB="-", PAS_INDSPCFIQM=0, PAS_INDCPEB="-", PAS_INDCPEM=0, PAS_INDAPPRENTIB="-", PAS_INDAPPRENTIM=0');
  st := 'UPDATE MAINTIEN SET PMT_HISTOMAINT ="X" WHERE (PMT_DATEDEBUT = "'
        + usdatetime(strtodate('01/01/1959')) + '" AND PMT_DATEFIN = "'
        + usdatetime(strtodate('01/01/1959')) + '") OR  (PMT_DATEDEBUT = "' + usdatetime(strtodate('31/12/1900')) + '" AND PMT_DATEFIN = "' + usdatetime(strtodate('31/12/1900')) + '")';
  ExecuteSQLContOnExcept(St);
  st := 'UPDATE MAINTIEN SET PMT_DATEDEBUT = "' + usdatetime(strtodate('31/12/1900')) + '", PMT_DATEFIN = "' + usdatetime(strtodate('31/12/1900')) + '" WHERE PMT_DATEDEBUT = "' + usdatetime(strtodate('01/01/1959')) + '" AND PMT_DATEFIN = "' + usdatetime(strtodate('01/01/1959')) + '"';
  ExecuteSQLContOnExcept(St);
  ExecuteSQLContOnExcept('UPDATE MOTIFABSENCE SET PMA_RUBRIQUEJ="",PMA_ALIMENTJ="",PMA_PROFILABSJ="" WHERE PMA_PREDEFINI <> "CEG"');
  ExecuteSQLContOnExcept('UPDATE MOTIFABSENCE SET PMA_RUBRIQUEJ=PMA_RUBRIQUE,PMA_ALIMENTJ=PMA_ALIMENT,PMA_PROFILABSJ=PMA_PROFILABS WHERE PMA_JOURHEURE="JOU" AND PMA_PREDEFINI <> "CEG"');
  ExecuteSQLContOnExcept('UPDATE MOTIFABSENCE SET PMA_RUBRIQUE="",PMA_ALIMENT="",PMA_PROFILABS="" WHERE PMA_JOURHEURE="JOU" AND PMA_PREDEFINI <> "CEG" ');

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
    ExecuteSQLContOnExcept('UPDATE TIERSPIECE SET GTP_FAR_FAE=""');
    ExecuteSQLContOnExcept('UPDATE TXCPTTVA SET TV_CPTACHFARFAE="",  TV_CPTVTEFARFAE=""');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_FAR_FAE="",GPP_PIECESAV="-"');
    // Gestion mode dégradé de la caisse :
    ExecuteSQLContOnExcept('UPDATE PARCAISSE SET GPK_AUTDEGRADE="-", GPK_REPTICKETS="", GPK_ARTICLEDEGRADE=""');

    AglNettoieListes('RTRECHDOCGED', 'RTD_DOCGUID;YDF_FILEGUID',nil, 'RTD_DOCID;YDF_FILEID');
    AglNettoieListes('RTRECHDOCGEDGLO', 'RTD_DOCGUID;YDF_FILEGUID',nil, 'RTD_DOCID;YDF_FILEID');


  end;

  ExecuteSQLContOnExcept('UPDATE RUBRIQUE SET RB_DATEVALIDITE="' + USDATETIME(iDate2099) + '"');
  AGLNettoieListes('CPRUBRIQUE;RUBRIQUE','RB_NODOSSIER;RB_PREDEFINI;RB_DATEVALIDITE');

  //PAIE
  ExecuteSQLContOnExcept('UPDATE FORMATIONS SET PFO_CENTREFORMGU="" WHERE PFO_CENTREFORMGU IS NULL');
  ExecuteSQLContOnExcept('UPDATE FORMATIONS SET PFO_ORGCOLLECTGU="" WHERE PFO_ORGCOLLECTGU IS NULL');
  ExecuteSQLContOnExcept('UPDATE EMPLOIINTERIM SET PEI_CENTREFORMGU="" WHERE PEI_CENTREFORM IS NULL');
  ExecuteSQLContOnExcept('UPDATE EMPLOIINTERIM SET PEI_AGENCEINTGU="" WHERE PEI_AGENCEINTGU IS NULL');
  ExecuteSQLContOnExcept('UPDATE FRAISSALPLAF SET PFP_ORGCOLLECTGU="" WHERE PFP_ORGCOLLECTGU IS NULL');
  ExecuteSQLContOnExcept('UPDATE INVESTFORMATION SET PIF_ORGCOLLECTGU="" WHERE PIF_ORGCOLLECTGU IS NULL');
  ExecuteSQLContOnExcept('UPDATE SESSIONSTAGE SET PSS_CENTREFORMGU="" WHERE PSS_CENTREFORMGU IS NULL');
  ExecuteSQLContOnExcept('UPDATE SESSIONSTAGE SET PSS_ORGCOLLECTSGU="" WHERE PSS_ORGCOLLECTSGU IS NULL');
  ExecuteSQLContOnExcept('UPDATE SESSIONSTAGE SET PSS_ORGCOLLECTPGU="" WHERE PSS_ORGCOLLECTPGU IS NULL');
  ExecuteSQLContOnExcept('UPDATE STAGE SET PST_CENTREFORMGU="" WHERE PST_CENTREFORMGU IS NULL');
  ExecuteSQLContOnExcept('UPDATE STAGE SET PST_ORGCOLLECTPGU="" WHERE PST_ORGCOLLECTPGU IS NULL');
  ExecuteSQLContOnExcept('UPDATE STAGE SET PST_ORGCOLLECTSGU="" WHERE PST_ORGCOLLECTSGU IS NULL');
  ExecuteSQLContOnExcept('UPDATE ENVOIFORMATION SET PVF_ORGCOLLECTGU="" WHERE PVF_ORGCOLLECTGU IS NULL');
  ExecuteSQLContOnExcept('UPDATE VISITEMEDTRAV SET PVM_MEDTRAVGU="" WHERE PVM_MEDTRAVGU IS NULL');
  ExecuteSQLContOnExcept('UPDATE RETENUESALAIRE SET PRE_BENEFRSGU="" WHERE PRE_BENEFRSGU IS NULL');
  ExecuteSQLContOnExcept('UPDATE ETABCOMPL SET ETB_CODEDDTEFPGU="" WHERE ETB_CODEDDTEFPGU IS NULL');
  ExecuteSQLContOnExcept('UPDATE ETABCOMPL SET ETB_MEDTRAVGU="" WHERE ETB_MEDTRAVGU IS NULL');
  AglNettoieListes('PGANNUAIRE', 'ANN_GUIDPER',nil);


End;

Procedure MajVer739 ;
var st: string ;
    iMaxSSCC: Integer;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSQLContOnExcept('update JUEVENEMENT set JEV_OCCURENCEEVT="SIN", JEV_FOREIGNID="", JEV_FOREIGNAPP=""');
    // DP
    If IsMonoOuCommune then
    begin
      ExecuteSQLContOnExcept('update ANNUBIS set ANB_GUIDPERGRPINT=""');
      ExecuteSQLContOnExcept('update DPTABCOMPTA set DTC_TOTACTIFBILAN=0, DTC_TOTPASSIFBILAN=0');
      ExecuteSQLContOnExcept('update DOSSIER set DOS_ABSENT="X" where DOS_VERROU="ABS"');
      ExecuteSQLContOnExcept('update DOSSIER set DOS_ABSENT="-" where DOS_VERROU<>"ABS"');
    end;

    ExecuteSQLContOnExcept('update YGEDDICO set YGD_EWSREGLEPUB=""');
    // GC
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_GEREARTICLELIE="AUT" ');
    // J SICH  12/04/2006 ExecuteSQLContOnExcept('update articlelie set gal_article=isnull((select ga_article from article where ga_codearticle=gal_article and ga_statutart in ("UNI","GEN")),gal_article) where len(gal_article)<34 ');
    //GPAO

    ExecuteSQLContOnExcept('UPDATE PORT SET GPO_PDRNIV="" WHERE GPO_PDRNIV IS NULL');
    ExecuteSQLContOnExcept('UPDATE PORT SET GPO_PDRNIV="002" WHERE GPO_TYPEPDR<>""');
    ExecuteSQLContOnExcept('UPDATE PORT SET GPO_PDRMTBASE="" WHERE GPO_PDRMTBASE IS NULL');
    ExecuteSQLContOnExcept('UPDATE PORT SET GPO_PDRMTBASE="TOT" WHERE GPO_TYPEPDR<>""');
    ExecuteSQLContOnExcept('UPDATE WPDRTET SET WPE_VALEURPDRD=0'
     + ', WPE_VALPDRLIBRE01=0, WPE_VALPDRLIBRE02=0, WPE_VALPDRLIBRE03=0, WPE_VALPDRLIBRE04=0, WPE_VALPDRLIBRE05=0, WPE_VALPDRLIBRE06=0,'
     + ' WPE_VALPDRLIBRE07=0, WPE_VALPDRLIBRE08=0, WPE_VALPDRLIBRE09=0, WPE_VALPDRLIBRE10=0, WPE_VALPDRLIBRE11=0, WPE_VALPDRLIBRE12=0'
     + ' WHERE WPE_VALEURPDRD IS NULL');
    ExecuteSQLContOnExcept('UPDATE WPDRTYPE SET WRT_MAJDPADPR="-" WHERE WRT_MAJDPADPR IS NULL');
    ExecuteSQLContOnExcept('UPDATE WPARAM SET WPA_BOOLEAN10="-"');

    ExecuteSQLContOnExcept('UPDATE PARACTIONS set RPA_QNATUREACTION="DIV", RPA_NATURETRAVAIL="" where RPA_QNATUREACTION is null');

   { ExecuteSQLContOnExcept('UPDATE ACTIONS SET RAC_IDENTIFIANT=0, RAC_QORIGINERQ="",RAC_QUALITETYPE="",RAC_TYPELOCA="",'
                            +'RAC_QDEMDEROGNUM=0,RAC_QPLANCORRNUM=0,RAC_QNCNUM=0,RAC_NATURETRAVAIL="",'
                            +'RAC_LIGNEORDRE=0,RAC_AREALISERPAR="",RAC_QPCTAVANCT=0,RAC_REALISEELE="'+string(UsDateTime(iDate1900))+'",'
                            +'RAC_REALISEEPAR="",RAC_VERIFIEELE="'+ string(UsDateTime(iDate1900))+'",'
                            +'RAC_VERIFIEEPAR="",RAC_EFFICACITE="",RAC_EFFJUGEELE="'+string(UsDateTime(iDate1900))+'",'
                            +'RAC_EFFJUGEEPAR="",RAC_CLOTUREELE="'+string(UsDateTime(iDate1900))+'",RAC_CLOTUREEPAR="" '
                            +'WHERE RAC_IDENTIFIANT is null');    }

    { Mise à jour du contexte de jetons pour le colisage }
    ExecuteSQLContOnExcept('UPDATE WJETONS'
             + ' SET WJT_CONTEXTE="UL"||(SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_SSCCCHAREXT")'
             + ' WHERE WJT_PREFIXE="GCS" AND WJT_CONTEXTE="ULI"');

    { Mise à jour du jeton en reformant le code unité logistique le plus grand }
    if (wGetSQLFieldValue('SOC_DATA', 'PARAMSOC', 'SOC_NOM="SO_CNUF"') <> '')
    and ExisteSQL('SELECT 1 FROM COLISAGE WHERE GCS_ESTUL="X" AND GCS_ORDRE>0') then
    begin
      iMaxSSCC := wGetSQLFieldValue('CAST(SUBSTRING(GCS_SSCC, 3 +LEN(TRIM((SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_CNUF"))),'
                                   +                         '15-LEN(TRIM((SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_CNUF")))'
                                   +              ') AS INTEGER'
                                   +    ') + 1',
                                   'COLISAGE',
                                   'GCS_ESTUL="X" AND GCS_ORDRE>0',
                                   'CAST(SUBSTRING(GCS_SSCC, 3 +LEN(TRIM((SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_CNUF"))),'
                                   +                       ' 15-LEN(TRIM((SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_CNUF")))'
                                   +             ') AS INTEGER'
                                   +   ') DESC'
                                   )
                                   ;
      ExecuteSQLContOnExcept('UPDATE WJETONS SET WJT_JETON=' + IntToStr(iMaxSSCC)
                + ' WHERE WJT_PREFIXE="GCS" AND WJT_CONTEXTE="UL"||(SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_SSCCCHAREXT")');
    end;


    { Mise à jour table EDILIGNE }
    ExecuteSQLContOnExcept('UPDATE EDILIGNE SET ELI_LIGNEORILU="", ELI_LIGNEORI="", ELI_QUALIFQTELU=""'
             + ' WHERE ELI_LIGNEORILU IS NULL'
             + ' AND ELI_LIGNEORI IS NULL'
             + ' AND ELI_QUALIFQTELU IS NULL');

    { Mise à jour table EDIMESSAGES }
    ExecuteSQLContOnExcept('UPDATE EDIMESSAGES SET'
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
    ExecuteSQLContOnExcept('DELETE FROM WPARAM WHERE WPA_CODEPARAM="EMG_AUTOMATE"');

    { Mise à jour table LIGNECOMPL }
    {1. 1er passage}
    ExecuteSQLContOnExcept('UPDATE LIGNECOMPL'
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
    ExecuteSQLContOnExcept('UPDATE LIGNECOMPL'
             + ' SET GLC_AFFCONTREM="", GLC_AUTOGENCONTREM="-"'
             + ' WHERE GLC_AUTOGENCONTREM IS NULL'
             + ' AND GLC_AFFCONTREM IS NULL');

    { Modif. de listes }
    AglNettoieListesPlus('EDIMESSAGES', 'EMG_IDENTIFIANT;EMG_TYPETRANS;EMG_CODEMESSAGE', nil, True, '');
    // GC
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_GEREARTICLELIE="AUT"');
    // J Sich 18/04/2006 ExecuteSQLContOnExcept('update articlelie set gal_article=isnull((select ga_article from article where ga_codearticle=gal_article and ga_statutart in ("UNI","GEN")),gal_article) where len(gal_article)<34');
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
    ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DESIGN="36;Service après vente;0;W" WHERE SOC_NOM="SCO_SAV"');
    //GPAO
    ExecuteSQLContOnExcept('UPDATE TIERSCOMPL SET YTC_PALMARESTRA = "-", YTC_CODEPORT="" WHERE YTC_PALMARESTRA IS NULL');
    ExecuteSQLContOnExcept('UPDATE WPARAMFONCTION SET WPF_SECTEURGEO="", WPF_PAYS="", WPF_CODEPOSTAL="",WPF_REGION="",WPF_VENTEACHAT="" WHERE WPF_SECTEURGEO IS NULL');
    ExecuteSQLContOnExcept('UPDATE PIEDPORT SET GPT_TIERSFRAIS="", GPT_ORIGINE="" WHERE GPT_TIERSFRAIS IS NULL');
    ExecuteSQLContOnExcept('UPDATE LIGNEFRAIS SET LF_TIERSFRAIS="" WHERE LF_TIERSFRAIS IS NULL');
    AglNettoieListesPlus('GCSTKDISPO','GA_COEFCONVQTEACH;GA_COEFCONVQTEVTE;GA_COEFPROD;GA_COEFCONVCONSO' ,nil);
    //MES
    ExecuteSQLContOnExcept('UPDATE WORDREGAMME SET WOG_AUTOMATISME="" WHERE WOG_AUTOMATISME IS NULL');
    MajMarkMESApres;
	  //GIGA
    ExecuteSQLContOnExcept('UPDATE PROFILGENER SET APG_TRIPROFIL=""');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_REPRISEENTAFF="",GPP_REPRISELIGAFF=""');
    AglNettoieListes('AFMULEAFFAIRE', 'EAF_TIERS',nil);
    //
		if (GetParamSocSecur('SO_AFFRAISCOMPTA', false) = true) then
		begin
      ExecuteSql ('UPDATE ACTIVITE SET ACT_MODESAISFRAIS="NOR" WHERE ACT_TYPEARTICLE<>"FRA" OR ACT_MONTANTTTC=0');
      ExecuteSql ('UPDATE EACTIVITE SET EAC_MODESAISFRAIS="NOR" WHERE EAC_TYPEARTICLE<>"FRA" OR EAC_MONTANTTTC=0');
      ExecuteSql ('UPDATE ACTIVITE SET ACT_MODESAISFRAIS="CPT" WHERE ACT_TYPEARTICLE="FRA" AND ACT_MONTANTTTC<>0');
      ExecuteSql ('UPDATE EACTIVITE SET EAC_MODESAISFRAIS="CPT" WHERE EAC_TYPEARTICLE="FRA" AND EAC_MONTANTTTC<>0');
		end
		else
		begin
      ExecuteSql ('UPDATE ACTIVITE SET ACT_MODESAISFRAIS="NOR"');
      ExecuteSql ('UPDATE EACTIVITE SET EAC_MODESAISFRAIS="NOR"');
		end;
	  //fin GIGA
    If IsMonoOuCommune then
    begin
      if not (isOracle or isDB2) then
      begin
        ExecuteSQLContOnExcept( 'UPDATE ANNUAIRE SET'
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
      ExecuteSQLContOnExcept( 'UPDATE ANNUAIRE SET ANN_CODEINSEE     = ""  WHERE ANN_CODEINSEE     IS NULL');
      ExecuteSQLContOnExcept( 'UPDATE ANNUAIRE SET ANN_COOP          = "-" WHERE ANN_COOP          IS NULL');
      ExecuteSQLContOnExcept( 'UPDATE ANNUAIRE SET ANN_FORMEGRPPRIVE = ""  WHERE ANN_FORMEGRPPRIVE IS NULL');
      ExecuteSQLContOnExcept( 'UPDATE ANNUAIRE SET ANN_FORMESTE      = ""  WHERE ANN_FORMESTE      IS NULL');
      ExecuteSQLContOnExcept( 'UPDATE ANNUAIRE SET ANN_FORMESCI      = ""  WHERE ANN_FORMESCI      IS NULL');

      if not (isOracle or isDB2) then
      begin
        ExecuteSQLContOnExcept( 'UPDATE DPORGA SET'
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
      ExecuteSQLContOnExcept('UPDATE HRCAISSE SET HRC_GESTIONPDA="-" WHERE HRC_GESTIONPDA IS NULL');
      ExecuteSQLContOnExcept('UPDATE HRPARAMPLANNING SET HPP_LPARAMPLANNING="" WHERE HPP_LPARAMPLANNING IS NULL');

      ExecuteSQLContOnExcept('UPDATE HRDOSSIER SET HDC_CBINTERNET="",HDC_CBLIBELLE="",HDC_DATEEXPIRE="",HDC_TYPECARTE="",HDC_CBNUMCTRL="",HDC_CBNUMAUTOR="",HDC_NUMCHEQUE=""  WHERE HDC_CBINTERNET IS NULL');

      //GC
      ExecuteSQLContOnExcept('update etables set edt_nbenreg=0, edt_taillemoyenne=0, edt_tempstraite=0, edt_active="X"');
      ExecuteSQLContOnExcept('UPDATE PARCAISSE SET GPK_IP="", GPK_IMPMODTICDGD=""');
      //GPAO
      ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA="CMC" WHERE SOC_NOM="SO_CDEMARCHEVTE"');
      ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA="CC;" WHERE SOC_NOM="SO_CDEMARCHEAPPELVTE"');
      ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA="GL_DATELIVRAISON" WHERE SOC_NOM="SO_CDEMARCHEAPPELDATEVTE"');
      //SERVANTISSIMO
      ExecuteSQLContOnExcept('UPDATE IMOREF SET IRF_CONFIDENTIEL ="0"');
    end;

  end;

// TRESO
AglNettoieListes('TRECRITURERAPPRO', 'TE_NODOSSIER;',nil);
AglNettoieListes('TRMODIFIERUBRIQUE', 'TE_NODOSSIER;',nil);
AglNettoieListes('TRECRITURE', 'TE_NODOSSIER;',nil);
AglNettoieListes('TRPREVISIONNELLES', 'TE_NODOSSIER;',nil);
AglNettoieListes('TRLISTESYNCHRO', 'TE_NODOSSIER;TE_NUMTRANSAC;TE_NUMLIGNE;TE_DATEVALEUR;TE_GENERAL;TE_MONTANTDEV;TE_DEVISE;TE_JOURNAL;TE_QUALIFORIGINE;TE_LIBELLE;TE_NUMEROPIECE;TE_CODECIB;TE_CODEFLUX;TE_DATECOMPTABLE;TE_NATURE;TE_EXERCICE;TE_DATECREATION; ',nil);
AglNettoieListes('TRSUPPRECRITURE', 'TE_NODOSSIER;',nil);
ExecuteSQLContOnExcept('UPDATE COURTSTERMES SET TCT_NODOSSIER = "000000"');
ExecuteSQLContOnExcept('UPDATE EQUILIBRAGE SET TEQ_SNODOSSIER = "000000"');
ExecuteSQLContOnExcept('UPDATE EQUILIBRAGE SET TEQ_DNODOSSIER = "000000"');
ExecuteSQLContOnExcept('UPDATE TRECRITURE SET TE_NODOSSIER = "000000"');
ExecuteSQLContOnExcept('UPDATE TROPCVM SET TOP_NODOSSIER = "000000"');
ExecuteSQLContOnExcept('UPDATE TRVENTEOPCVM SET TVE_NODOSSIER = "000000"');
ExecuteSQLContOnExcept('UPDATE BANQUECP SET BQ_NODOSSIER = "000000"');
//
ExecuteSQLContOnExcept('UPDATE SUIVCPTA SET SC_CONTROLEQTE = "-"' ) ;
ExecuteSQLContOnExcept('UPDATE SECTION SET S_FINCHANTIER = "' + UsDateTime(iDate2099) + '" WHERE S_FINCHANTIER = "' + UsDateTime(iDate1900) + '"');
//Compta
ExecuteSQLContOnExcept('UPDATE CBALSIT SET BSI_DATEMODIF="' + USDATETIME(iDate1900) + '",BSI_TYPEECR="N"');
ExecuteSQLContOnExcept('UPDATE CBALSITECR SET BSE_QUALIFPIECE="",BSE_ETABLISSEMENT="",BSE_DEVISE=""');
//PAie

  St := 'UPDATE MASQUESAISRUB SET PMR_AIDECOL1="",PMR_AIDECOL2="",PMR_AIDECOL3="",PMR_AIDECOL4="",'+
        'PMR_AIDECOL5="",PMR_AIDECOL6="",PMR_AIDECOL7="",PMR_UTILISATEUR="",PMR_DATECREATION="'+
        UsDateTime(Idate1900)+'",PMR_DATEMODIF="'+UsDateTime(Idate1900)+'"';
  ExecuteSQLContOnExcept(St);

  ExecuteSQLContOnExcept('UPDATE CRITMAINTIEN SET PCM_VALCATEGORIE = PCM_VALCATEG');
  ExecuteSQLContOnExcept('UPDATE FORMATIONS SET PFO_SIGNIFICATIF="-"');
  ExecuteSQLContOnExcept('UPDATE RHCOMPETRESSOURCE SET PCH_PGPROVCOMP="",PCH_RHPROFIL=""');
  ExecuteSQLContOnExcept('UPDATE RHCOMPETENCES SET PCO_RHNATURECOMP="",PCO_DUREEACQUIS=0,'+
  'PCO_RHVALIDEURCOMP="",PCO_RHMODEVALID="",PCO_RHNIVEAUCOMP="",PCO_RHMACROCOMP=""');
  ExecuteSQLContOnExcept('UPDATE REFERENTIELMETIER SET PRF_CODEEMPLOI=""');

  St:= 'UPDATE EMETTEURSOCIAL SET PET_CRECIVILITE="", PET_CRENOM="", PET_CIVIL1DADSU="", PET_TEL1DADSU="",'+
       ' PET_FAX1DADSU="", PET_CIVIL2DADSU="", PET_TEL2DADSU="", PET_FAX2DADSU="", PET_CIVIL3DADSU="",'+
       ' PET_TEL3DADSU="", PET_FAX3DADSU=""';
  ExecuteSQLContOnExcept(St);
  ExecuteSQLContOnExcept('UPDATE EMETTEURSOCIAL SET PET_TEL1DADSU=PET_APPEL1DUDS WHERE PET_MEDIADUDS1="01"');
  ExecuteSQLContOnExcept('UPDATE EMETTEURSOCIAL SET PET_APPEL1DUDS="", PET_MEDIADUDS1="" WHERE PET_MEDIADUDS1="01"');
  ExecuteSQLContOnExcept('UPDATE EMETTEURSOCIAL SET PET_FAX1DADSU=PET_APPEL1DUDS WHERE PET_MEDIADUDS1="02"');
  ExecuteSQLContOnExcept('UPDATE EMETTEURSOCIAL SET PET_APPEL1DUDS="", PET_MEDIADUDS1="" WHERE PET_MEDIADUDS1="02"');
  ExecuteSQLContOnExcept('UPDATE EMETTEURSOCIAL SET PET_MEDIADUDS1="" WHERE PET_MEDIADUDS1="03"');
  ExecuteSQLContOnExcept('UPDATE EMETTEURSOCIAL SET PET_TEL2DADSU=PET_APPEL2DUDS WHERE PET_MEDIADUDS2="01"');
  ExecuteSQLContOnExcept('UPDATE EMETTEURSOCIAL SET PET_APPEL2DUDS="", PET_MEDIADUDS2="" WHERE PET_MEDIADUDS2="01"');
  ExecuteSQLContOnExcept('UPDATE EMETTEURSOCIAL SET PET_FAX2DADSU=PET_APPEL2DUDS WHERE PET_MEDIADUDS2="02"');
  ExecuteSQLContOnExcept('UPDATE EMETTEURSOCIAL SET PET_APPEL2DUDS="", PET_MEDIADUDS2="" WHERE PET_MEDIADUDS2="02"');
  ExecuteSQLContOnExcept('UPDATE EMETTEURSOCIAL SET PET_MEDIADUDS2="" WHERE PET_MEDIADUDS2="03"');
  ExecuteSQLContOnExcept('UPDATE EMETTEURSOCIAL SET PET_TEL3DADSU=PET_APPEL3DUDS WHERE PET_MEDIADUDS3="01"');
  ExecuteSQLContOnExcept('UPDATE EMETTEURSOCIAL SET PET_APPEL3DUDS="", PET_MEDIADUDS3="" WHERE PET_MEDIADUDS3="01"');
  ExecuteSQLContOnExcept('UPDATE EMETTEURSOCIAL SET PET_FAX3DADSU=PET_APPEL3DUDS WHERE PET_MEDIADUDS3="02"');
  ExecuteSQLContOnExcept('UPDATE EMETTEURSOCIAL SET PET_APPEL3DUDS="", PET_MEDIADUDS3="" WHERE PET_MEDIADUDS3="02"');
  ExecuteSQLContOnExcept('UPDATE EMETTEURSOCIAL SET PET_MEDIADUDS3="" WHERE PET_MEDIADUDS3="03"');

  ExecuteSQLContOnExcept('UPDATE MOTIFABSENCE SET PMA_ALIMNETJ="-",PMA_ALIMNETH="-"');
  ExecuteSQLContOnExcept('UPDATE MOTIFABSENCE SET PMA_ALIMNETJ="X" WHERE PMA_RUBRIQUEJ <> ""');
  ExecuteSQLContOnExcept('UPDATE MOTIFABSENCE SET PMA_ALIMNETH="X" WHERE PMA_RUBRIQUE <> ""');
  ExecuteSQLContOnExcept('UPDATE REMUNERATION SET PRM_METHODARRONDI="",PRM_REMLIEREGUL="" WHERE PRM_PREDEFINI<> "CEG"');

  ExecuteSQLContOnExcept('UPDATE ETABCOMPL SET ETB_PAIEVALOMS="",ETB_MSAUNITEGES="",ETB_RATTACHEHANDI=""');
  ExecuteSQLContOnExcept('UPDATE ETABCOMPL SET ETB_PAIEVALOMS="ANT" WHERE ETB_CONGESPAYES="X"');
  ExecuteSQLContOnExcept('UPDATE SALARIES SET PSA_TYPPAIEVALOMS="ETB", PSA_PAIEVALOMS=""');
  St:= 'UPDATE SALARIES SET PSA_TYPPAIEVALOMS="ETB", PSA_PAIEVALOMS="ANT" WHERE PSA_ETABLISSEMENT IN'+
	' (SELECT ETB_ETABLISSEMENT FROM ETABCOMPL WHERE ETB_CONGESPAYES="X")';
  ExecuteSQLContOnExcept(St);
  St:= 'update motifabsence set PMA_TYPEPERMAXI="",PMA_PERMAXI=0,PMA_OUVRES="-",PMA_OUVRABLE="-",PMA_CALENDSAL="-"';
  ExecuteSQLContOnExcept(St);
  ExecuteSQLContOnExcept('UPDATE MOTIFABSENCE SET PMA_CALENDSAL="X" WHERE PMA_CALENDCIVIL="-"');
  St:= 'UPDATE SALARIES SET PSA_DATEVIEACTIVE="'+UsDateTime(IDate1900)+'"';
  ExecuteSQLContOnExcept(St);
  St:= 'UPDATE DEPORTSAL SET PSE_MSATYPUNITEG="ETB", PSE_MSAUNITEGES="", PSE_RESPONSAUG=PSE_RESPONSVAR,'+
	' PSE_ASSISTAUG=PSE_ASSISTVAR';
  ExecuteSQLContOnExcept(St);
  St:= 'UPDATE SERVICES SET PGS_RESPONSAUG=PGS_RESPONSVAR, PGS_SECRETAIREAUG=PGS_SECRETAIREVAR,'+
	' PGS_ADJOINTAUG=PGS_ADJOINTVAR';
  ExecuteSQLContOnExcept(St);
  ExecuteSQLContOnExcept('UPDATE FRAISSALPLAF SET PFP_CATEGFRAISFOR=""');
  ExecuteSQLContOnExcept('UPDATE SESSIONSTAGE SET PSS_NUMSESSION=-1, PSS_TYPESESSSTAGE="IND"');
  St:= 'UPDATE PGHISTODETAIL SET PHD_DATEDEBVALID="'+UsDateTime(IDate1900)+'", PHD_RUBRIQUE=""';
  ExecuteSQLContOnExcept(St);
  St:= 'UPDATE SALARIES SET PSA_TYPCONVENTION="ETB" WHERE PSA_CONVENTION IN (SELECT ETB_CONVENTION FROM ETABCOMPL'+
	' WHERE ETB_ETABLISSEMENT=PSA_ETABLISSEMENT)';
  ExecuteSQLContOnExcept(St);
  St:= 'UPDATE SALARIES SET PSA_TYPCONVENTION="PER" WHERE PSA_CONVENTION NOT IN (SELECT ETB_CONVENTION FROM ETABCOMPL'+
	' WHERE ETB_ETABLISSEMENT=PSA_ETABLISSEMENT)';
  ExecuteSQLContOnExcept(St);

  //IMMO
  ExecuteSQLContOnExcept('UPDATE IMMO SET I_SUSDEF="A",I_REGLECESSION="NOR",I_NONDED="-",I_REPRISEINT=0.00,I_REPRISEDEPCEDEE=0.00,I_REPRISEINTCEDEE=0.00,I_DPI="-",I_DPIEC="-",I_CORRECTIONVR=0.00,I_CORVRCEDDE=0.00,I_SUBVENTION="NON",I_SBVPRI=0.00');
  ExecuteSQLContOnExcept('UPDATE IMMO SET I_SBVMTC=0.00,I_SBVPRIC=0.00,I_SBVEC="C",I_SBVDATE="'+UsDateTime(iDate1900)+'",I_CPTSBVR="",I_CPTSBVB="",I_PFR="-",I_COEFDEG=0.00,I_AMTFOR=0.00,I_AMTFORC=0.00,I_ACHFOR="'+UsDateTime(iDate1900)+'"');
  ExecuteSQLContOnExcept('UPDATE IMMO SET I_PRIXACFORC=0.00,I_VNCFOR=0.00,I_DURRESTFOR=0,I_DATEDEBECO=I_DATEAMORT,I_DATEDEBFIS=I_DATEAMORT');

  ExecuteSQLContOnExcept('UPDATE IMMO SET I_DATEDEBECO=I_DATEPIECEA WHERE I_METHODEECO="DEG"');
  ExecuteSQLContOnExcept('UPDATE IMMO SET I_DATEDEBFIS=I_DATEPIECEA WHERE I_METHODEFISC="DEG"');

  ExecuteSQLContOnExcept('UPDATE IMMOLOG SET IL_TAUX=0.0');

  ExecuteSQLContOnExcept('UPDATE IMMOAMOR SET IA_NONDEDUCT=0.00,IA_CESSIONND=0.00,IA_MONTANTSBV=0.00,IA_CESSIONSBV=0.00,IA_MONTANTDPI=0.00,IA_CESSIONDPI=0.00,IA_MONTANTARD=0.00,IA_CESSIONARD=0.00');



End;

Procedure MajVer741 ;
var i: integer;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSQLContOnExcept('UPDATE SECTION SET S_FINCHANTIER = "'+USDateTime(idate2099)+'" WHERE S_FINCHANTIER = "'+USDateTime(idate1900)+'"');
    ExecuteSQLContOnExcept('UPDATE SECTION SET S_AFFAIREENCOURS="X"');
    //DP
    ExecuteSQL ('UPDATE ACTIONS SET RAC_PRIVATE="-"') ;
    ExecuteSQLContOnExcept('UPDATE YDOCUMENTS SET YDO_CREATEUR=YDO_UTILISATEUR, YDO_LIBREGED1="", YDO_LIBREGED2="", YDO_LIBREGED3="", YDO_LIBREGED4="", YDO_LIBREGED5=""');
    // Insertion des libellés des tablettes libres GED
    for i:=1 to 5 do InsertChoixCode ('ZLG', 'ZL'+IntToStr(i), '', '','');
    InsertChoixCode('AAO', 'POR', 'Portail Web', '', '');

    ExecuteSQLContOnExcept('UPDATE ANNULIEN SET'
             +' ANL_COOPTDATE = "' + UsDateTime(iDate1900) + '", ANL_DTAGJPRES = "' + UsDateTime(iDate1900) + '",'
             +' ANL_DTCAJPRES = "' + UsDateTime(iDate1900) + '", ANL_EXPEANTICIP = "' + UsDateTime(iDate1900) + '",'
             +' ANL_FRALIBSOUSCRI = 0, ANL_JETONSPRES = "-",'
             +' ANL_LIMPOUV = 0, ANL_MODREMUN = "",'
             +' ANL_MTRESTANTLIB = 0, ANL_MTVERSEMENT = 0,'
             +' ANL_NANTISSEMENT   = "-"  , ANL_PEA   = "-"  ');

    ExecuteSQLContOnExcept('UPDATE ANNUAIRE SET'
             +' ANN_CAPLIBDTLIMITE = "' + UsDateTime(iDate1900) + '", ANN_CAPLIBFRACTION = 0,'
             +' ANN_NBTITDIVPRIO = 0, ANN_PREPONDEIMMO = "-" ');

    ExecuteSQLContOnExcept('UPDATE DPFISCAL SET'
             +' DFI_DEFICITORDIN = "-", DFI_FRANCPRECPT = "-", DFI_RESIDENT = "-" ');

    ExecuteSQLContOnExcept('UPDATE DPSOCIAL SET'
             +' DSO_EPARGNESAL = "-", DSO_PREVCOLLECT = "-",'
             +' DSO_PREVFAC = "-", DSO_RETCOLLECTIF = "-", DSO_RETRAITEFAC = "-" ');

    ExecuteSQLContOnExcept('UPDATE JURIDIQUE SET'
             +' JUR_AGMAJORITE = 0, JUR_AGQUORUM = 0,'
             +' JUR_ARD = "-", JUR_CAMAJORITE = 0,'
             +' JUR_CAQUORUM = 0, JUR_DATEPARUTION = "' + UsDateTime(iDate1900) + '",'
             +' JUR_DTLIMINSCRIACT = "' + UsDateTime(iDate1900) + '", JUR_DTMODIFSTAT = "' + UsDateTime(iDate1900) + '",'
             +' JUR_GESTDRTSOCIAUX = "-", JUR_NUMSTATDURSOC = "" ');
    //GC
    ExecuteSql ('UPDATE QBPARBRE SET ' +
             ' QBR_HISTOCA=0,QBR_REALISECA=0,QBR_PREVUCA=0,QBR_CARETENU=0,QBR_MAILLE="",QBR_EVOLPREVQTE=0,QBR_EVOLPREVQTEPRC=0,QBR_SAISIPREVQTE=0,' +
             ' QBR_EVOLPREVCA=0,QBR_EVOLPREVCAPRCT=0,QBR_SAISIPREVCA=0,QBR_NBCLTPROSPECT=0,QBR_NBCLTC=0,QBR_PROSPECTQTE=0,QBR_VUREFQTE=0,' +
             ' QBR_PERDUQTE=0,QBR_VUCOURANTQTE=0,QBR_NOUVEAUQTE=0,QBR_EXTRAPOLABLE=""')  ;

    GCTraite_TIERSIMPACTPIECE;
    
    //GP
    AglNettoieListes('GCSTKDISPODETAIL','GA_QUALIFUNITESTO;GA_QUALIFUNITEVTE,GA_UNITEQTEVTE,GA_UNITEQTEACH,GA_UNITEPROD,GA_UNITECONSO',nil);
    AGLNettoieListes('GCSTKPREVISION','GA_COEFCONVQTEVTE;GA_COEFCONVQTEACH;GA_PRIXPOURQTEAC;GA_PRIXPOURQTE;GA_COEFPROD;GA_COEFCONVCONSO',nil);

  end;
// PAIE
AglNettoieListes('PGMULINSCFOR', 'PFI_TYPEPLANPREV',nil);
//COMPTA
ExecuteSQLContOnExcept('UPDATE IMMO SET I_OPEREG="-"');

End;

Procedure MajVer742 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
     //giga
     {AF_20060510_PGIMAJVER.TXT
     ExecuteSQL ('UPDATE AFFAIRE SET AFF_NATUREPIECEG="'+ GetParamSoc ('SO_AFNATAFFAIRE') + '" '
            + ' WHERE AFF_AFFAIRE0="A" AND (AFF_NATUREPIECEG="FPR" OR AFF_NATUREPIECEG="") ');
     ExecuteSQL ('UPDATE AFFAIRE SET AFF_NATUREPIECEG="'+ GetParamSoc ('SO_AFNATPROPOSITION') + '" '
            + ' WHERE AFF_AFFAIRE0="P" AND (AFF_NATUREPIECEG="FPR" OR AFF_NATUREPIECEG="") ');}
     ExecuteSQL ('UPDATE AFFAIRE SET AFF_NATUREPIECEG="'+ GetParamSoc ('SO_AFNATAFFAIRE') + '" '
            + ' WHERE AFF_AFFAIRE0="A" AND (AFF_NATUREPIECEG="FPR" OR AFF_NATUREPIECEG="" OR AFF_NATUREPIECEG="FAC") ');
     ExecuteSQL ('UPDATE AFFAIRE SET AFF_NATUREPIECEG="'+ GetParamSoc ('SO_AFNATPROPOSITION') + '" '
            + ' WHERE AFF_AFFAIRE0="P" AND (AFF_NATUREPIECEG="FPR" OR AFF_NATUREPIECEG="" OR AFF_NATUREPIECEG="FAC") ');
	   //gpao
     AGLNettoieListes('WORDREBESL','WOB_COEFLIEN',nil);
     AglNettoieListes('GCLISTINVLIGNE','GIL_LOTEXTERNE;GIL_SERIEEXTERNE',nil);
     ExecuteSQLContOnExcept('UPDATE PORT SET GPO_PDRAVECFRAIS="-" WHERE GPO_PDRAVECFRAIS IS NULL');
     ExecuteSQLContOnExcept('UPDATE PORT SET GPO_PDRMTBASE="MAC;MAT;SAL;STR;" WHERE GPO_PDRAVECFRAIS = "TOT"');
     AglNettoieListes('GCSTKDISPODETAIL','GA_QUALIFUNITESTO;GA_QUALIFUNITEVTE;GA_UNITEQTEVTE;GA_UNITEQTEACH;GA_UNITEPROD;GA_UNITECONSO;GA_COEFCONVQTEACH;GA_COEFCONVQTEVTE;GA_COEFPROD;GA_COEFCONVCONSO',nil);
     AglNettoieListes('GCSTKPHYSIQUE','GA_COEFCONVQTEACH;GA_COEFCONVQTEVTE;GA_COEFPROD;GA_COEFCONVCONSO',nil);
     AglNettoieListes('GCSTKPREVISION','GA_COEFCONVQTEACH;GA_COEFCONVQTEVTE;GA_COEFPROD;GA_COEFCONVCONSO',nil);
     AglNettoieListes('GCSTKRUPSTA','GA_COEFCONVQTEACH;GA_COEFCONVQTEVTE;GA_COEFPROD;GA_COEFCONVCONSO',nil);
     AglNettoieListes('GCSTKTRF','GA_COEFCONVQTEACH;GA_COEFCONVQTEVTE;GA_COEFPROD;GA_COEFCONVCONSO',nil);
     //PGISIDE
     AglNettoieListes('PSE_MULEXPORTS', 'EDT_FICHIEREXT',nil);
     AglNettoieListes('PSE_MULIMPORTS', 'EDT_FICHIEREXT',nil);


  end;
End;

Procedure MajVer743 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //GC
    // D brosset 24042006  supprimé ExecuteSQLContOnExcept('update tiers set t_sexe="H" where t_sexe="M"');
//    LanceMoulinetteArticleLie;   ///** Ajouter le source   Y:\PgiS5\DIFFUSION\Gescom\Lib\BasculeArticleLie.pas  **///
// js1 080606 gc    MajBundlePartage('YS7;YS8');
    //GPAO
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_TARIFMODULE="301" WHERE GPP_NATUREPIECEG IN ("CSA","BSA")');

  end;
End;


Procedure MajVer744 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //DP
    If IsMonoOuCommune then
    begin
      ExecuteSQLContOnExcept('UPDATE JUTYPECIVIL SET JTC_SEXE="M" WHERE JTC_SEXE="H"');
      ExecuteSQLContOnExcept('UPDATE ANNUAIRE SET ANN_SEXE="M" WHERE ANN_SEXE="H"');
      ExecuteSQLContOnExcept('UPDATE ANNUAIRE SET ANN_SEXE="" WHERE ANN_SEXE="I"');
    end;
    //GPAO
    { EDI - Commande achat }
    if not ExisteSQL('SELECT 1 FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="CDA"') then
      ExecuteSQLContOnExcept('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE     , EFM_NATURESPIECE, EFM_SERIALISE)'
                +                   ' VALUES("CDA"          , "Commande achat", "CF,CFR"        , "-")')
    else
      ExecuteSQLContOnExcept('UPDATE EDIFAMILLEEMG SET EFM_NATURESPIECE="CF,CFR" WHERE EFM_CODEFAMILLE="CDA"');

    { EDI - Réception achat }
    if not ExisteSQL('SELECT 1 FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="BLF"') then
      ExecuteSQLContOnExcept('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE       , EFM_NATURESPIECE, EFM_SERIALISE)'
                +                   ' VALUES("BLF"          , "Bon de réception", "PRF,BLF,BRC"   , "-")')
    else
      ExecuteSQLContOnExcept('UPDATE EDIFAMILLEEMG SET EFM_NATURESPIECE="PRF,BLF,BRC" WHERE EFM_CODEFAMILLE="BLF"');

    { EDI - Facture achat }
    if not ExisteSQL('SELECT 1 FROM EDIFAMILLEEMG WHERE EFM_CODEFAMILLE="FAA"') then
      ExecuteSQLContOnExcept('INSERT INTO EDIFAMILLEEMG (EFM_CODEFAMILLE, EFM_LIBELLE    , EFM_NATURESPIECE   , EFM_SERIALISE)'
                +                   ' VALUES("FAA"          , "Facture achat", "FF,FRA,AF,AFP,AVS", "-")');

    ExecuteSQLContOnExcept('UPDATE WPDRTET SET WPE_PERIODESAUV ='
        +' IIF((WPE_UNPDRPAR="010"),SUBSTRING(WPE_PERIODESAUV,1,4)||"-1-1-01-01-000001",'
        +' IIF(((WPE_UNPDRPAR="020") AND (SUBSTRING(WPE_PERIODESAUV,6,1)="1"))'
        +' ,SUBSTRING(WPE_PERIODESAUV,1,6)||"-1-01-01-000001",'
        +' IIF(((WPE_UNPDRPAR="020") AND (SUBSTRING(WPE_PERIODESAUV,6,1)="2"))'
        +' ,SUBSTRING(WPE_PERIODESAUV,1,6)||"-3-07-01-000001",'
        +' IIF(((WPE_UNPDRPAR="030") AND (SUBSTRING(WPE_PERIODESAUV,8,1)="1"))'
        +' ,SUBSTRING(WPE_PERIODESAUV,1,8)||"-01-01-000001",'
        +' IIF(((WPE_UNPDRPAR="030") AND (SUBSTRING(WPE_PERIODESAUV,8,1)="2"))'
        +' ,SUBSTRING(WPE_PERIODESAUV,1,8)||"-04-01-000001",'
        +' IIF(((WPE_UNPDRPAR="030") AND (SUBSTRING(WPE_PERIODESAUV,8,1)="3"))'
        +' ,SUBSTRING(WPE_PERIODESAUV,1,8)||"-07-01-000001",'
        +' IIF(((WPE_UNPDRPAR="030") AND (SUBSTRING(WPE_PERIODESAUV,8,1)="4"))'
        +' ,SUBSTRING(WPE_PERIODESAUV,1,8)||"-10-01-000001",'
        +' IIF((WPE_UNPDRPAR="040")'
        +' ,SUBSTRING(WPE_PERIODESAUV,1,11)||"-01-000001",'
        +' IIF((WPE_UNPDRPAR="050"),SUBSTRING(WPE_PERIODESAUV,1,14)||"-000001"'
        +' , WPE_PERIODESAUV)))))))))');

        // Requêtes de mise à jour des nouvelles unités dans la fiche ARTICLE
    {pour le cas ou ...}
    ExecuteSQL('UPDATE MEA SET GME_QUOTITE=1 WHERE GME_QUOTITE=0');
    { Unité de prix vente -> Unité de quantité de vente }
    ExecuteSQLContOnExcept('UPDATE ARTICLE SET GA_UNITEQTEVTE=GA_QUALIFUNITEVTE WHERE (ISNULL(GA_UNITEQTEVTE, "") = "") OR (GA_UNITEQTEVTE="")');
    { Unité de prix vente -> Unité de prix d'achat }
    ExecuteSQLContOnExcept('UPDATE ARTICLE SET GA_UNITEPRIXACH=GA_QUALIFUNITEVTE WHERE (ISNULL(GA_UNITEPRIXACH, "") = "") OR (GA_UNITEPRIXACH="")');
    { Unité de prix vente -> Unité de quantité d'achat }
    ExecuteSQLContOnExcept('UPDATE ARTICLE SET GA_UNITEQTEACH=GA_QUALIFUNITEVTE WHERE (ISNULL(GA_UNITEQTEACH,"") = "") OR (GA_UNITEQTEACH="")');

    // Requêtes de mise à jour des coefficients de conversion dans la fiche ARTICLE
    if GetParamSocSecur('SO_GCGESTUNITEMODE', '001') = '001' then
    begin
      // GA_COEFCONVQTEVTE = QUALIFUNITEVTE / QUALIFUNITESTO
      ExecuteSQLContOnExcept('UPDATE ARTICLE '
               + 'SET GA_COEFCONVQTEVTE = ISNULL( (SELECT GME_QUOTITE FROM MEA WHERE GME_MESURE=GA_QUALIFUNITEVTE), 1) / ISNULL( (SELECT GME_QUOTITE FROM MEA WHERE GME_MESURE=GA_QUALIFUNITESTO), 1) '
               + 'WHERE (ISNULL(GA_QUALIFUNITEVTE, " ") <> " ") '
               +  'AND (ISNULL(GA_QUALIFUNITESTO, " ") <> " ") '
               +  'AND (ISNULL((SELECT GME_QUALIFMESURE FROM MEA WHERE GME_MESURE=GA_QUALIFUNITEVTE), " ") = ISNULL((SELECT GME_QUALIFMESURE FROM MEA WHERE GME_MESURE=GA_QUALIFUNITESTO), " "))');
      // GA_COEFCONVQTEACH = GA_COEFCONVQTEVTE
      ExecuteSQLContOnExcept('UPDATE ARTICLE '
               + 'SET GA_COEFCONVQTEACH = GA_COEFCONVQTEVTE '
               + 'WHERE (ISNULL(GA_QUALIFUNITEVTE, " ") <> " ") '
               +  'AND (ISNULL(GA_QUALIFUNITESTO, " ") <> " ") '
               +  'AND (ISNULL((SELECT GME_QUALIFMESURE FROM MEA WHERE GME_MESURE=GA_QUALIFUNITEVTE), " ") = ISNULL((SELECT GME_QUALIFMESURE FROM MEA WHERE GME_MESURE=GA_QUALIFUNITESTO), " "))');
    end
    else
    begin
    { Coefficient : Quantité de vente vers Quantité de stock }
    ExecuteSQLContOnExcept('UPDATE ARTICLE '
             + 'SET GA_COEFCONVQTEACH = ISNULL( (SELECT GME_QUOTITE FROM MEA WHERE GME_MESURE=GA_UNITEQTEACH), 1) / ISNULL( (SELECT GME_QUOTITE FROM MEA WHERE GME_MESURE=GA_QUALIFUNITESTO), 1) '
             + 'WHERE (ISNULL(GA_UNITEQTEACH, " ") <> " ") '
             +  'AND (ISNULL(GA_QUALIFUNITESTO, " ") <> " ") '
             +  'AND (ISNULL((SELECT GME_QUALIFMESURE FROM MEA WHERE GME_MESURE=GA_UNITEQTEACH), " ") = ISNULL((SELECT GME_QUALIFMESURE FROM MEA WHERE GME_MESURE=GA_QUALIFUNITESTO), " "))');
    { Coefficient : Quantité d'achat vers Quantité de stock }
    ExecuteSQLContOnExcept('UPDATE ARTICLE '
             + 'SET GA_COEFCONVQTEVTE = ISNULL( (SELECT GME_QUOTITE FROM MEA WHERE GME_MESURE=GA_UNITEQTEVTE), 1) / ISNULL( (SELECT GME_QUOTITE FROM MEA WHERE GME_MESURE=GA_QUALIFUNITESTO), 1) '
             + 'WHERE (ISNULL(GA_UNITEQTEVTE, " ") <> " ") '
             + 'AND (ISNULL(GA_QUALIFUNITESTO, " ") <> " ") '
             + 'AND (ISNULL((SELECT GME_QUALIFMESURE FROM MEA WHERE GME_MESURE=GA_UNITEQTEVTE), " ") = ISNULL((SELECT GME_QUALIFMESURE FROM MEA WHERE GME_MESURE=GA_QUALIFUNITESTO), " "))');
    end;
    { Coefficient : Quantité production vers Quantité de stock }
    ExecuteSQLContOnExcept('UPDATE ARTICLE '
             + 'SET GA_COEFPROD = ISNULL( (SELECT GME_QUOTITE FROM MEA WHERE GME_MESURE=GA_UNITEPROD), 1) / ISNULL( (SELECT GME_QUOTITE FROM MEA WHERE GME_MESURE=GA_QUALIFUNITESTO), 1) '
             + 'WHERE (ISNULL(GA_UNITEPROD, " ") <> " ") '
             + 'AND (ISNULL(GA_QUALIFUNITESTO, " ") <> " ") '
             + 'AND (ISNULL((SELECT GME_QUALIFMESURE FROM MEA WHERE GME_MESURE=GA_UNITEPROD), " ") = ISNULL((SELECT GME_QUALIFMESURE FROM MEA WHERE GME_MESURE=GA_QUALIFUNITESTO), " "))');
    { Coefficient : Quantité de conso vers quantité de stock }
    ExecuteSQLContOnExcept('UPDATE ARTICLE '
             + 'SET GA_COEFCONVCONSO = ISNULL( (SELECT GME_QUOTITE FROM MEA WHERE GME_MESURE=GA_UNITECONSO), 1) / ISNULL( (SELECT GME_QUOTITE FROM MEA WHERE GME_MESURE=GA_QUALIFUNITESTO), 1) '
             + 'WHERE (ISNULL(GA_UNITECONSO, " ") <> " ") '
             + 'AND (ISNULL(GA_QUALIFUNITESTO, " ") <> " ") '
             + 'AND (ISNULL((SELECT GME_QUALIFMESURE FROM MEA WHERE GME_MESURE=GA_UNITECONSO), " ") = ISNULL((SELECT GME_QUALIFMESURE FROM MEA WHERE GME_MESURE=GA_QUALIFUNITESTO), " "))');
  end;
End;

Procedure MajVer745 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
     AglNettoieListes('WPARCELEM','WPC_TIERS',nil);
     //GIGA oubli 738
     executeSql ('Update Tierscompl set ytc_codecata=""');
     AglNettoieListes('PGMULVARIABLE', 'PVA_VARIABLE',nil);
     AglNettoieListes('PGCOMPLSALARIEINT', 'PSE_SALARIE',nil);
  end;
end;

Procedure MajVer746 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  if not isOracle then
  begin
// js1 210606 DupliqueFiltre ('MULMVT3N','CPRECHERCHEECR');
  DupliqueFiltre ('CPMODECRN','CPRECHERCHEECR');
  end;

  // Correction arborescence DP
  if IsMonoOuCommune then
  begin
  ExecuteSQLContOnExcept('UPDATE YGEDDICO SET YGD_LIBELLEGED="Données juridiques" WHERE YGD_TAGMENU=75410');
  ExecuteSQLContOnExcept('UPDATE YGEDDICO SET YGD_LIBELLEGED="Intervenants spécifiques" WHERE YGD_TAGMENU=75430');
  ExecuteSQLContOnExcept('UPDATE ANNULIEN ' +
              'SET ANL_RACINE = (SELECT JTP_RACINE FROM JUTYPEPER WHERE ANL_TYPEPER = JTP_TYPEPER) ' +
              'WHERE ANL_TYPEDOS = "DP" AND ANL_TYPEPER <> ""');
  end;
end;

Procedure MajVer747 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  // K BORGHETTI 160506
  AGLNettoieListes('WORDRELIG;WORDRELIG2','WOL_AUTOMATISME',nil);

  //C AYEL 160506
  AGLNettoieListes('CPICCGENERAUX','',nil,'ICG_SOLDEDEBEX;ICG_SOLDEFINEX');
  SupprimeEtat( 'E','GLG','GLH');
  //P DUMET 160506
  ExecuteSQLContOnExcept('UPDATE ABSENCESALARIE SET PCN_TYPEIMPUTE="AC2" WHERE PCN_TYPEMVT="CPA" AND (PCN_TYPECONGE="REP" OR PCN_TYPECONGE="CPA")');
end;

Procedure MajVer748 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  // ANNECY 220506
  AGLNettoieListes('WORDRELIG;WORDRELIG2','WOL_AUTOMATISME',nil);
  //PAYE  230506
  ExecuteSQLContOnExcept('update motifabsence set pma_controlmotif="-"');
  //TRESO 230506
  AglNettoieListes('CPEXPORT', 'E_NUMCFONB;E_CFONBOK;E_DEVISE;', nil);
end;

Procedure MajVer749 ;
Var stLesBases,UneBase : String;
    SavEnableDeShare:boolean;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //MCD 240506
  AglNettoieListes('AFMULACTIVITE', 'ACT_TIERS',nil);
  // p dumet 300506
  AglNettoieListes('PGCRITMAINTIEN', 'PCM_VALCATEGORIE',nil);

  // js1 080606 suppression des parametres
  MajBundlePartage();

  SavEnableDeShare := V_PGI.enableDEShare;
  V_PGI.enableDEShare := True;
  try
  // js1 290506 On mouline les bases multi pour enrichir les
  // paramsocs et la table dossier de la base commune
  // Récupération des bases de regroupement
    stLesBases := GetBasesMS ('##MULTISOC', true);

    // Remplissage table temporaire pour chaque base :
    UneBase := ReadTokenSt (stLesBases) ;

    while UneBase <> '' do
    begin
      UpdateMultiDossier (UneBase);

      UneBase := ReadTokenSt (stLesBases) ;
    end ;
  finally
    V_PGI.enableDEShare := SavEnableDeShare;
  end;

end;

Procedure MajVer750 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  // r rohault 060606
  AglNettoieListes('CPEXPORT', 'E_NUMLIGNE;E_NUMECHE;', nil);
  // mcd 060606
  AglNettoieListes('AFMULLIGNEACOMPTE', 'GL_PUHTDEV;GL_REMISELIGNE;GL_TYPEREMISE',nil);
  AglNettoieListes('AFMULACTIVITEREP', 'ACT_MNTREMISE;ACT_UNITE;ACT_UNITEFAC',nil);
  //aix 060606
  if not IsDossierPCL then
  begin
    if not ExisteSQL('SELECT 1 FROM SOUCHE WHERE SH_TYPE="GES" AND SH_SOUCHE="CDA" ') then
      ExecuteSQLContOnExcept('INSERT INTO SOUCHE values ("GES","CDA","Compteur Demande Achat",'
      +'"Compteur DA","   ",1,"   ","","001","' + string(USDATETIME(iDate1900)) +'","'
      + string(USDATETIME(iDate1900)) + '","-","-","   ","-",1,1,"-","-")' );
    if not ExisteSQL('SELECT 1 FROM SOUCHE WHERE SH_TYPE="GES" AND SH_SOUCHE="COT" ') then
      ExecuteSQLContOnExcept('INSERT INTO SOUCHE values ("GES","COT","Compteur Ordre de Transport","'
      +'Compteur OT","   ",1,"   ","","001","' + string(USDATETIME(iDate1900)) + '","'
      + string(USDATETIME(iDate1900)) +'","-","-","   ","-",1,1,"-","-")' );
    if not ExisteSQL('SELECT 1 FROM CHOIXCOD WHERE CC_TYPE="TRT" AND CC_CODE="1  " ') then
      ExecuteSQLContOnExcept('INSERT INTO CHOIXCOD values ("TRT","1  ","Express","Express","")' );
    if not ExisteSQL('SELECT 1 FROM CHOIXCOD WHERE CC_TYPE="TRT" AND CC_CODE="2  " ') then
      ExecuteSQLContOnExcept('INSERT INTO CHOIXCOD values ("TRT","2  ","Rapide","Rapide","")' );
    if not ExisteSQL('SELECT 1 FROM CHOIXCOD WHERE CC_TYPE="TRT" AND CC_CODE="3  " ') then
      ExecuteSQLContOnExcept('INSERT INTO CHOIXCOD values ("TRT","3  ","Messagerie","Messagerie","")' );
    if not ExisteSQL('SELECT 1 FROM CHOIXCOD WHERE CC_TYPE="TRT" AND CC_CODE="4  " ') then
      ExecuteSQLContOnExcept('INSERT INTO CHOIXCOD values ("TRT","4  ","Monocolis","Monocolis","")' );
    if not ExisteSQL('SELECT 1 FROM CHOIXCOD WHERE CC_TYPE="TEX" AND CC_CODE="C  " ') then
      ExecuteSQLContOnExcept('INSERT INTO CHOIXCOD values ("TEX","C  ","Colis","Colis","")' );
    if not ExisteSQL('SELECT 1 FROM CHOIXCOD WHERE CC_TYPE="TEX" AND CC_CODE="S  " ') then
      ExecuteSQLContOnExcept('INSERT INTO CHOIXCOD values ("TEX","S  ","Livraison Samedi","Livraison Samedi","")' );
    // r rohault 060606
    ExecuteSQLContOnExcept('DELETE FROM modedata WHERE md_cle LIKE "EBANEMA%"');
    ExecuteSQLContOnExcept('DELETE FROM modeles WHERE mo_nature = "BAN" AND mo_code = "EMA"');
    ExecuteSQLContOnExcept('UPDATE SECTION SET S_FINCHANTIER = "'+USDateTime(idate2099)+'" WHERE S_FINCHANTIER IS NULL');
    ExecuteSQLContOnExcept('UPDATE SECTION SET S_DEBCHANTIER = "'+USDateTime(idate1900)+'" WHERE S_DEBCHANTIER IS NULL');
    // annecy 060606 puis 120606 : plantage oracle ==> rajout ||
    ExecuteSQLContOnExcept('UPDATE WPDRTET SET WPE_PERIODESAUV = IIF((WPE_UNPDRPAR="010")'
    +' ,SUBSTRING(WPE_PERIODESAUV,1,4)||"-1-1-01-01-000001",'
    +' IIF(((WPE_UNPDRPAR="020") AND (SUBSTRING(WPE_PERIODESAUV,6,1)="1"))'
    +' ,SUBSTRING(WPE_PERIODESAUV,1,6)||"-1-01-01-000001",'
    +' IIF(((WPE_UNPDRPAR="020") AND (SUBSTRING(WPE_PERIODESAUV,6,1)="2"))'
    +' ,SUBSTRING(WPE_PERIODESAUV,1,6)||"-3-07-01-000001",'
    +' IIF(((WPE_UNPDRPAR="030") AND (SUBSTRING(WPE_PERIODESAUV,8,1)="1"))'
    +' ,SUBSTRING(WPE_PERIODESAUV,1,8)||"-01-01-000001",'
    +' IIF(((WPE_UNPDRPAR="030") AND (SUBSTRING(WPE_PERIODESAUV,8,1)="2"))'
    +' ,SUBSTRING(WPE_PERIODESAUV,1,8)||"-04-01-000001",'
    +' IIF(((WPE_UNPDRPAR="030") AND (SUBSTRING(WPE_PERIODESAUV,8,1)="3"))'
    +' ,SUBSTRING(WPE_PERIODESAUV,1,8)||"-07-01-000001",'
    +' IIF(((WPE_UNPDRPAR="030") AND (SUBSTRING(WPE_PERIODESAUV,8,1)="4"))'
    +' ,SUBSTRING(WPE_PERIODESAUV,1,8)||"-10-01-000001",'
    +' IIF((WPE_UNPDRPAR="040"),SUBSTRING(WPE_PERIODESAUV,1,11)||"-01-000001",'
    +' IIF((WPE_UNPDRPAR="050"),SUBSTRING(WPE_PERIODESAUV,1,14)||"-000001",'
    +' IIF((WPE_UNPDRPAR="060"), WPE_PERIODESAUV,SUBSTRING(WPE_PERIODESAUV,1,4)||"-1-1-01-01-000001"))))))))))');

    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_LIBELLE="Attendu de sous-traitance d''achat" WHERE GSN_QUALIFMVT="ASA"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_LIBELLE="Entrée de sous-traitance d''achat" WHERE GSN_QUALIFMVT="ESA"');

    if ExisteSQL('SELECT 1 FROM PARPIECE WHERE GPP_NATUREPIECEG="CMF"') then
    begin
      ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_LIBELLE="Commande marché achat" WHERE GPP_NATUREPIECEG="CMF"');
    end;

    if ExisteSQL('SELECT 1 FROM PARPIECE WHERE GPP_NATUREPIECEG="CMC"') then
    begin
      ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_LIBELLE="Commande marché vente" WHERE GPP_NATUREPIECEG="CMC"');
    end;

    // gc 060606
    executesql ('UPDATE STKNATURE SET GSN_SIGNEMVT="MIX" WHERE GSN_QUALIFMVT="EVE"');

    if not ExisteSql ('Select gpp_libelle from parpiece where gpp_naturepieceg="EEX"') then
    ExecuteSql ('INSERT INTO PARPIECE (GPP_NATUREPIECEG) values( "EEX")');

    ExecuteSql ('update parpiece set '+
    '   GPP_LIBELLE="Entrée exceptionnelle", GPP_NIVEAUPARAM="", GPP_NATUREORIGINE="", GPP_MENU="", GPP_NATURETIERS="CLI", ' +
    '   GPP_SENSPIECE="ENT", GPP_QTEPLUS="ESE;PHY;", GPP_QTEMOINS="", GPP_MAJPRIXVALO="", GPP_LOT="-",  ' +
    '   GPP_CALCRUPTURE="AUC", GPP_FORCERUPTURE="-", GPP_RELIQUAT="-", GPP_RECALCULPRIX="-", GPP_GEREECHEANCE="-",  ' +
    '   GPP_TYPEECRCPTA="RIE", GPP_TYPEPASSCPTA="REE", GPP_TYPEPASSACC="REE", GPP_TYPEECRSTOCK="RIE", GPP_NATURESUIVANTE="", ' +
    '   GPP_APPELPRIX="DPA", GPP_MODIFCOUT="-", GPP_CONDITIONTARIF="-", GPP_HISTORIQUE="X", ' +
    '   GPP_ACOMPTE="-", GPP_VALLIBART1="AUC", GPP_VALLIBART2="AUC", GPP_VALLIBART3="AUC", GPP_DATELIBART1="AUC",  ' +
    '   GPP_DATELIBART2="AUC", GPP_DATELIBART3="AUC", GPP_VALLIBTIERS1="AUC", GPP_VALLIBTIERS2="AUC", GPP_VALLIBTIERS3="AUC", ' +
    '   GPP_DATELIBTIERS1="AUC", GPP_DATELIBTIERS2="AUC", GPP_DATELIBTIERS3="AUC", GPP_VALLIBCOM1="", GPP_VALLIBCOM2="",  ' +
    '   GPP_VALLIBCOM3="", GPP_DATELIBCOM1="", GPP_DATELIBCOM2="", GPP_DATELIBCOM3="", GPP_CUMULCOM1="-", GPP_CUMULCOM2="-",  ' +
    '   GPP_CUMULCOM3="-", GPP_CUMULART1="-", GPP_CUMULART2="-", GPP_CUMULART3="-", GPP_CUMULTIERS1="-", GPP_CUMULTIERS2="-", ' +
    '   GPP_CUMULTIERS3="-", GPP_TYPEARTICLE="MAR", GPP_EDITIONNOMEN="AUC", GPP_TAXE="-", GPP_VISA="-", GPP_CONTROLEMARGE="", ' +
    '   GPP_MESSAGEEDIOUT="", GPP_MESSAGEEDIIN="", GPP_SOUCHE="GME", GPP_NBEXEMPLAIRE=0, GPP_EQUIPIECE="", GPP_ENCOURS="-", ' +
    '   GPP_LIENAFFAIRE="-", GPP_LIENTACHE="-", GPP_TYPEACTIVITE="", GPP_PREVUAFFAIRE="-", GPP_OBJETDIM="-", ' +
    '   GPP_LISTESAISIE="GCSAISIEEX", GPP_PRIORECHART1="ART", GPP_PRIORECHART2="BAR", GPP_PRIORECHART3="AUC",  ' +
    '   GPP_CONTRECHART1="-", GPP_CONTRECHART2="-", GPP_CONTRECHART3="-", GPP_TYPEDIMOBLI1="-", GPP_TYPEDIMOBLI2="-",  ' +
    '   GPP_TYPEDIMOBLI3="-", GPP_TYPEDIMOBLI4="-", GPP_TYPEDIMOBLI5="-", GPP_VENTEACHAT="AUT", GPP_IMPIMMEDIATE="-",  ' +
    '   GPP_IMPMODELE="", GPP_IMPETIQ="-", GPP_ETATETIQ="", GPP_COMMENTENT="", GPP_COMMENTPIED="", ' +
    '   GPP_DUPLICPIECE="", GPP_ACTIONFINI="ENR", GPP_LISTEAFFAIRE="", GPP_IFL1="001", GPP_IFL2="002", GPP_IFL3="005", ' +
    '   GPP_IFL4="015", GPP_IFL5="", GPP_IFL6="", GPP_IFL7="", GPP_IFL8="", GPP_MONTANTVISA=0, GPP_QUALIFMVT="EXC", ' +
    '   GPP_BLOBLIENART="", GPP_BLOBLIENTIERS="", GPP_JOURNALCPTA="", GPP_NATURECPTA="", GPP_COMPANALLIGNE="SAN", ' +
    '   GPP_COMPANALPIED="SAN", GPP_COMPSTOCKLIGNE="SAN", GPP_COMPSTOCKPIED="SAN", GPP_ESTAVOIR="-", GPP_IMPETAT="",  ' +
    '   GPP_DIMSAISIE="GEN", GPP_AFAFFECTTB="", GPP_OUVREAUTOPORT="-", GPP_MODEGROUPEPORT="", GPP_TIERS="",  ' +
    '   GPP_PARAMGRILLEDIM="-", GPP_ACHATACTIVITE="-", GPP_PARAMDIM="-", GPP_ARTFOURPRIN="-", GPP_ARTSTOCK="-",  ' +
    '   GPP_OBLIGEREGLE="-", GPP_FILTREARTCH="", ' +
    '   GPP_FILTREARTVAL="", GPP_AFFPIECETABLE="-", GPP_PIECETABLE1="", GPP_PIECETABLE2="", GPP_PIECETABLE3="-", ' +
    '   GPP_CODPIECEDEF1="", GPP_CODPIECEDEF2="", GPP_CODPIECEDEF3="", GPP_CODEPIECEOBL1="-", GPP_CODEPIECEOBL2="-", ' +
    '   GPP_CODEPIECEOBL3="-", GPP_CONTREMARQUE="-", GPP_CONTREMREF="-", GPP_NUMEROSERIE="-", GPP_ECLATEDOMAINE="-",  ' +
    '   GPP_RACINELIBECR1="", GPP_RACINELIBECR2="", GPP_VALEURLIBECR1="", GPP_VALEURLIBECR2="", GPP_RACINEREFINT1="", ' +
    '   GPP_RACINEREFINT2="", GPP_VALEURREFINT1="", GPP_VALEURREFINT2="", GPP_TYPEECRALIM="", GPP_REGROUPCPTA="",  ' +
    '   GPP_CPTCENTRAL="-", GPP_GESTIONGRATUIT="-", GPP_MULTIGRILLE="X", GPP_PROCLI="-", GPP_TYPEFACT="", GPP_NATPIECEANNUL="", ' +
    '   GPP_MASQUERNATURE="-", GPP_APERCUAVIMP="X", GPP_APERCUAVETIQ="-", GPP_VALMODELE="-", GPP_CONTEXTES="GC;", ' +
    '   GPP_ECLATEAFFAIRE="-", GPP_TYPEPRESENT=0, GPP_RECUPPRE="PRE", GPP_TYPEPRESDOC="DEF", GPP_APPLICRG="-", ' +
    '   GPP_TRSFACHAT="-", GPP_TRSFVENTE="-", GPP_PIECEPILOTE="-", GPP_TYPECOMMERCIAL="REP", GPP_FILTRECOMM="",  ' +
    '   GPP_MAJINFOTIERS="-", GPP_INSERTLIG="X", GPP_REFINTEXT="INT", GPP_REFINTCTRL="000", GPP_REFEXTCTRL="000", GPP_CFGART="-", ' +
    '   GPP_CFGARTASSIST="", GPP_TARIFMODULE="", GPP_PRIXNULOK="-", GPP_PILOTEORDRE="-", GPP_IMPBESOIN="-",  ' +
    '   GPP_IMPAUTOETATCBN="-", GPP_IMPAUTOBESCBN="-", GPP_PIECEEDI="-", GPP_MONTANTMINI=0, GPP_STKQUALIFMVT="EEX", ' +
    '   GPP_MODEECHEANCES="RS", GPP_NATUREREPRISE="", GPP_INFOSCOMPL="-", GPP_ACTIVITEPUPR="-", GPP_REGROUPE="X", ' +
    '   GPP_SOLDETRANSFO="-", GPP_RECHTARIF501="", GPP_CHAINAGE="", GPP_MODPLANIFIABLE="", GPP_MODELEWORD="", ' +
    '   GPP_INFOSCPLPIECE="-", GPP_TARIFGENDATE="010", GPP_TARIFGENDEPOT="-", GPP_TARIFGENSAISIE="010", GPP_TARIFGENSPECIA="010", ' +
    '   GPP_TARIFGENTRANSF="010", GPP_GEREARTICLELIE="", GPP_FAR_FAE="", GPP_PIECESAV="-", GPP_REPRISEENTAFF="",  ' +
    '   GPP_REPRISELIGAFF="" Where gpp_naturepieceg="EEX"');

    if not ExisteSql ('Select gpp_libelle from parpiece where gpp_naturepieceg="SEX"') then
      ExecuteSql ('INSERT INTO PARPIECE (GPP_NATUREPIECEG) values( "SEX")');
    ExecuteSql ('update parpiece set'+
    '   GPP_LIBELLE="sortie exceptionnelle", GPP_NIVEAUPARAM="", GPP_NATUREORIGINE="", GPP_MENU="", GPP_NATURETIERS="CLI",  ' +
    '   GPP_SENSPIECE="SOR", GPP_QTEPLUS="LC;", GPP_QTEMOINS="ESE;PHY;", GPP_MAJPRIXVALO="", GPP_LOT="-",  ' +
    '   GPP_CALCRUPTURE="PHY", GPP_FORCERUPTURE="-", GPP_RELIQUAT="-", GPP_RECALCULPRIX="-", GPP_GEREECHEANCE="-",  ' +
    '   GPP_TYPEECRCPTA="RIE", GPP_TYPEPASSCPTA="REE", GPP_TYPEPASSACC="REE", GPP_TYPEECRSTOCK="RIE", GPP_NATURESUIVANTE="", ' +
    '   GPP_APPELPRIX="PUH", GPP_MODIFCOUT="-", GPP_CONDITIONTARIF="-", GPP_HISTORIQUE="X", ' +
    '   GPP_ACOMPTE="-", GPP_VALLIBART1="AUC", GPP_VALLIBART2="AUC", GPP_VALLIBART3="AUC", GPP_DATELIBART1="AUC",  ' +
    '   GPP_DATELIBART2="AUC", GPP_DATELIBART3="AUC", GPP_VALLIBTIERS1="AUC", GPP_VALLIBTIERS2="AUC", GPP_VALLIBTIERS3="AUC", ' +
    '   GPP_DATELIBTIERS1="AUC", GPP_DATELIBTIERS2="AUC", GPP_DATELIBTIERS3="AUC", GPP_VALLIBCOM1="", GPP_VALLIBCOM2="",  ' +
    '   GPP_VALLIBCOM3="", GPP_DATELIBCOM1="", GPP_DATELIBCOM2="", GPP_DATELIBCOM3="", GPP_CUMULCOM1="-", GPP_CUMULCOM2="-",  ' +
    '   GPP_CUMULCOM3="-", GPP_CUMULART1="-", GPP_CUMULART2="-", GPP_CUMULART3="-", GPP_CUMULTIERS1="-", GPP_CUMULTIERS2="-",  ' +
    '   GPP_CUMULTIERS3="-", GPP_TYPEARTICLE="MAR", GPP_EDITIONNOMEN="AUC", GPP_TAXE="-", GPP_VISA="-", GPP_CONTROLEMARGE="",  ' +
    '   GPP_MESSAGEEDIOUT="", GPP_MESSAGEEDIIN="", GPP_SOUCHE="GME", GPP_NBEXEMPLAIRE=0, GPP_EQUIPIECE="", GPP_ENCOURS="-", ' +
    '   GPP_LIENAFFAIRE="-", GPP_LIENTACHE="-", GPP_TYPEACTIVITE="", GPP_PREVUAFFAIRE="-", GPP_OBJETDIM="-", ' +
    '   GPP_LISTESAISIE="GCSAISIEEX", GPP_PRIORECHART1="ART", GPP_PRIORECHART2="BAR", GPP_PRIORECHART3="AUC",  ' +
    '   GPP_CONTRECHART1="-", GPP_CONTRECHART2="-", GPP_CONTRECHART3="-", GPP_TYPEDIMOBLI1="-", GPP_TYPEDIMOBLI2="-",  ' +
    '   GPP_TYPEDIMOBLI3="-", GPP_TYPEDIMOBLI4="-", GPP_TYPEDIMOBLI5="-", GPP_VENTEACHAT="AUT", GPP_IMPIMMEDIATE="-",  ' +
    '   GPP_IMPMODELE="", GPP_IMPETIQ="-", GPP_ETATETIQ="", GPP_COMMENTENT="", GPP_COMMENTPIED="", ' +
    '   GPP_DUPLICPIECE="", GPP_ACTIONFINI="ENR", GPP_LISTEAFFAIRE="", GPP_IFL1="001", GPP_IFL2="002", GPP_IFL3="005", ' +
    '   GPP_IFL4="015", GPP_IFL5="", GPP_IFL6="", GPP_IFL7="", GPP_IFL8="", GPP_MONTANTVISA=0, GPP_QUALIFMVT="SEC", ' +
    '   GPP_BLOBLIENART="", GPP_BLOBLIENTIERS="", GPP_JOURNALCPTA="", GPP_NATURECPTA="", GPP_COMPANALLIGNE="SAN", ' +
    '   GPP_COMPANALPIED="SAN", GPP_COMPSTOCKLIGNE="SAN", GPP_COMPSTOCKPIED="SAN", GPP_ESTAVOIR="-", GPP_IMPETAT="",    ' +
    '   GPP_DIMSAISIE="GEN", GPP_AFAFFECTTB="", GPP_OUVREAUTOPORT="-", GPP_MODEGROUPEPORT="", GPP_TIERS="",  ' +
    '   GPP_PARAMGRILLEDIM="-", GPP_ACHATACTIVITE="-", GPP_PARAMDIM="-", GPP_ARTFOURPRIN="-", GPP_ARTSTOCK="-",  ' +
    '   GPP_OBLIGEREGLE="-", GPP_FILTREARTCH="", ' +
    '   GPP_FILTREARTVAL="", GPP_AFFPIECETABLE="-", GPP_PIECETABLE1="", GPP_PIECETABLE2="", GPP_PIECETABLE3="-", ' +
    '   GPP_CODPIECEDEF1="", GPP_CODPIECEDEF2="", GPP_CODPIECEDEF3="", GPP_CODEPIECEOBL1="-", GPP_CODEPIECEOBL2="-",  ' +
    '   GPP_CODEPIECEOBL3="-", GPP_CONTREMARQUE="-", GPP_CONTREMREF="-", GPP_NUMEROSERIE="-", GPP_ECLATEDOMAINE="-",    ' +
    '   GPP_RACINELIBECR1="", GPP_RACINELIBECR2="", GPP_VALEURLIBECR1="", GPP_VALEURLIBECR2="", GPP_RACINEREFINT1="", ' +
    '   GPP_RACINEREFINT2="", GPP_VALEURREFINT1="", GPP_VALEURREFINT2="", GPP_TYPEECRALIM="", GPP_REGROUPCPTA="",  ' +
    '   GPP_CPTCENTRAL="-", GPP_GESTIONGRATUIT="-", GPP_MULTIGRILLE="X", GPP_PROCLI="-", GPP_TYPEFACT="", GPP_NATPIECEANNUL="", ' +
    '   GPP_MASQUERNATURE="-", GPP_APERCUAVIMP="X", GPP_APERCUAVETIQ="-", GPP_VALMODELE="-", GPP_CONTEXTES="GC;", ' +
    '   GPP_ECLATEAFFAIRE="-", GPP_TYPEPRESENT=0, GPP_RECUPPRE="PRE", GPP_TYPEPRESDOC="DEF", GPP_APPLICRG="-", ' +
    '   GPP_TRSFACHAT="-", GPP_TRSFVENTE="-", GPP_PIECEPILOTE="-", GPP_TYPECOMMERCIAL="REP", GPP_FILTRECOMM="",  ' +
    '   GPP_MAJINFOTIERS="-", GPP_INSERTLIG="X", GPP_REFINTEXT="INT", GPP_REFINTCTRL="000", GPP_REFEXTCTRL="000", GPP_CFGART="-", ' +
    '   GPP_CFGARTASSIST="", GPP_TARIFMODULE="", GPP_PRIXNULOK="-", GPP_PILOTEORDRE="-", GPP_IMPBESOIN="-",  ' +
    '   GPP_IMPAUTOETATCBN="-", GPP_IMPAUTOBESCBN="-", GPP_PIECEEDI="-", GPP_MONTANTMINI=0, GPP_STKQUALIFMVT="SEX", ' +
    '   GPP_MODEECHEANCES="RS", GPP_NATUREREPRISE="", GPP_INFOSCOMPL="-", GPP_ACTIVITEPUPR="-", GPP_REGROUPE="X", ' +
    '   GPP_SOLDETRANSFO="-", GPP_RECHTARIF501="", GPP_CHAINAGE="", GPP_MODPLANIFIABLE="", GPP_MODELEWORD="", ' +
    '   GPP_INFOSCPLPIECE="-", GPP_TARIFGENDATE="010", GPP_TARIFGENDEPOT="-", GPP_TARIFGENSAISIE="010", GPP_TARIFGENSPECIA="010", ' +
    '   GPP_TARIFGENTRANSF="010", GPP_GEREARTICLELIE="", GPP_FAR_FAE="", GPP_PIECESAV="-", GPP_REPRISEENTAFF="",  ' +
    '   GPP_REPRISELIGAFF="" Where gpp_naturepieceg="SEX"');

    //mcd 080606
    ExecuteSQLContOnExcept('update PARPIECE set gpp_listeaffaire="AFSAISIEFCDE"  where gpp_naturepieceg="PRF" or gpp_naturepieceg="DEF"');
   //d brosset 140606
    executesql ('update parpiece set gpp_majprixvalo="DPA;DPR;PPA;PPR;" where gpp_naturepieceg="EEX"');

  end;

  // p dumet 060606
  ExecuteSQLContOnExcept('UPDATE CRITMAINTIEN SET PCM_RUBGARANTIE=PCM_RUBMAINTIEN WHERE PCM_RUBGARANTIE=""');
  AglNettoieListes('PGMULCONGES', 'PSA_TRAVAILN1;PSA_TRAVAILN2;PSA_TRAVAILN3;PSA_TRAVAILN4',nil);
//js1 230606 pour souad  psa_confidentiel==>  psd_confidentiel
  AglNettoieListes('PGVIREMENTACP', 'PSD_CONFIDENTIEL',nil);
  ExecuteSQLContOnExcept('UPDATE DEPORTSAL SET PSE_ISNUMIDENT=(SELECT PSA_NUMEROSS FROM SALARIES WHERE PSA_SALARIE=PSE_SALARIE) WHERE PSE_ISNUMIDENT=""');

  // CA 070606
  If GetParamSocSecur('SO_CPLIENGAMME','') = 'S1' then
	Setparamsoc ('SO_CPMODESYNCHRO', TRUE);

  //aix 080606
  AglNettoieListes('GCPRODUITTRA', 'GTC_CODEPRODTRA;GTC_LIBELLE;GTC_REGIMETRA', nil,'');
  AglNettoieListes('GCMULPIECEEXPE', 'GP_ETABLISSEMENT;GPA_TRANSPORTEUR;GPL_SOUCHEOT;(@SI([GPL_NUMBONTRA]<>0~[GPL_NUMBONTRA]~));GPL_NUMBONTRA;GP_NATUREPIECEG;GP_SOUCHE;GP_NUMERO;GP_TIERSLIVRE;GP_TIERS;GP_DATELIVRAISON;'+
    'GP_DATEPIECE;GP_VIVANTE;GP_INDICEG;GPL_NBCOLIS;GPL_POIDSPESEE;GPL_QUALIFPOIDS;GPL_DATEOT', nil,'');
  AglNettoieListes('GCMULPIECEEXPIMP', 'GP_ETABLISSEMENT;GPA_TRANSPORTEUR;GPL_SOUCHEOT;GPL_NUMBONTRA;GPL_DATEOT;NBCOLIS;POIDSTRANSPORT;LIBUNITE;YTC_MODELEBON;YTC_MODELE;YTC_EDITRA;YTC_TYPEECHANGE;YTC_MODELETXT', nil,'');

{
  //gilles costes 130606
  AGLNettoieFiltre('TCPBALGEN', 'E_EXERCICE;E_ETABLISSEMENT;E_DEVISE', 'EXERCICE;ETABLISSEMENT;DEVISE', False);
  AGLNettoieFiltre('TCPBALGEN', 'E_EXERCICE;E_ETABLISSEMENT;E_DEVISE', 'EXERCICE;ETABLISSEMENT;DEVISE', True);
  AGLNettoieFiltre('TCPBALAUXI', 'E_EXERCICE;E_ETABLISSEMENT;E_DEVISE', 'EXERCICE;ETABLISSEMENT;DEVISE', False);
  AGLNettoieFiltre('TCPBALAUXI', 'E_EXERCICE;E_ETABLISSEMENT;E_DEVISE', 'EXERCICE;ETABLISSEMENT;DEVISE', True);
  AGLNettoieFiltre('TCPBALANAL', 'E_EXERCICE;E_ETABLISSEMENT;E_DEVISE', 'EXERCICE;ETABLISSEMENT;DEVISE', False);
  AGLNettoieFiltre('TCPBALANAL', 'E_EXERCICE;E_ETABLISSEMENT;E_DEVISE', 'EXERCICE;ETABLISSEMENT;DEVISE', True);
  AGLNettoieFiltre('TCPBALAGEE', 'E_ETABLISSEMENT;E_DEVISE', 'ETABLISSEMENT;DEVISE', False);
  AGLNettoieFiltre('TCPBALAGEE', 'E_ETABLISSEMENT;E_DEVISE', 'ETABLISSEMENT;DEVISE', True);
}

  //p dumet 150606
  AglNettoieListes('PGENVOIDUCS', 'PES_SUPPORTEMIS',nil);

  // js1 pour o lambert 260606
  ExecuteSQLContOnExcept('UPDATE MENU SET MN_VERSIONDEV="-" WHERE MN_1=52 AND MN_TAG IN(52010,7677,7261)');
  ExecuteSQLContOnExcept('UPDATE MENU SET MN_VERSIONDEV="-" WHERE MN_1=54 AND MN_TAG IN(54710,54720)');

  // C.B 25/08/2006
  AGLNETTOIELISTES('AFLIGPLANNING', 'APL_HEUREDEB_PLA;APL_HEUREFIN_PLA', nil, 'APL_HEUREDEBPLA;APL_HEUREFINPLA');

End;

Procedure MajVer800 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //mcd 040706
  ExecuteSql ('UPDATE TIERSCOMPL SET YTC_MOTIFFERMETURE=""');
  InsertChoixCode('ATU', '038', 'des tiers', '', 'des clients');
end;

//=====================================
Procedure MajVer801 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  AglNettoieListesPlus('GCSTKPHYSIQUE','GSM_ETATMVT;GSM_ARTICLE',nil,False);

  if not IsDossierPCL then
  begin
    //kb270706
    InsertTablettesRQ;
    ExecuteSQLContOnExcept('UPDATE LISTEINVLIG SET GIL_VALLIBRE1=0,GIL_VALLIBRE2=0,GIL_VALLIBRE3=0 WHERE GIL_VALLIBRE1 IS NULL');

    ExecuteSQLContOnExcept('UPDATE WPDRTYPE SET WRT_AVECCATALOGUE="-" ,WRT_AVECDISPO="-"'
    +',WRT_AVECARTICLE=iif((WRT_TVACHAT="01" OR WRT_TVACHAT="21"),"-","X")'
    +',WRT_AVECARTICLEDEF=iif(WRT_TVACHAT="01","X","-")'
    +',WRT_AVECARTICLESIM=iif(WRT_TVACHAT="21","X","-")'
    +',WRT_TVACHAT=iif((WRT_TVACHAT="01" OR WRT_TVACHAT="21"),"",WRT_TVACHAT)');

    ExecuteSQLContOnExcept('UPDATE WPARAM SET WPA_BOOLEAN11="-" ,WPA_BOOLEAN12="-"'
    +',WPA_BOOLEAN13=iif((WPA_COMBO04="01" OR WPA_COMBO04="21"),"-","X")'
    +',WPA_BOOLEAN14=iif(WPA_COMBO04="01","X","-")'
    +',WPA_BOOLEAN15=iif(WPA_COMBO04="21","X","-")'
    +',WPA_COMBO04=iif((WPA_COMBO04="01" OR WPA_COMBO04="21"),"",WPA_COMBO04)'
    +' WHERE WPA_CODEPARAM="PRIXDEREVIENT"');

    ExecuteSQLContOnExcept('UPDATE WPDRTET SET WPE_AVECCATALOGUE="-",WPE_AVECDISPO="-"'
    +',WPE_AVECARTICLE=iif((WPE_TVACHAT="01" OR WPE_TVACHAT="21"),"-","X")'
    +',WPE_AVECARTICLEDEF=iif(WPE_TVACHAT="01","X","-")'
    +',WPE_AVECARTICLESIM=iif(WPE_TVACHAT="21","X","-")'
    +',WPE_TVACHAT=iif((WPE_TVACHAT="01" OR WPE_TVACHAT="21"),"",WPE_TVACHAT)');

    ExecuteSQLContOnExcept( 'UPDATE WPARAMFONCTION'
              + ' SET WPF_BOOLEAN11="-", WPF_BOOLEAN12="-"'
              + ' WHERE WPF_BOOLEAN11 IS NULL AND WPF_BOOLEAN12 IS NULL');
  end;
  //md 270706
  If IsMonoOuCommune then
  begin
    ExecuteSQLContOnExcept('update DPFISCAL SET DFI_DGE="-", DFI_IMPODIRDATE="'+usDateTime(iDate1900)+'", DFI_ISDECLA="",'+
      'DFI_REGETRANGER="-", DFI_ARRET="", DFI_DATEARRET="'+usDateTime(iDate1900)+'", '+
      'DFI_2065LIEUSIGNE="", DFI_REGIMETVA="", DFI_TVADECLA="",'+
      'DFI_TVARBTCE="-", DFI_TVATERR="-"') ;

    ExecuteSQLContOnExcept('update DPSOCIAL SET DSO_PARTICIPATDATE="'+usDateTime(iDate1900)+'", DSO_ACCORDENTDATE="'+usDateTime(iDate1900)+'",'+
      'DSO_COMITEENTDATE="'+usDateTime(iDate1900)+'"') ;
  end;

end;

Procedure MajVer804 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  if not IsDossierPCL then
  begin
    ExecuteSQLContOnExcept('UPDATE QWRELEVELIG SET QWL_GROUPAGEMES=0 WHERE QWL_GROUPAGEMES IS NULL');
    ExecuteSQLContOnExcept('UPDATE QWHISTORES SET QWH_GROUPAGEMES=0,QWH_SOUSLIGNE=0,QWH_LANCEPHASE="-",QWH_RECEPPHASE="-",'+
       'QWH_LIBOPERATION=IIF((QWH_TYPELIGNEPDR="OM" OR QWH_TYPELIGNEPDR="OL" OR QWH_TYPELIGNEPDR="BAC"),'+
       '(SELECT WOG_LIBELLE FROM WORDREGAMME'+
          ' WHERE WOG_NATURETRAVAIL=QWH_NATURETRAVAIL AND WOG_LIGNEORDRE=QWH_LIGNEORDRE'+
          ' AND WOG_OPECIRC=QWH_OPECIRC AND WOG_NUMOPERGAMME=QWH_NUMOPERGAMME),"")'+
       ' WHERE QWH_GROUPAGEMES IS NULL');
  end;
end;

Procedure MajVer805 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  If IsMonoOuCommune then
  begin
    ExecuteSQLContOnExcept('update DPFISCAL set DFI_DGEENTREE ="'+usDateTime(iDate1900)+'",' +
                  'DFI_DGEESORTIE ="'+usDateTime(iDate1900)+'",' +
                  'DFI_DATEDEBGROUPE ="'+usDateTime(iDate1900)+'",' +
                  'DFI_DATEFINGROUPE ="'+usDateTime(iDate1900)+'",' +
                  'DFI_ACTDIFF ="-",' +
                  'DFI_TPISCCI ="",' +
                  'DFI_TPPLAFOND ="-",' +
                  'DFI_TPPLAFONDSUR ="",' +
                  'DFI_ABATTEPRESSE ="-",' +
                  'DFI_ABATTENWINVEST="-",' +
                  'DFI_ABATTEPLAFOND="-"' ) ;

{ ludovic montavon 150906
 ExecuteSQLContOnExcept('update ANNUBIS set ANB_GROUPE ="",' +
                  'ANB_SOUSGROUPE ="",' +
                  'ANB_POLE ="",' +
                  'ANB_AXE1 ="",' +
                  'ANB_AXE2 ="",' +
                  'ANB_AXE3 =""' );  erreur init inutile le bob correspondant n'est plus intégré}

    if not ExisteSQL('SELECT 1 FROM JUTYPEPER WHERE JTP_TYPEPER="CCI"') then
      ExecuteSQLContOnExcept('insert into JUTYPEPER (JTP_UTILISATEUR, JTP_APPMODIF, JTP_VERSION, JTP_TYPEPER, JTP_TYPEPERABREG, JTP_TYPEPERLIB, JTP_RACINE, JTP_FAMPER, JTP_AFFICHE, JTP_PREDEFINI) ' +
                 'values ("000", "Decla", 1, "CCI", "CCI", "Chambres du commerce et de l''industrie", "CFE", "INS", "CFE", "CEG")');

  end;

   AglNettoieListes('AFCOMPTEURCLT_MUL', 'APX_TIERS;APX_AFFAIRE;APX_AFFAIRE1;APX_AFFAIRE2;APX_AVENANT;APX_COMPTEURGUID',nil);
   AglNettoielIstes('AFEPURAFFAIRE_MUL','AFF_TIERS;AFF_AFFAIRE',nil);
   AGLNETTOIELISTES('AFLIGPLANNING', 'APL_HEUREDEB_PLA;APL_HEUREFIN_PLA', nil, 'APL_HEUREDEBPLA;APL_HEUREFINPLA');
   AglNettoielIstes('AFMULAFFAIRE_CPT','AFF_TIERS;AFF_AFFAIRE1',nil);
   AglNettoielIstes('AFMULCOMPTEURS','ACX_COMPTEURGUID',nil);
	//fin GIGA

  if not IsDossierPCL then
  begin
    	// debut GIGA
	//pour correction bug acceptationa ffaire V7
    ExecuteSql ('UPDATE AFFAIRE set AFF_NATUREPIECEG="AFF" where aff_affaire0="A" and aff_naturepieceg="PAF"');
     // gerard jugde
    ExecuteSQLContOnExcept('UPDATE WGAMMERES  SET WGR_NATURERESMES="PIL" WHERE WGR_NATURERESMES="CPT" ');
    ExecuteSQLContOnExcept('UPDATE WGAOPERRES SET WGS_NATURERESMES="PIL" WHERE WGS_NATURERESMES="CPT" ');
    ExecuteSQLContOnExcept('UPDATE WORDRERES  SET WOR_NATURERESMES="PIL" WHERE WOR_NATURERESMES="CPT" ');
    ExecuteSQLContOnExcept('UPDATE QWHISTORES SET QWH_NATURERESMES="PIL" WHERE QWH_NATURERESMES="CPT" ');
  end;
end;

Procedure MajVer806 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
      // debut GIGA
      //pour correction bug acceptationa ffaire V7
     ExecuteSql ('UPDATE ARTICLECOMPL set GA2_STATUTPLA="" ');
     ExecuteSql ('UPDATE AFPLANNINGPARAM set APP_MODIFIABLE="X" ');
     ExecuteSql ('UPDATE TACHE set ATA_STATUTPLA="" ');
      //fin GIGA
   end;
end;

Procedure MajVer809 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    ExecuteSQLContOnExcept('UPDATE YTARIFSPARAMETRES SET YFO_OKMARQUE="----", YFO_OKCHOIXQUALITE="----", YFO_OKINDICE="----" WHERE YFO_OKMARQUE IS NULL');

    ExecuteSQLContOnExcept('UPDATE YTARIFS SET YTS_MARQUE="", YTS_CASCMARQUE="-", YTS_CASCTOUSMARQUE="-", YTS_CHOIXQUALITE="", YTS_CASCCHOIXQ="-",'
    +' YTS_CASCTOUSCHOIXQ="-", YTS_INDICEARTICLE="", YTS_CASCINDICE="-", YTS_CASCTOUSINDICE="-" WHERE YTS_MARQUE IS NULL');

    ExecuteSQLContOnExcept('UPDATE LIGNEFRAIS SET LF_ORIFORF=SUBSTRING(LF_ORIFORF,1,65)||"-...|"'
    +'||iif((SUBSTRING(LF_ORIFORF,67,65)<>""),SUBSTRING(LF_ORIFORF,67,65)||"-...|"||SUBSTRING(LF_ORIFORF,133,18),"")'
    +' WHERE LF_ORIFORF<>"" AND LF_ORIFORF<>"m" AND ISNULL(SUBSTRING(LF_ORIFORF,66,5)," ")<>"-...|"');

    ExecuteSQLContOnExcept('UPDATE LIGNEFRAIS SET LF_ORIFIXE=SUBSTRING(LF_ORIFIXE,1,65)||"-...|"'
    +'||iif((SUBSTRING(LF_ORIFIXE,67,65)<>""),SUBSTRING(LF_ORIFIXE,67,65)||"-...|"||SUBSTRING(LF_ORIFIXE,133,18),"")'
    +' WHERE LF_ORIFIXE<>"" AND LF_ORIFIXE<>"m" AND ISNULL(SUBSTRING(LF_ORIFIXE,66,5)," ")<>"-...|"');

    ExecuteSQLContOnExcept('UPDATE LIGNEFRAIS SET LF_ORIREMMT=SUBSTRING(LF_ORIREMMT,1,65)||"-...|"'
    +'||iif((SUBSTRING(LF_ORIREMMT,67,65)<>""),SUBSTRING(LF_ORIREMMT,67,65)||"-...|"||SUBSTRING(LF_ORIREMMT,133,18),"")'
    +' WHERE LF_ORIREMMT<>"" AND LF_ORIREMMT<>"m" AND ISNULL(SUBSTRING(LF_ORIREMMT,66,5)," ")<>"-...|"');

    ExecuteSQLContOnExcept('UPDATE LIGNEFRAIS SET LF_ORICOUTBRUT=SUBSTRING(LF_ORICOUTBRUT,1,65)||"-...|"'
    +'||iif((SUBSTRING(LF_ORICOUTBRUT,67,65)<>""),SUBSTRING(LF_ORICOUTBRUT,67,65)||"-...|"||SUBSTRING(LF_ORICOUTBRUT,133,18),"")'
    +' WHERE LF_ORICOUTBRUT<>"" AND LF_ORICOUTBRUT<>"m" AND ISNULL(SUBSTRING(LF_ORICOUTBRUT,66,5)," ")<>"-...|"');

    ExecuteSQLContOnExcept('UPDATE LIGNEFRAIS SET LF_ORIPOURCENT=SUBSTRING(LF_ORIPOURCENT,1,65)||"-...|"'
    +'||iif((SUBSTRING(LF_ORIPOURCENT,67,65)<>""),SUBSTRING(LF_ORIPOURCENT,67,65)||"-...|"||SUBSTRING(LF_ORIPOURCENT,133,18),"")'
    +' WHERE LF_ORIPOURCENT<>"" AND LF_ORIPOURCENT<>"m" AND ISNULL(SUBSTRING(LF_ORIPOURCENT,66,5)," ")<>"-...|"');

    ExecuteSQLContOnExcept('UPDATE LIGNEFRAIS SET LF_ORICOUTNET=SUBSTRING(LF_ORICOUTNET,1,65)||"-...|"'
    +'||iif((SUBSTRING(LF_ORICOUTNET,67,65)<>""),SUBSTRING(LF_ORICOUTNET,67,65)||"-...|"||SUBSTRING(LF_ORICOUTNET,133,18),"")'
    +' WHERE LF_ORICOUTNET<>"" AND LF_ORICOUTNET<>"m" AND ISNULL(SUBSTRING(LF_ORICOUTNET,66,5)," ")<>"-...|"');
{$IFNDEF BTP}
    if (V_PGI.Driver <> dbORACLE7 ) then
      MajPoidsTarifs;
{$ENDIF}
  end;
end;

Procedure MajVer811 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  if not IsDossierPCL then
  begin
    //g jugde 25102006
    ExecuteSQLContOnExcept('UPDATE QWHISTORES SET QWH_PHASE=ISNULL((SELECT WOP_PHASE FROM WORDREPHASE '+
      ' WHERE (WOP_NATURETRAVAIL=QWH_NATURETRAVAIL) AND (WOP_LIGNEORDRE=QWH_LIGNEORDRE) AND (WOP_OPECIRC=QWH_OPECIRC)) ,"" ) '+
      ' WHERE QWH_PHASE IS NULL');
    //jl sauzet 25102006
    ExecuteSQLContOnExcept('UPDATE ARTICLECOMPL SET GA2_NUMEROSERIEGR="-"');
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_GERESERIEGRP="-"');
  end;
  AglNettoieListesPlus('WORDREPHASE','WOP_DATEDEBTHEO',nil,True);
end;

Procedure MajVer813 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  // M DESGOUTTE 071106 GED V2 : Initialisation des nouveaux champs :
  If IsMonoOuCommune then
  begin
    ExecuteSQLContOnExcept('UPDATE YDOCUMENTS SET YDO_PRIVE="-", YDO_DROITGED=""');
    ExecuteSQLContOnExcept('UPDATE DPDOCUMENT SET DPD_CODEGED="D"||'+DB_MID('DPD_CODEGED',1,3)+', DPD_ARBOGED="D", DPD_GROUPEGED=0');
    ExecuteSQLContOnExcept('UPDATE YGEDDICO SET YGD_CODEGED="D"||'+DB_MID('YGD_CODEGED',1,3)+', YGD_ARBOGED="D"');

    // Propagation dans les tablettes hiérarchiques
    ExecuteSQLContOnExcept('UPDATE YDATATYPETREES SET YDT_MCODE="D"||'+DB_MID('YDT_MCODE',1,3)+', YDT_SCODE="D"||'+DB_MID('YDT_SCODE',1,3)+' WHERE YDT_CODEHDTLINK="YYGEDNIV1GEDNIV2"');
  end;
end;

Procedure MajVer814 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //grc mn garnier 141106
  AglNettoieListes('WPARCELEM', 'WAP_GESTVERSION;WPN_ARTICLEPARC',nil);
  AglNettoieListes('WPARCELEMTIERS', 'WAP_GESTVERSION;WPN_ARTICLEPARC',nil);

  //Jean-Luc SAUZET Le 29/05/2008 Version 8.1.850.86 Demande N° 2473
  {
  //jls 141106
  //-- SUPPRESSION DES LISTES  --//
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE ="WPF_CBNCIRCDEP"');
  }

  //-- AGL Nettoie Liste --//
  AGLNettoieListesPlus('RTQEVALDELAI','RQE_INDICEORI',nil,True);
  AGLNettoieListesPlus('RTQEVALZDM','RQE_QNCNUM;RQE_NATUREORI;RQE_SOUCHEORI;RQE_NUMEROORI;RQE_INDICEORI;RQE_NUMLIGNEORI',nil,True);

  // md 141106
  If IsMonoOuCommune then
  begin
    ExecuteSQLContOnExcept('UPDATE ANNUAIRE SET ANN_REGMAT="AUT" WHERE ANN_REGMAT = "INC"');
    ExecuteSQLContOnExcept('UPDATE ANNUAIRE SET ANN_SITUFAM="MAR" WHERE ANN_SITUFAM IN ("SEP","VEU")');
  end;
end;

Procedure MajVer815 ;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  ExecuteSQLContOnExcept('update choixcod set cc_libelle="Lettre d''affaire" where cc_type="den" and cc_code="lmi"');
  ExecuteSQLContOnExcept('update choixcod set cc_libelle="Proposition d''affaire" where cc_type="den" and cc_code="pmi"');

  if not IsDossierPCL then
  begin
    ExecuteSQLContOnExcept('UPDATE QWHISTORES SET '
      + ' QWH_PROVENANCE="",QWH_REFEXTERNE="", '
      + ' QWH_LIBREQWH1="",QWH_LIBREQWH2="",QWH_LIBREQWH3="", '
      + ' QWH_VALLIBRE1=0,QWH_VALLIBRE2=0,QWH_VALLIBRE3=0, '
      + ' QWH_DATELIBRE1="' + UsDateTime(iDate1900) + '",QWH_DATELIBRE2="' + UsDateTime(iDate1900) + '",QWH_DATELIBRE3="' + UsDateTime(iDate1900) + '", '
      + ' QWH_BOOLLIBRE1="-",QWH_BOOLLIBRE2="-",QWH_BOOLLIBRE3="-", '
      + ' QWH_ALEAPRIME=(IIF(QWH_ALEA="","-",(SELECT QOF_ALEAPRIME FROM QOALEA '
      + ' LEFT JOIN QOFAMILLEALEA ON QOF_FAMILLEALEA=QOA_FAMILLEALEA  WHERE QOA_ALEA=QWH_ALEA))) '
      + 'WHERE QWH_PROVENANCE IS NULL ' );
    ExecuteSQLContOnExcept('UPDATE QWMESFICHESV SET QWY_IDENTIFIANTWOB=0 WHERE QWY_IDENTIFIANTWOB IS NULL ');
  end;
end;



Procedure MajVer816 ;
Begin
  // MCD
  if not IsDossierPCL then
  begin
    ExecuteSql ('UPDATE AFFORMULE SET AFE_DATECREATION = "' + UsDateTime(iDate1900) + '", AFE_DATEMODIF = "' + UsDateTime(iDate1900) + '", AFE_UTILISATEUR = "", AFE_CREATEUR = ""');
    ExecuteSql ('UPDATE AFTACHEJOUR SET ATJ_DATECREATION = "' + UsDateTime(iDate1900) + '", ATJ_DATEMODIF = "' + UsDateTime(iDate1900) + '", ATJ_UTILISATEUR = "", ATJ_CREATEUR = ""');
    ExecuteSql ('UPDATE AFMODELEtache SET AFM_DATECREATION = "' + UsDateTime(iDate1900) + '", AFM_DATEMODIF = "' + UsDateTime(iDate1900) + '", AFM_UTILISATEUR = "", AFM_CREATEUR = ""');
    ExecuteSql ('UPDATE AFPUBLICATION SET AFP_DATECREATION = "' + UsDateTime(iDate1900) + '", AFP_DATEMODIF = "' + UsDateTime(iDate1900) + '", AFP_UTILISATEUR = "", AFP_CREATEUR = ""');
    ExecuteSql ('UPDATE AFINDICE SET AIN_DATECREATION = "' + UsDateTime(iDate1900) + '", AIN_DATEMODIF = "' + UsDateTime(iDate1900) + '", AIN_UTILISATEUR = "", AIN_CREATEUR = ""');
    ExecuteSql ('UPDATE AFPLANNINGPARAM SET APP_DATECREATION = "' + UsDateTime(iDate1900) + '", APP_DATEMODIF = "' + UsDateTime(iDate1900) + '", APP_UTILISATEUR = "", APP_CREATEUR = ""');
    ExecuteSql ('UPDATE AFVALINDICE SET AFV_UTILISATEUR = ""');
    ExecuteSql ('UPDATE AFREVISION SET AFR_UTILISATEUR = ""');
    ExecuteSql ('UPDATE AFPARAMFORMULE SET AFC_DATECREATION = "' + UsDateTime(iDate1900) + '", AFC_DATEMODIF = "' + UsDateTime(iDate1900) + '", AFC_UTILISATEUR = "", AFC_CREATEUR = ""');
  end;
  // MD A Ajouter PGIMAJVER : MODIFICATION DE LA LISTE DPDOCUMENTS_MUL
  //+ correction de la table DPDOCUMENT
  If ISmonoOuCommune then
  begin
    AGLNettoieListes ('DPDOCUMENTS_MUL','YDO_UTILISATEUR',nil,'DPD_UTILISATEUR');

    // Rectification d'un traitement de la 813
    ExecuteSQLContOnExcept('UPDATE DPDOCUMENT SET DPD_CODEGED="###" WHERE DPD_CODEGED="D###"');
  end;
  // JLS
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE="LDC" WHERE DH_PREFIXE="LF" AND DH_NUMCHAMP NOT IN (10,20,30,40,50,90,110,130,150,190)');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE="LDC" WHERE DH_PREFIXE="GPT" AND DH_NUMCHAMP NOT IN (10,20,30,40,85,385,400)');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE="LDZC" WHERE DH_PREFIXE="GPT" AND DH_NUMCHAMP IN (85,385,400)');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE=DH_CONTROLE || "C" WHERE DH_PREFIXE="GPA" AND DH_NUMCHAMP NOT IN (2,4,6,8,10)');

  AGLNettoieListesPlus('WCOMPAREWNLWOB','WOB_ETATBES',nil,True);

  AGLNettoieListes('WPF_CHXQTIERS;WPF_CHXQARTI;WPF_MARQARTI;WPF_MARQTIERS;WPF_AUTOWOG;WPF_AUTOWOL;WPF_AUTOWOP','WPF_ARTICLE');
end;

Procedure MajVer817 ;
Begin

  if not IsDossierPCL then
  begin
  { MCD }
	// debut GIGA
    ExecuteSQLContOnExcept('update afftiers set aft_tiers = (select t_tiers from tiers where t_auxiliaire = aft_auxiliaire) where aft_typeinterv = "CON"');
	//fin GIGA
  // JLS
    AGLNettoieListes('WGAMMECIR' ,'WGC_PHASE');
    AGLNettoieListes('WNOMEDEC'  ,'WND_PHASE');
    AGLNettoieListes('WGAMMELIG1','WGL_PHASELIB');
  end;

end;

Procedure MajVer818 ;
Begin
  { MNG}
  ExecuteSQLContOnExcept('update yalertes set yal_modeblocage="001",yal_predefini="",yal_active="X",yal_datedebut="' + UsDateTime(iDate1900) + '",yal_versionalerte="07.001.001"');

  if not IsDossierPCL then
  begin
    { JPL}
    ExecuteSQLContOnExcept('Update ARTICLECOMPL SET GA2_NUMONU=0,GA2_INDICEONU=0 Where GA2_NUMONU is null and GA2_INDICEONU is null');
    // MCD
    ExecuteSQLContOnExcept('UPDATE ACTIVITE SET ACT_REFCOMPTABLE = ""');
    InsertChoixCode('ATU', '037', 'autres tiers', '', 'autres clients');
  end;
end;

Procedure MajVer819 ;
Begin
  { MES }
  ExecuteSQLContOnExcept('UPDATE RESSOURCE SET ARS_CODEPRODMES="" WHERE ARS_CODEPRODMES IS NULL')
end;

Procedure MajVer820 ;
Begin
  if IsMonoOuCommune then
  begin
    // JP
    ExecuteSQL ('UPDATE YDOCUMENTS SET YDO_DOCTYPE="COM",YDO_URL="" WHERE NOT EXISTS (SELECT 1 FROM YDOCFILES WHERE YDF_DOCGUID=YDO_DOCGUID)');
    ExecuteSQL ('UPDATE YDOCUMENTS SET YDO_DOCTYPE="DOC",YDO_URL="" WHERE YDO_DOCTYPE IS NULL OR YDO_DOCTYPE=""');
    // MD
    // Suppression des fiches de l'arborescence Ged/DP
    ExecuteSQLContOnExcept('DELETE FROM YGEDDICO WHERE YGD_TYPEGED="FIC" OR YGD_TYPEGED="LIS"');
  end;

  if not IsDossierPCL then
  begin
    // JLS
    ExecuteSQLContOnExcept('UPDATE WPDRTYPE SET WRT_PDRENREGWPL="X" WHERE WRT_PDRENREGWPL IS NULL');
    ExecuteSQLContOnExcept('UPDATE WPARAM SET WPA_BOOLEAN16="-" WHERE WPA_BOOLEAN16 IS NULL');
    ExecuteSQLContOnExcept('UPDATE WPARAM SET WPA_BOOLEAN16="X" WHERE WPA_CODEPARAM="PRIXDEREVIENT"');
    // JG
    ExecuteSQLContOnExcept('UPDATE QWRELEVETET SET QWT_TYPETRAVAIL="" WHERE QWT_TYPETRAVAIL IS NULL');
    ExecuteSQLContOnExcept('UPDATE QWRELEVELIG SET QWL_TYPETRAVAIL="",QWL_TPREVSUPSAIS=0,QWL_TPREVSUPHHCC=0 WHERE QWL_TYPETRAVAIL IS NULL');
    ExecuteSQLContOnExcept('UPDATE QWHISTORES SET QWH_TYPETRAVAIL="",QWH_DEDUCEFFIC="-",QWH_TPREVSUPSAIS=0,QWH_TPREVSUPHHCC=0 WHERE QWH_TYPETRAVAIL IS NULL');

    // DBR
    { Rendre le Signe des natures de mouvements MIX }
    ExecuteSql ('UPDATE STKNATURE SET GSN_SIGNEMVT="MIX" WHERE GSN_QUALIFMVT in ("SAC", "SDM", "STR")');
  end;

  // MCD
  ExecuteSQLContOnExcept('update paramsoc set soc_data=(select soc_data from paramsoc where soc_nom="so_afvaloactpr") where soc_nom = "so_afvalofraispr"');
  ExecuteSQLContOnExcept('update paramsoc set soc_data=(select soc_data from paramsoc where soc_nom="so_afvaloactpr") where soc_nom = "so_afvalofourpr"');
  ExecuteSQLContOnExcept('update paramsoc set soc_data=(select soc_data from paramsoc where soc_nom="so_afvaloactpv") where soc_nom = "so_afvalofourpv"');
  ExecuteSQLContOnExcept('update paramsoc set soc_data=(select soc_data from paramsoc where soc_nom="so_afvaloactpv") where soc_nom = "so_afvalofraispv"');
end;

Procedure MajVer821 ;
Begin
  if IsMonoOuCommune then
  begin
    ExecuteSQLContOnExcept('UPDATE JUEVENEMENT SET JEV_ALLDAY="-"');
    ExecuteSQLContOnExcept('UPDATE JUEVTRECURRENCE SET JEC_IGNOREFERIE="-"');
  end;

  //Vincent GALLIOT Demande N° 1528
  {AglNettoieListes('PGDECLARANTATTEST', 'EtablissementsT',nil);}
  //Vincent GALLIOT Demande N° 1583
  {AglNettoieListes('PGDECLARANTATTEST', 'Etablissements',nil);}

  AglNettoieListes('PGENVOIDUCS', 'PES_FICHIERRECU',nil);
  ExecuteSQLContOnExcept('UPDATE ABSENCESALARIE SET PCN_GUID=PGIGUID WHERE (PCN_GUID IS NULL OR PCN_GUID="")')  ;
  AglNettoieListes('TRLISTETRANSAC', 'TCT_NODOSSIER;');
  AglNettoieListes('TRTRANSACOPCVMACH', 'TOP_NODOSSIER;');
  if not IsDossierPCL then
  begin
    // JLS
    ExecuteSQLContOnExcept('DELETE FROM STKMOUVEMENT WHERE GSM_STKTYPEMVT="CLO" AND GSM_QUALIFMVT="CLO" AND GSM_PHYSIQUE=0');
    ExecuteSQLContOnExcept('UPDATE STKMOUVEMENT SET GSM_MONTANTCPTA=0');
  end;
end;

Procedure MajVer822 ;
Begin
  if not IsDossierPCL then
  begin
    // MCD
    ExecuteSQLContOnExcept('update afplanning set apl_guid = pgiguid');
    ExecuteSQLContOnExcept('Update ETABLES set EDT_SENS="REQ" where EDT_PREFIXE <= "999"');
    ExecuteSQLContOnExcept('Update ECHAMPS set EDH_SENS="REQ" where EDH_PREFIXE <= "999"');
    //giga suite bug en version 2007
    executesql ('update activite set act_qteuniteref=0 where act_typeactivite="REA" and (act_typearticle="FRA" or act_typearticle="MAR") and act_qteuniteref<>0');

    //giga suite suppression aff_contact des vues.. suppression dans liste des bases
    AglNettoieListes('WPARCELEM', '',nil);
    AglNettoieListes('WPARCELEMM', '',nil);
    AglNettoieListes('WPARCELEMSELECT', '',nil);
    AglNettoieListes('WPARCELEMTIERS', '',nil);
    AglNettoieListes('WHISTOVERSIONS', '',nil);
    //giga init nouveau champs
    ExecuteSQLContOnExcept('update afplanning set apl_doublon="-"');
    ExecuteSQLContOnExcept('update tacheressource set atr_principale="-"');
    // Maj valeurs tablettes
    if ExisteSQL('select CO_CODE from commun where co_type="GHP"') then
    begin
    	// suite plantage dans le cas ou le V existe deja
    	if ExisteSQL('select CO_CODE from commun where co_type="GHP" AND CO_CODE="V"') then
      begin
      	ExecuteSQLContOnExcept('DELETE FROM COMMUN WHERE CO_TYPE="GHP" AND CO_CODE="V"');
      end;
      ExecuteSQLContOnExcept('UPDATE COMMUN set CO_CODE="V" WHERE CO_TYPE="GHP" AND CO_CODE="E"');
    end;
    // initialisation nouvelles zones
    ExecuteSQLContOnExcept('Update CIRCUITDA SET DAV_TYPERESPONS="V" Where DAV_TYPERESPONS="E"');

    {
      Laurent Abélard le 27 Septembre 2007, Il apparaît que plusieurs UPDATE sur LIGNECOMPL sont très long
      et nécessitent de regrouper tout ces UPDATE en un seul (Voir en Fin de MajVer822()).
    ExecuteSQLContOnExcept('Update LIGNECOMPL SET GLC_TYPEDA="" Where GLC_TYPEDA is Null');
    ExecuteSQLContOnExcept('Update LIGNECOMPL SET GLC_CODESERVICE="" Where GLC_CODESERVICE is Null');
    }

    ExecuteSQLContOnExcept('Update HISTODA SET DAH_CTRLBUDGET="-" Where DAH_CTRLBUDGET is Null');
    ExecuteSQLContOnExcept('Update LIGNEDA SET DAL_NUMORDRE=0 Where DAL_NUMORDRE is Null');
    ExecuteSQLContOnExcept('Update LIGNEDA SET DAL_WBMEMO="-" Where DAL_BLOCNOTE is Null');
    ExecuteSQLContOnExcept('Update LIGNEDA SET DAL_WBMEMO="X" Where DAL_BLOCNOTE is not Null');
    ExecuteSQLContOnExcept('Update PIECEDA SET DA_CODESESSION="" Where DA_CODESESSION is Null');
    ExecuteSQLContOnExcept('Update PIECEDA SET DA_CIRCUITCTRL="" Where DA_CIRCUITCTRL is Null');
    ExecuteSQLContOnExcept('Update PIECEDA SET DA_CTRLBUDGET="-" Where DA_CTRLBUDGET is Null');
    ExecuteSQLContOnExcept('Update PIECEDA SET DA_WBMEMO="-" Where DA_BLOCNOTE is Null');
    ExecuteSQLContOnExcept('Update PIECEDA SET DA_WBMEMO="X" Where DA_BLOCNOTE is not Null');
    ExecuteSQLContOnExcept('Update TYPEDA SET DAT_ETAPECIRCUIT="" Where DAT_ETAPECIRCUIT is Null');

  end;
  // JG
  if not ExisteSQL('SELECT 1 FROM CHOIXCOD WHERE CC_TYPE="GT1" AND CC_CODE="MES"') then
    ExecuteSQLContOnExcept('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) VALUES ("GT1","MES","Import MES","Import MES","") ');

  //Laurent Abélard le 27 Septembre 2007, Regroupement des UPDATE LIGNECOMPL entre le MajVer822() et MajVer844().
  //On part du principe que tous ces champs sont nouveaux et que la condition GLC_TYPEDA IS NULL vaut pour
  //tout les champs.
  If Not IsDossierPCL Then
    ExecuteSQLContOnExcept( 'UPDATE LIGNECOMPL SET GLC_TYPEDA="", GLC_CODESERVICE="", GLC_SECTIONPDR=""' +
                            ', GLC_RUBRIQUEPDR="", GLC_IMMOBILISATION="", GLC_CONSODIRECTE="-", GLC_COEFFREVALO=0' +
                            ', GLC_DATEDEBUTFAC="'+USDATETime(iDate1900)+'", GLC_DATEFINFAC="'+USDATETime(iDate2099)+'"' +
                            ', GLC_DATEAUGMENT="'+USDATETime(iDate1900)+'", GLC_DATEDEM="'+USDATETime(iDate1900)+'"' +
                            ', GLC_MONTANTTX1=0, GLC_MONTANTTX3=0, GLC_MONTANTTX4=0, GLC_MONTANTTX5=0 ' +
                            'WHERE GLC_TYPEDA IS NULL' );
end;

Procedure MajVer823 ;
Begin
  if not IsDossierPCL then
  begin
    // MCD
    executesql ('update afplanning set apl_competence=""');
  end;
end;

Procedure MajVer824 ;
Begin
  // MD
  // Transfert de l'indicateur de réalisation des obligations
  If IsMonoOuCommune then
    begin
    ExecuteSQLContOnExcept('update dpobligationrealise set dpo_etat=1, dpo_commentaire=' + DB_MID('dpo_commentaire', 4, 251) + ' where ' + DB_MID('dpo_commentaire',1,3) + '="[1]"');
    ExecuteSQLContOnExcept('update dpobligationrealise set dpo_etat=2, dpo_commentaire=' + DB_MID('dpo_commentaire', 4, 251) + ' where ' + DB_MID('dpo_commentaire',1,3) + '="[2]"');
    ExecuteSQLContOnExcept('update dpobligationrealise set dpo_etat=3, dpo_commentaire=' + DB_MID('dpo_commentaire', 4, 251) + ' where ' + DB_MID('dpo_commentaire',1,3) + '="[3]"');
    ExecuteSQLContOnExcept('update dpobligationrealise set dpo_etat=0 where dpo_etat is null');
    end;
  // CA
  AglNettoieListes('TRECRITURERAPPRO', 'TE_QUALIFORIGINE;');
  // MCD
  ExecuteSQLContOnExcept('update tierscompl set ytc_duns=""');
  // VG (vermot gauchy)
  ExecuteSQLContOnExcept('update immo set i_futurvnfisc=i_journala, i_reprisedr=i_reprisefiscal-i_repriseeco, i_reprisefdrcedee=i_repcedfisc-i_repcedeco,i_typederoglia=""');
  ExecuteSQLContOnExcept('update immo set i_typederoglia="DEG" where i_methodefisc="DEG"');
  ExecuteSQLContOnExcept('update immo set i_typederoglia="DUR" where i_methodefisc<>"" and i_typederoglia="" and i_dureeeco>i_dureefisc');
  ExecuteSQLContOnExcept('update immo set i_typederoglia="EXC" where i_methodefisc<>"" and i_typederoglia=""');


end;


Procedure MajVer825 ;
Begin
  if not IsDossierPCL then
  begin

    // JLS + modif ALR pour longueur 150 LFXXX
    ExecuteSQLContOnExcept('UPDATE LIGNEFRAIS SET LF_ORIFORF=SUBSTRING(LF_ORIFORF,1,69)||"....-.....|"'
    +'||iif((SUBSTRING(LF_ORIFORF,71,69)<>""),SUBSTRING(LF_ORIFORF,71,69)||"....-.....|"||SUBSTRING(LF_ORIFORF,141,9),"")'
    +' WHERE LF_ORIFORF<>"" AND LF_ORIFORF<>"m" AND ISNULL(SUBSTRING(LF_ORIFORF,70,11)," ")<>"....-.....|"');

    ExecuteSQLContOnExcept('UPDATE LIGNEFRAIS SET LF_ORIFIXE=SUBSTRING(LF_ORIFIXE,1,69)||"....-.....|"'
    +'||iif((SUBSTRING(LF_ORIFIXE,71,69)<>""),SUBSTRING(LF_ORIFIXE,71,69)||"....-.....|"||SUBSTRING(LF_ORIFIXE,141,9),"")'
    +' WHERE LF_ORIFIXE<>"" AND LF_ORIFIXE<>"m" AND ISNULL(SUBSTRING(LF_ORIFIXE,70,11)," ")<>"....-.....|"');

    ExecuteSQLContOnExcept('UPDATE LIGNEFRAIS SET LF_ORIREMMT=SUBSTRING(LF_ORIREMMT,1,69)||"....-.....|"'
    +'||iif((SUBSTRING(LF_ORIREMMT,71,69)<>""),SUBSTRING(LF_ORIREMMT,71,69)||"....-.....|"||SUBSTRING(LF_ORIREMMT,141,9),"")'
    +' WHERE LF_ORIREMMT<>"" AND LF_ORIREMMT<>"m" AND ISNULL(SUBSTRING(LF_ORIREMMT,70,11)," ")<>"....-.....|"');

    ExecuteSQLContOnExcept('UPDATE LIGNEFRAIS SET LF_ORICOUTBRUT=SUBSTRING(LF_ORICOUTBRUT,1,69)||"....-.....|"'
    +'||iif((SUBSTRING(LF_ORICOUTBRUT,71,69)<>""),SUBSTRING(LF_ORICOUTBRUT,71,69)||"....-.....|"||SUBSTRING(LF_ORICOUTBRUT,141,9),"")'
    +' WHERE LF_ORICOUTBRUT<>"" AND LF_ORICOUTBRUT<>"m" AND ISNULL(SUBSTRING(LF_ORICOUTBRUT,70,11)," ")<>"....-.....|"');

    ExecuteSQLContOnExcept('UPDATE LIGNEFRAIS SET LF_ORIPOURCENT=SUBSTRING(LF_ORIPOURCENT,1,69)||"....-.....|"'
    +'||iif((SUBSTRING(LF_ORIPOURCENT,71,69)<>""),SUBSTRING(LF_ORIPOURCENT,71,69)||"....-.....|"||SUBSTRING(LF_ORIPOURCENT,141,9),"")'
    +' WHERE LF_ORIPOURCENT<>"" AND LF_ORIPOURCENT<>"m" AND ISNULL(SUBSTRING(LF_ORIPOURCENT,70,11)," ")<>"....-.....|"');

    ExecuteSQLContOnExcept('UPDATE LIGNEFRAIS SET LF_ORICOUTNET=SUBSTRING(LF_ORICOUTNET,1,69)||"....-.....|"'
    +'||iif((SUBSTRING(LF_ORICOUTNET,71,69)<>""),SUBSTRING(LF_ORICOUTNET,71,69)||"....-.....|"||SUBSTRING(LF_ORICOUTNET,141,9),"")'
    +' WHERE LF_ORICOUTNET<>"" AND LF_ORICOUTNET<>"m" AND ISNULL(SUBSTRING(LF_ORICOUTNET,70,11)," ")<>"....-.....|"');

    // JLS
    ExecuteSQLContOnExcept('UPDATE YTARIFSPARAMETRES SET YFO_ENTETEINLIGNE="X", YFO_OKTIERSCOMPL="----", YFO_OKDIMENSIONS="----" WHERE YFO_ENTETEINLIGNE IS NULL');

    ExecuteSQLContOnExcept('UPDATE YTARIFS SET YTS_POIDSRECHFORT=0'
    +', YTS_ORIGINETIERS="", YTS_SECTEUR="", YTS_CODEENSEIGNE="", YTS_SOCIETEGROUPE=""'
    +', YTS_CASCTIERSORIG="-", YTS_CASCTIERSACTI="-", YTS_CASCTIERSENSE="-", YTS_CASCTIERSGROU="-"'
    +', YTS_GRILLEDIM1="", YTS_GRILLEDIM2="", YTS_GRILLEDIM3="", YTS_GRILLEDIM4="", YTS_GRILLEDIM5=""'
    +', YTS_CODEDIM1="", YTS_CODEDIM2="", YTS_CODEDIM3="", YTS_CODEDIM4="", YTS_CODEDIM5=""'
    +', YTS_CASCDIM1="-", YTS_CASCDIM2="-", YTS_CASCDIM3="-", YTS_CASCDIM4="-", YTS_CASCDIM5="-"'
    +' WHERE YTS_POIDSRECHFORT IS NULL');

    ExecuteSQLContOnExcept('UPDATE WPARAM SET WPA_VARCHAR16="" WHERE WPA_VARCHAR16 IS NULL');
{$IFNDEF BTP}
    if (V_PGI.Driver <> dbORACLE7 ) then
      MajPoidsTarifs;
{$ENDIF}

    // JLS
    ExecuteSQLContOnExcept('UPDATE WNOMEDEC SET WND_PHASE="" WHERE WND_PHASE IS NULL');

  end;

  // VG
  ExecuteSQLContOnExcept('UPDATE VISITEMEDTRAV SET PVM_LIBRECMB2="" WHERE (PVM_LIBRECMB2 IS NULL)');
  ExecuteSQLContOnExcept('UPDATE VISITEMEDTRAV SET PVM_LIBREDBL1=0 WHERE (PVM_LIBREDBL1 IS NULL)');
  ExecuteSQLContOnExcept('UPDATE VISITEMEDTRAV SET PVM_LIBREDBL2=0 WHERE (PVM_LIBREDBL2 IS NULL)');
  ExecuteSQLContOnExcept('UPDATE VISITEMEDTRAV SET PVM_LIBREDBL3=0 WHERE (PVM_LIBREDBL3 IS NULL)');

  ExecuteSQL ('update PAIEENCOURS SET PPU_BULLCONTRAT="-"');
  ExecuteSQLContOnExcept('update PAIEENCOURS SET PPU_UNITEPRISEFF= (SELECT PSA_UNITEPRISEFF FROM SALARIES WHERE PSA_SALARIE=PPU_SALARIE)');
  ExecuteSQLContOnExcept('update PAIEENCOURS SET PPU_BULLSOLDE = "-"');
  ExecuteSQLContOnExcept('update PAIEENCOURS SET PPU_BULLSOLDE = "X" WHERE PPU_DATESORTIE  > "'+UsDateTime(Idate1900)+'" AND PPU_DATESORTIE IS NOT NULL');
  ExecuteSQLContOnExcept('update REMUNERATION SET PRM_LIBCONTRAT="-"');
  ExecuteSQLContOnExcept('update EMETTEURSOCIAL SET PET_DUEIDENT = ""');
  ExecuteSQLContOnExcept('update EMETTEURSOCIAL SET PET_DUEADRAR = ""');
  //ExecuteSQLContOnExcept('UPDATE DUCSPARAM SET PDP_CODIFALSACE=PDP_CODIFICATION');
  //ExecuteSQLContOnExcept('UPDATE DUCSAFFECT SET PDF_CODIFALSACE=PDF_CODIFICATION');
  ExecuteSQLContOnExcept('UPDATE DUCSDETAIL SET PDD_REGIMEALSACE="-"');

  // VG
  if ExisteSQL ('SELECT "1" FROM PARAMSOC  WHERE SOC_NOM="SO_PGTYPECDETICKET" AND SOC_DATA <> "003"') then
  begin
    ExecuteSQL ('UPDATE PARAMSOC SET SOC_DATA = "-" '
        +' WHERE SOC_NOM IN ("SO_PGTKEDNPCARNET", "SO_PGTKEDRSCARNET","SO_PGTKEDCPVCARNET", "SO_PGTKDATELIVR")');

    ExecuteSQL ('UPDATE ETABCOMPL SET ETB_TICKLIVR=""')

  end
  else if (ExisteSQL  ('SELECT "1" FROM PARAMSOC  WHERE SOC_NOM="SO_PGTYPECDETICKET" AND SOC_DATA = "003"') and
        ExisteSQL  ('SELECT "1" FROM PARAMSOC  WHERE SOC_NOM="SO_PGMSA" AND SOC_DATA<>"X"'))
  then
  begin
    ExecuteSQL ('UPDATE ETABCOMPL SET ETB_TICKLIVR=ETB_AUTRENUMERO, ETB_AUTRENUMERO=""');
  end
  else if (ExisteSQL  ('SELECT "1" FROM PARAMSOC  WHERE SOC_NOM="SO_PGTYPECDETICKET" AND SOC_DATA = "003"')
    and ExisteSQL  ('SELECT "1" FROM PARAMSOC  WHERE SOC_NOM="SO_PGMSA" AND SOC_DATA="X"')) then
  begin
    ExecuteSQL ('UPDATE ETABCOMPL SET ETB_TICKLIVR=""');
  end;

  if ExisteSQL ('SELECT "1" FROM PARAMSOC  WHERE SOC_NOM="SO_PGTYPECDETICKET" AND SOC_DATA <> "003"') then
  begin
    ExecuteSQL ('UPDATE DEPORTSAL SET PSE_TYPTICKLIVR="ETB", PSE_TICKLIVR=""');
  end
  else if (ExisteSQL  ('SELECT "1" FROM PARAMSOC  WHERE SOC_NOM="SO_PGTYPECDETICKET" AND SOC_DATA = "003"')
    and ExisteSQL  ('SELECT "1" FROM PARAMSOC  WHERE SOC_NOM="SO_PGMSA" AND SOC_DATA<>"X"')) then
  begin
    ExecuteSQL ('UPDATE DEPORTSAL SET PSE_TICKLIVR= PSE_MSAUNITEGES,PSE_TYPTICKLIVR= PSE_MSATYPUNITEG,'
                +'PSE_MSAUNITEGES="",PSE_MSATYPUNITEG="ETB"');
  end
  else if (ExisteSQL  ('SELECT "1" FROM PARAMSOC  WHERE SOC_NOM="SO_PGTYPECDETICKET" AND SOC_DATA = "003"') and
    ExisteSQL  ('SELECT "1" FROM PARAMSOC  WHERE SOC_NOM="SO_PGMSA" AND SOC_DATA="X"')) then
  begin
    ExecuteSQL ('UPDATE DEPORTSAL SET PSE_TICKLIVR= "", PSE_TYPTICKLIVR= "ETB"');
  end;

end;


Procedure MajVer826 ;
var
iDat2007 : TDateTime;
savEnableDeshare : boolean;
Begin
  if not IsDossierPCL then
  begin
    ExecuteSQLContOnExcept(' update tache set ata_planningperiod = ata_periodicite');
    ExecuteSQLContOnExcept(' Update tache set ata_periodicite = ""');
    ExecuteSQLContOnExcept(' update afmodeletache set afm_planningperiod = afm_periodicite');
    ExecuteSQLContOnExcept(' Update afmodeletache set afm_periodicite = "" ');

    // CD
    ExecuteSQLContOnExcept('UPDATE SUSPECTS SET RSU_ENSEIGNE=""');

    ExecuteSQLContOnExcept('UPDATE SUSPECTSCOMPL SET RSC_REPRESENTANT2="",RSC_REPRESENTANT3="",RSC_TAUXREPR1=0.0,RSC_TAUXREPR2=0.0,RSC_TAUXREPR3=0.0');
    ExecuteSQLContOnExcept('UPDATE PARSUSPECTLIG SET RRL_FORMULE=""');

    // MCD
    if AfPlanCharge then
      executesql ('update paramsoc set soc_data="PDC" where soc_nom="so_afquelplanning"')
    else
      executesql ('update paramsoc set soc_data="DEU" where soc_nom="so_afquelplanning"');

    // JLS
    ExecuteSQLContOnExcept('UPDATE AFFAIRE SET AFF_RESPONSABLETEC="",AFF_DATEETATAFF="'+UsDateTime(iDate1900)+'",AFF_DATELIVPREVUE="'+UsDateTime(iDate1900)+'",AFF_CONTRATSAV="-",'
    +'AFF_POURCENTMARGE=0,AFF_POURCENTALEAS=0,AFF_TOTALCOMMANDE=0,AFF_TOTALFACTURE=0,AFF_OKMAJBUDGET="X",AFF_DCALCULBUDGET="'+UsDateTime(iDate1900)+'",'
    +'AFF_TYPEAFFAIRE="NOR",AFF_RECODIFOK="-" WHERE AFF_RESPONSABLETEC IS NULL');
    ExecuteSQLContOnExcept('UPDATE WPDRSECTION SET WRS_ORDREAFFSECT="0" WHERE WRS_ORDREAFFSECT IS NULL');

    ExecuteSQLContOnExcept('UPDATE WPDRRUBRIQUE SET WRR_FAMRUBRIQUE="", WRR_TYPESAISIE="", WRR_TAUXUNITAIRE=0,WRR_ORDREAFFRUBR=0,WRR_CUMULAUTORISE="-" WHERE WRR_TYPESAISIE IS NULL');

    {
      Laurent Abélard le 27 Septembre 2007, Il apparaît que plusieurs UPDATE sur LIGNECOMPL sont très long
      et nécessitent de regrouper tout ces UPDATE en un seul (Voir en Fin de MajVer822()).
    ExecuteSQLContOnExcept('UPDATE LIGNECOMPL SET GLC_SECTIONPDR="", GLC_RUBRIQUEPDR="" WHERE GLC_SECTIONPDR IS NULL');
    }
    ExecuteSQLContOnExcept('UPDATE WCBNEVOLUTION SET WEV_SECTIONPDR="", WEV_RUBRIQUEPDR="" WHERE WEV_SECTIONPDR IS NULL');
    ExecuteSQLContOnExcept('UPDATE QWRELEVELIG SET QWL_AFFAIRE="",QWL_SECTIONPDR="",QWL_RUBRIQUEPDR="" WHERE QWL_AFFAIRE IS NULL');
    ExecuteSQLContOnExcept('UPDATE QWHISTORES SET QWH_AFFAIRE="",QWH_SECTIONPDR="",QWH_RUBRIQUEPDR="",QWH_TAUXUNITAIRE=0 WHERE QWH_AFFAIRE IS NULL');
    ExecuteSQLContOnExcept('UPDATE QOALEA SET QOA_GESTAFFAIRE="-" WHERE QOA_GESTAFFAIRE IS NULL');
    ExecuteSQL ('UPDATE WPDRTYPE SET WRT_PARTYPEPARTA="-",WRT_PARTYPECOUTA="-",WRT_PARTYPEENTITEA="-",WRT_PARENTITEA="-",WRT_PARSECTIONA="-",WRT_PARRUBRIQUEA="-",WRT_PARORIGINEA="-",WRT_PAROPECIRCA="-",WRT_DEFAUTA="X" WHERE WRT_PARTYPEPARTA IS NULL');
    ExecuteSQL ('UPDATE WPDRTET SET WPE_AFFAIRE="" WHERE WPE_AFFAIRE IS NULL');
  end;

  // MNG
  ExecuteSQLContOnExcept('UPDATE YALERTES SET YAL_VERSIONALERTE="08.001.001",YAL_MEMO="-",YAL_EMAIL="-"');
  ExecuteSQLContOnExcept('UPDATE YALERTES SET YAL_EMAIL ="X" WHERE YAL_NOTIFICATION = "E"');
  ExecuteSQLContOnExcept('UPDATE YALERTES SET YAL_MEMO ="X" WHERE YAL_NOTIFICATION = "M"');
  ExecuteSQLContOnExcept('UPDATE YALERTES SET YAL_NOTIFICATION = ""');

  //Laurent Abélard le 25/04/2007
  //Cette requête génére plusieurs erreurs sur DB2 :
  //SQL0404 - La valeur destinée à la colonne ou la variable YTA_C00002 est trop longue.
  //SQL0803 - La valeur indiquée est incorrecte car elle produirait un clé en double.
  //ExecuteSQLContOnExcept('INSERT INTO YTABLEALERTES (YTA_ACTIF,YTA_CONTEXTES,YTA_CREATEUR,YTA_DATECREATION,YTA_DATEMODIF,YTA_LIBELLE,'+
  //  'YTA_NODOSSIER,YTA_PREDEFINI,YTA_PREFIXE,YTA_UTILISATEUR) SELECT "X","<<Tous>>","","'+UsDateTime(iDate1900)+'","'+UsDateTime(iDate1900)+'"'+
  //  ',CO_LIBELLE,"","",SUBSTRING(CO_CODE,1,LEN(CO_CODE)),"" FROM COMMUN WHERE CO_TYPE="YAL"');

  //Après avoir remis la liste des champs dans l'ordre du CREATE TABLE et supprimés le SUBSTRING :
  //IBM est visiblement chatouilleux sur les INSERT INTO SELECT !!!

  {
  savEnableDeshare:=V_PGI.EnableDeshare;
  V_PGI.EnableDeshare:=True;
  ExecuteSQLContOnExcept( 'INSERT INTO YTABLEALERTES (YTA_ACTIF,YTA_CONTEXTES,YTA_DATECREATION,YTA_DATEMODIF'
                        + ',YTA_CREATEUR,YTA_UTILISATEUR,YTA_PREDEFINI,YTA_NODOSSIER,YTA_PREFIXE,YTA_LIBELLE)'
                        + ' VALUES ("X","<<Tous>>","'+UsDateTime(iDate1900)+'","'+UsDateTime(iDate1900)+'",""'
                        + ',"","","","AFF","Affaire")');

  ExecuteSQLContOnExcept( 'INSERT INTO YTABLEALERTES (YTA_ACTIF,YTA_CONTEXTES,YTA_DATECREATION,YTA_DATEMODIF'
                        + ',YTA_CREATEUR,YTA_UTILISATEUR,YTA_PREDEFINI,YTA_NODOSSIER,YTA_PREFIXE,YTA_LIBELLE)'
                        + ' VALUES ("X","<<Tous>>","'+UsDateTime(iDate1900)+'","'+UsDateTime(iDate1900)+'",""'
                        + ',"","","","C","Contact")');

  ExecuteSQLContOnExcept( 'INSERT INTO YTABLEALERTES (YTA_ACTIF,YTA_CONTEXTES,YTA_DATECREATION,YTA_DATEMODIF'
                        + ',YTA_CREATEUR,YTA_UTILISATEUR,YTA_PREDEFINI,YTA_NODOSSIER,YTA_PREFIXE,YTA_LIBELLE)'
                        + ' VALUES ("X","<<Tous>>","'+UsDateTime(iDate1900)+'","'+UsDateTime(iDate1900)+'",""'
                        + ',"","","","GA","Article")');

  ExecuteSQLContOnExcept( 'INSERT INTO YTABLEALERTES (YTA_ACTIF,YTA_CONTEXTES,YTA_DATECREATION,YTA_DATEMODIF'
                        + ',YTA_CREATEUR,YTA_UTILISATEUR,YTA_PREDEFINI,YTA_NODOSSIER,YTA_PREFIXE,YTA_LIBELLE)'
                        + ' VALUES ("X","<<Tous>>","'+UsDateTime(iDate1900)+'","'+UsDateTime(iDate1900)+'",""'
                        + ',"","","","GP","Pièce")');

  ExecuteSQLContOnExcept( 'INSERT INTO YTABLEALERTES (YTA_ACTIF,YTA_CONTEXTES,YTA_DATECREATION,YTA_DATEMODIF'
                        + ',YTA_CREATEUR,YTA_UTILISATEUR,YTA_PREDEFINI,YTA_NODOSSIER,YTA_PREFIXE,YTA_LIBELLE)'
                        + ' VALUES ("X","<<Tous>>","'+UsDateTime(iDate1900)+'","'+UsDateTime(iDate1900)+'",""'
                        + ',"","","","T","Tiers")');
  V_PGI.EnableDeshare:=savEnableDeshare;
  }

  idat2007 := EncodeDate(2007, 1, 1);
  if not (ExisteSQL  ('SELECT "1" FROM YTABLEALERTELIEES  WHERE YTB_PREFIXEYTA="C" AND YTB_PREFIXEYTB = "RD6"')) then
  	ExecuteSQLContOnExcept('INSERT INTO YTABLEALERTELIEES (YTB_PREFIXEYTA,YTB_PREFIXEYTB,YTB_DATECREATION,YTB_DATEMODIF,YTB_CREATEUR,YTB_UTILISATEUR) VALUES ("C","RD6","'+UsDateTime(idat2007)+'","'+UsDateTime(idat2007)+'","CEG","CEG")');

  if not (ExisteSQL  ('SELECT "1" FROM YTABLEALERTELIEES  WHERE YTB_PREFIXEYTA="GA" AND YTB_PREFIXEYTB = "GA2"')) then
  	ExecuteSQLContOnExcept('INSERT INTO YTABLEALERTELIEES (YTB_PREFIXEYTA,YTB_PREFIXEYTB,YTB_DATECREATION,YTB_DATEMODIF,YTB_CREATEUR,YTB_UTILISATEUR) VALUES ("GA","GA2","'+UsDateTime(idat2007)+'","'+UsDateTime(idat2007)+'","CEG","CEG")');

  if not (ExisteSQL  ('SELECT "1" FROM YTABLEALERTELIEES  WHERE YTB_PREFIXEYTA="GA" AND YTB_PREFIXEYTB = "RD4"')) then
  	ExecuteSQLContOnExcept('INSERT INTO YTABLEALERTELIEES (YTB_PREFIXEYTA,YTB_PREFIXEYTB,YTB_DATECREATION,YTB_DATEMODIF,YTB_CREATEUR,YTB_UTILISATEUR) VALUES ("GA","RD4","'+UsDateTime(idat2007)+'","'+UsDateTime(idat2007)+'","CEG","CEG")');

  if not (ExisteSQL  ('SELECT "1" FROM YTABLEALERTELIEES  WHERE YTB_PREFIXEYTA="GP" AND YTB_PREFIXEYTB = "RDD"')) then
  	ExecuteSQLContOnExcept('INSERT INTO YTABLEALERTELIEES (YTB_PREFIXEYTA,YTB_PREFIXEYTB,YTB_DATECREATION,YTB_DATEMODIF,YTB_CREATEUR,YTB_UTILISATEUR) VALUES ("GP","RDD","'+UsDateTime(idat2007)+'","'+UsDateTime(idat2007)+'","CEG","CEG")');

  if not (ExisteSQL  ('SELECT "1" FROM YTABLEALERTELIEES  WHERE YTB_PREFIXEYTA="T" AND YTB_PREFIXEYTB = "YTC"')) then
  	ExecuteSQLContOnExcept('INSERT INTO YTABLEALERTELIEES (YTB_PREFIXEYTA,YTB_PREFIXEYTB,YTB_DATECREATION,YTB_DATEMODIF,YTB_CREATEUR,YTB_UTILISATEUR) VALUES ("T","YTC","'+UsDateTime(idat2007)+'","'+UsDateTime(idat2007)+'","CEG","CEG")');

  if not (ExisteSQL  ('SELECT "1" FROM YTABLEALERTELIEES  WHERE YTB_PREFIXEYTA="T" AND YTB_PREFIXEYTB = "RPR"')) then
  	ExecuteSQLContOnExcept('INSERT INTO YTABLEALERTELIEES (YTB_PREFIXEYTA,YTB_PREFIXEYTB,YTB_DATECREATION,YTB_DATEMODIF,YTB_CREATEUR,YTB_UTILISATEUR) VALUES ("T","RPR","'+UsDateTime(idat2007)+'","'+UsDateTime(idat2007)+'","CEG","CEG")');


  ExecuteSQLContOnExcept('UPDATE YTABLEALERTES SET YTA_ACTIF="'+BoolToStr_(GetParamSocSecur('SO_GEREALERTET',false))+'" WHERE YTA_PREFIXE="T"');
  ExecuteSQLContOnExcept('UPDATE YTABLEALERTES SET YTA_ACTIF="'+BoolToStr_(GetParamSocSecur('SO_GEREALERTEC',false))+'" WHERE YTA_PREFIXE="C"');
  ExecuteSQLContOnExcept('UPDATE YTABLEALERTES SET YTA_ACTIF="'+BoolToStr_(GetParamSocSecur('SO_GEREALERTEGA',false))+'" WHERE YTA_PREFIXE="GA"');
  ExecuteSQLContOnExcept('UPDATE YTABLEALERTES SET YTA_ACTIF="'+BoolToStr_(GetParamSocSecur('SO_GEREALERTEGP',false))+'" WHERE YTA_PREFIXE="GP"');
  ExecuteSQLContOnExcept('UPDATE YTABLEALERTES SET YTA_ACTIF="'+BoolToStr_(GetParamSocSecur('SO_GEREALERTEAFF',false))+'" WHERE YTA_PREFIXE="AFF" ');


end;



Procedure MajVer827 ;
var
 i : integer;
Begin

  if not IsDossierPCL then
  begin
    //ABU il faut inverser la case à cocher du paramètre societe suivant:
    setparamsoc('SO_AFALIMCUTOFF',not getparamsoc('SO_AFALIMCUTOFF'))
  end;
  // MNG
  for i := 1 to 3 do
    InsertChoixCode('ZLB', 'BB' + intTostr(i), 'Décision libre ' + intTostr(i), '', '');
  InsertChoixCode('ZLB', 'BM1', 'Multi-choix', '', '');
  // CD
  InsertChoixCode('DEN', 'GCO', 'Pièces', 'Pièces', '');
  InsertChoixCode('DEN', 'MLC', 'Maquettes GRC', 'GRC', '');
  InsertChoixCode('DEN', 'MLF', 'Maquettes GRF', 'GRF', '');

  //VG
//  ExecuteSQLContOnExcept('UPDATE ELTNATIONAUX SET PEL_CONVENTION= "000" WHERE PEL_CONVENTION IS NULL');
// TODO: voir si AddChamp conservé ou modif table ?

  // CA
  ExecuteSQLContOnExcept('Delete from formes where dfm_forme="CPRFISCAL2006"');
end;


Procedure MajVer828 ;
var
 i : integer;
Begin

  if not IsDossierPCL then
  begin
    // JLS
    ExecuteSQLContOnExcept('UPDATE STKNATURE SET GSN_QTEPLUS="GQ_PREAFFECTE;" || GSN_QTEPLUS WHERE GSN_STKTYPEMVT = "ATT" AND NOT(GSN_QTEPLUS LIKE "%GQ_PREAFFECTE%")');

    ExecuteSQLContOnExcept('update Dispo set gq_Preaffecte=0, Gq_PrepaFou=0');
    {
      Laurent Abélard le 27 Septembre 2007, Il apparaît que plusieurs UPDATE sur LIGNECOMPL sont très long
      et nécessitent de regrouper tout ces UPDATE en un seul (Voir en Fin de MajVer822()).
    ExecuteSQLContOnExcept('UPDATE LIGNECOMPL SET GLC_IMMOBILISATION = ""');
    ExecuteSQLContOnExcept('UPDATE LIGNECOMPL SET GLC_CONSODIRECTE="-" WHERE GLC_CONSODIRECTE IS NULL');
    }

    // JLS
    ExecuteSQLContOnExcept('UPDATE TIERSFRAIS SET GTF_FRANCO=0 WHERE GTF_FRANCO IS NULL');
    AGLNettoieListesPlus('TIERSFRAIS','GTF_FRANCO',nil,True);
    ExecuteSQLContOnExcept('UPDATE wartnat SET WAN_ACTIF="X" WHERE WAN_ACTIF IS NULL');
    AGLNettoieListesPlus('AFMULRECHAFFAIRE','AFF_TYPEAFFAIRE',nil,False);
    ExecuteSQLContOnExcept('delete from paramsoc where soc_nom in ("SCO_PDRDEFAUTGA", "SCO_PDRCONSOAVANCE", "SCO_PDRGENEREWPL", "SCO_PDRMETHVALO", "SCO_PDRNBORDREGA2")');
  end
  else
  //IsDossierPCL
  {Laurent ABELARD le 29/10/2007 Correction suite à la remarque de Jean PASTERIS sur l'erreur de la
  l'interprétation de la Demande N° 682 où le BEGIN END sont de trop.
  begin}
    // CA
    ExecuteSQLContOnExcept('UPDATE BANQUECP SET BQ_NODOSSIER = "' + V_PGI.NoDossier + '" WHERE BQ_NODOSSIER = "000000" ' +
                           'OR BQ_NODOSSIER = "" OR BQ_NODOSSIER IS NULL');

  if not GetParamSocSecur('SO_MODETRESO', False) and (V_PGI.ModePCL <> '1') then
    ExecuteSQLContOnExcept('UPDATE RUBRIQUE SET RB_NODOSSIER = "' + GetParamSocSecur('SO_NODOSSIER', '') + '" WHERE RB_NODOSSIER = "000000" ' +
                           'OR RB_NODOSSIER = "" OR RB_NODOSSIER IS NULL');

  if V_PGI.ModePCL <> '1' then ExecuteSQLContOnExcept('UPDATE REGLEACCRO SET TRG_PREDEFINI = "STD"');

  if V_PGI.ModePCL <> '1' then
  begin
    if not existeSql('SELECT 1 FROM CIB WHERE TCI_PREDEFINI="STD"') then
    begin
  		ExecuteSQLContOnExcept('UPDATE CIB SET TCI_PREDEFINI = "STD" WHERE TCI_PREDEFINI<>"CEG"');
    end;
  end;

  AglNettoieListes('TRLISTECIBACCRO', 'TCI_PREDEFINI;');
  AglNettoieListes('TRLISTECIBREF', 'TCI_PREDEFINI;');
  //end;

  // LMO
  ExecuteSQLContOnExcept('update ANNUAIRE set ANN_NOM4 ="", ANN_PAYSNAIS ="", ANN_CODENAF2=""');
  ExecuteSQLContOnExcept('update ANNUAIRE set 	ANN_TYPESCI = ( select DOR_TYPESCI from DPORGA where DOR_GUIDPER = ANN_GUIDPER)');
  //update ANNUAIRE set 	ANN_TYPESCI = DOR_TYPESCI '
  //+' from ANNUAIRE A left join DPORGA D on A.ANN_GUIDPER = D.DOR_GUIDPER');

//  ExecuteSQLContOnExcept('update JURIDIQUE set 	JUR_LOCAGERANCE = DOR_LOCAGERANCE ,	JUR_DATEDEBUTEX = DOR_DUREE ,'
//  +'	JUR_DATEFINEX = DOR_DATEFINEX,	JUR_DUREEEX = DOR_DUREE,	JUR_DUREEEXPREC = DOR_DUREEPREC '
//  +' from JURIDIQUE left join DPORGA on JUR_GUIDPERDOS = DOR_GUIDPER');
  ExecuteSQLContOnExcept('update JURIDIQUE set 	JUR_LOCAGERANCE = (select DOR_LOCAGERANCE FROM DPORGA where JUR_GUIDPERDOS = DOR_GUIDPER) ,'
  + ' JUR_DATEDEBUTEX = (select DOR_DATEDEBUTEX FROM DPORGA where JUR_GUIDPERDOS = DOR_GUIDPER) ,'
  +'	JUR_DATEFINEX = (select DOR_DATEFINEX FROM DPORGA where JUR_GUIDPERDOS = DOR_GUIDPER),'
  +' 	JUR_DUREEEX = (select DOR_DUREE FROM DPORGA where JUR_GUIDPERDOS = DOR_GUIDPER),	'
  +'  JUR_DUREEEXPREC = (select DOR_DUREEPREC FROM DPORGA where JUR_GUIDPERDOS = DOR_GUIDPER)');

  // LMO
  ExecuteSQLContOnExcept('update DPFISCAL set '
  +' 	DFI_TSDOMACT = "",'
  +'  DFI_REGLEFISC = ( select DOR_REGLEFISC  from DPORGA where DFI_GUIDPER = DOR_GUIDPER),'
  +'  DFI_DATECONVEDITVA = "'+ UsDateTime(iDate1900) + '",'
  +' 	DFI_TSABATASSO = "-",'
  +' 	DFI_TSACTDIFF = "-",'
  +' 	DFI_TSPERIODICITE = "-",'
  +'  DFI_TAXESALAIRES = (select DSO_TAXESALARIE from DPSOCIAL where  DFI_GUIDPER = DSO_GUIDPER)');

  ExecuteSQLContOnExcept('update DPSOCIAL'
  +' set	DSO_RETCOLLECTIF39 = "-",'
  +'	DSO_IFC = "-",'
  +'	DSO_STATUT = "",'
  +'	DSO_CONJOINTAVEC = "-",'
  +'	DSO_CONTRATPREVOY="-",'
  +'	DSO_CONTRATRET = "-",'
  +'	DSO_CONJOINTSTATUT  =""');

  ExecuteSQLContOnExcept('update ANNULIEN'
  +' set	ANL_STATUT = "",'
  +' 	ANL_CONJOINTAVEC="-",'
  +' 	ANL_CONJOINTSTATUT="",'
  +' 	ANL_CONTRATRET = "",'
  +' 	ANL_CONTRATPREVOY="",'
  +' 	ANL_QUALITESIGN="",'
  +' 	ANL_DECLARATION=""') ;

  ExecuteSQLContOnExcept('UPDATE ADRESSES SET ADR_CDE="-" WHERE ADR_CDE IS NULL');
  
end;



Procedure MajVer829 ;
Begin

  ExecuteSQLContOnExcept( 'UPDATE immo set i_futurvnfisc=i_journala,'   +
                          'i_reprisedr=i_reprisefiscal-i_repriseeco,'   +
                          'i_reprisefdrcedee=i_repcedfisc-i_repcedeco,' +
                          'i_typederoglia="" where i_methodeeco<>"VAR"');

  ExecuteSQLContOnExcept( 'UPDATE immo set i_futurvnfisc=i_journala,'  +
                          'i_reprisedr=i_reprisefiscal-(i_repriseeco + i_reprisedep),' +
                          'i_reprisefdrcedee=i_repcedfisc -(i_repcedeco+ i_reprisedepcedee),' +
                          'i_typederoglia="" where i_methodeeco="VAR"');

  ExecuteSQLContOnExcept('UPDATE immo set i_typederoglia="DEG" where i_methodefisc="DEG"');

  ExecuteSQLContOnExcept('UPDATE immo set i_typederoglia="DUR"where i_methodefisc<>"" and i_typederoglia="" and i_dureeeco>i_dureefisc');
  ExecuteSQLContOnExcept('UPDATE immo  set i_typederoglia="EXC" where i_methodefisc<>"" and i_typederoglia=""');

  //Martine VERMOT-GAUCHY Le 15/11/2007 Version 8.1.850.15 Demande N° 1995
  {ExecuteSQLContOnExcept('UPDATE immo SET i_datecession= isnull((select il_dateop FROM immolog il'
  + ' WHERE immo.i_immo=il.il_immo and il.il_typeop="CES"),"' + UsDateTime(iDate1900) + '")') ;}
  //J'abandonne la modification car on retrouve exactement la même requête en MajVer80
  {ExecuteSQLContOnExcept( 'UPDATE immo SET i_datecession = isnull((select MAX(il_dateop) FROM immolog il'
                        + ' WHERE immo.i_immo=il.il_immo and il.il_typeop="CES"),"' + UsDateTime(iDate1900) + '")');}

  //Martine VERMOT-GAUCHY Le 15/11/2007 Version 8.1.850.15 Demande N° 1995
  {ExecuteSQLContOnExcept('UPDATE immo SET i_dateleveeoption=( SELECT il_dateop'
  +' FROM immolog il WHERE immo.i_immo=il.il_immo and il.il_typeop="LEV" )');}
  //J'abandonne la modification car on retrouve exactement la même requête en MajVer80
  {ExecuteSQLContOnExcept( 'UPDATE immo SET i_dateleveeoption = isnull((select MAX(il_dateop) FROM immolog il'
                        + '  WHERE immo.i_immo=il.il_immo and il.il_typeop="LEV"),"' + UsDateTime(iDate1900) + '")' );}

end;

Procedure MajVer830 ;
var i : integer;
Begin
  // CA
  ExecuteSQLContOnExcept('update tierscompl set ytc_retenuesource=""');
  // CD
  //for i:=3 to 9 do InsertChoixCode ('RSZ', 'BL'+IntToStr(i), 'Décision libre ' + intTostr(i + 1), '','');
  //for i:=3 to 9 do InsertChoixCode ('RSZ', 'DL'+IntToStr(i), 'Date libre ' + intTostr(i + 1), '','');
  //for i:=3 to 9 do InsertChoixCode ('RSZ', 'TL'+IntToStr(i), 'Texte libre ' + intTostr(i + 1), '','');
  //for i:=3 to 9 do InsertChoixCode ('RSZ', 'VL'+IntToStr(i), 'Valeur libre ' + intTostr(i + 1), '','');

  //Marie-Christine DESSEIGNET Demande N° 1341
  for i:=3 to 9 do InsertChoixCode ('RSZ', 'BL'+IntToStr(i), '.-Suspect Décision libre ' + intTostr(i + 1), '','');
  for i:=3 to 9 do InsertChoixCode ('RSZ', 'DL'+IntToStr(i), '.-Suspect Date libre ' + intTostr(i + 1), '','');
  for i:=3 to 9 do InsertChoixCode ('RSZ', 'TL'+IntToStr(i), '.-Suspect Texte libre ' + intTostr(i + 1), '','');
  for i:=3 to 9 do InsertChoixCode ('RSZ', 'VL'+IntToStr(i), '.-Suspect Valeur libre t' + intTostr(i + 1), '','');

  // CA
  If IsMonoOuCommune then
  begin
    ExecuteSQLContOnExcept('UPDATE DPTABCOMPTA SET DTC_SUBVPRO=0, DTC_SUBVSTR=0, DTC_SUBVREV=0, DTC_SUBVAUT=0');
    // MD Initialisation des nouveaux champs de la table DOSSIER
    ExecuteSQLContOnExcept('UPDATE DOSSIER SET DOS_WINOUV="", DOS_WINSTR=""');
    // MD Glissement de données
    ExecuteSQLContOnExcept('update ANNUAIRE set ANN_NOM4 = (SELECT DPP_NOMUSAGE from DPPERSO where DPP_GUIDPER=ANN_GUIDPER)');
  end;

  if not IsDossierPCL then
  begin
    // CD
    ExecuteSql ('update paractions set rpa_planifiable = "-"') ;
    // DB
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_INITQTE = "001"');
    ExecuteSQLContOnExcept('UPDATE PIECE SET GP_ECHEBLOQUE = "-", GP_TYPETAUXFIXING = "TF", GP_TAUXFIXING = (SELECT D_PARITEEUROFIXING FROM DEVISE WHERE D_DEVISE=GP_DEVISE) WHERE GP_VIVANTE="X"');
  end;

  // VG
  ExecuteSQLContOnExcept('UPDATE immo set i_futurvnfisc=i_journala,'+ 'i_reprisedr=i_reprisefiscal-i_repriseeco,'
  + 'i_reprisefdrcedee=i_repcedfisc-i_repcedeco,'+ 'i_typederoglia=""'+ 'where i_methodeeco<>"VAR"');

  ExecuteSQLContOnExcept('UPDATE immo set i_futurvnfisc=i_journala,'+ 'i_reprisedr=i_reprisefiscal-(i_repriseeco + i_reprisedep),'
  + 'i_reprisefdrcedee=i_repcedfisc -(i_repcedeco+ i_reprisedepcedee),'
  + 'i_typederoglia="" where i_methodeeco="VAR"');

  ExecuteSQLContOnExcept('UPDATE immo set i_typederoglia="DEG" where i_methodefisc="DEG"');
  ExecuteSQLContOnExcept('UPDATE immo set i_typederoglia="DUR"where i_methodefisc<>"" and i_typederoglia="" and i_dureeeco>i_dureefisc');
  ExecuteSQLContOnExcept('UPDATE immo  set i_typederoglia="EXC" where i_methodefisc<>"" and i_typederoglia=""');

  //Martine VERMOT-GAUCHY Le 15/11/2007 Version 8.1.850.15 Demande N° 1995
  {ExecuteSQLContOnExcept('UPDATE immo SET i_datecession= isnull((select il_dateop FROM immolog il'
  +'  WHERE immo.i_immo=il.il_immo and il.il_typeop="CES"),"' + UsDateTime(iDate1900) + '")' ) ;}
  ExecuteSQLContOnExcept( 'UPDATE immo SET i_datecession = isnull((select MAX(il_dateop) FROM immolog il'
                        + ' WHERE immo.i_immo=il.il_immo and il.il_typeop="CES"),"' + UsDateTime(iDate1900) + '")');

  //Martine VERMOT-GAUCHY Le 15/11/2007 Version 8.1.850.15 Demande N° 1995
  {ExecuteSQLContOnExcept('UPDATE immo SET i_dateleveeoption= isnull((select il_dateop FROM immolog il'
  +'  WHERE immo.i_immo=il.il_immo and il.il_typeop="LEV"),"' + UsDateTime(iDate1900) + '")' ) ;}
  ExecuteSQLContOnExcept( 'UPDATE immo SET i_dateleveeoption = isnull((select MAX(il_dateop) FROM immolog il'
                        + '  WHERE immo.i_immo=il.il_immo and il.il_typeop="LEV"),"' + UsDateTime(iDate1900) + '")' );

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/03/2007
Modifié le ... :   /  /
Description .. : Pour PGIMAGVER
Mots clefs ... :
*****************************************************************}
//Christophe AYEL Demande N° 1424
procedure TransfertPlanRevision;
begin
  ExecuteSQLContOnExcept( 'INSERT INTO CREVPLAN ' +
                          '(CPR_PLANREVISION, CPR_LIBELLE, CPR_ABREGE, CPR_ACTIVATION, ' +
                          'CPR_PREDEFINI, CPR_DATECREATION, CPR_DATEMODIF, ' +
                          'CPR_CREATEUR, CPR_UTILISATEUR) ' +
                          '(SELECT YDS_CODE, SUBSTRING(YDS_LIBELLE, 1, 35), YDS_ABREGE, "", ' +
                          'YDS_PREDEFINI, "' + UsDateTime( Now ) + '","' + UsDateTime( Now ) + '", '+
                          '"CEG", "CEG" FROM CHOIXDPSTD WHERE YDS_TYPE = "PDR")');

  ExecuteSQLContOnExcept('DELETE FROM CHOIXDPSTD WHERE YDS_TYPE = "PDR"');
end;

////////////////////////////////////////////////////////////////////////////////

{***********A.G.L.***********************************************

Auteur  ...... : Gilles COSTE

Créé le ...... : 06/03/2007

Modifié le ... : 21/03/2007

Description .. : Ok fonctionne pour PGIMAJVER

Mots clefs ... :

*****************************************************************}

procedure TransfertCycleRevision;
var lQuery : TQuery;
    lStRubrique   : string;
    lStAbrege     : string;
    lStCompte1    : string;
    lStExclusion1 : string;

begin

  lQuery := nil;

  try

      lQuery := OpenSQL('SELECT RB_RUBRIQUE, RB_LIBELLE, RB_FAMILLES, ' +
                        'RB_COMPTE1, RB_EXCLUSION1, RB_PREDEFINI FROM RUBRIQUE WHERE ' +
                        'RB_CLASSERUB = "CDR" ORDER BY RB_RUBRIQUE', True);

      if lQuery.Eof then Exit;

      while not lQuery.Eof do
      begin

        lStRubrique   := lQuery.FindField('RB_RUBRIQUE').AsString;
        lStAbrege     := Copy(lStRubrique, 1, 6);
        lStCompte1    := lQuery.FindField('RB_COMPTE1').asString;
        lStExclusion1 := lQuery.FindField('RB_EXCLUSION1').asString;

        lStCompte1    := FindEtReplace(lStCompte1, '(SM)', '', True);
        lStCompte1    := FindEtReplace(lStCompte1, '(SD)', '', True);
        lStCompte1    := FindEtReplace(lStCompte1, '(SC)', '', True);
        lStExclusion1 := FindEtReplace(lStExclusion1, '(SM)', '', True);
        lStExclusion1 := FindEtReplace(lStExclusion1, '(SD)', '', True);
        lStExclusion1 := FindEtReplace(lStExclusion1, '(SC)', '', True);

        ExecuteSQLContOnExcept('INSERT INTO CREVPARAMCYCLE' +
        '(CPC_CODECYCLE, CPC_CODEABREGE, CPC_LIBELLECYCLE, CPC_PLANASSOCIE, ' +
        'CPC_LISTECOMPTE, CPC_LISTEEXCLUSION, CPC_LISTEACTIVE, CPC_ACTIVATION, ' +
        'CPC_ACTIVATIONSQL, CPC_DATECREATION, CPC_DATEMODIF, CPC_CREATEUR, ' +
        'CPC_UTILISATEUR, CPC_PREDEFINI) ' +
        ' VALUES ("' + lStRubrique + '",' +
         '"' + lStAbrege + '",' +
         '"' + lQuery.FindField('RB_LIBELLE').AsString + '",' +
         '"' + lQuery.FindField('RB_FAMILLES').AsString + '",' +
         '"' + lStCompte1 + '",' +
         '"' + lStExclusion1 + '", "", "", "",' + // LISTEACTIVE + ACTIVATION + ACTIVATIONSQL
         '"' + UsDateTime( Now ) + '","' + UsDateTime( Now ) + '",'+
         '"CEG", "CEG", ' +
         '"' + lQuery.FindField('RB_PREDEFINI').AsString + '")');
        lQuery.Next;
      end;
      ExecuteSQLContOnExcept('DELETE FROM RUBRIQUE WHERE RB_CLASSERUB = "CDR"');
  finally
    Ferme( lQuery );
  end;
end;


Procedure MajVer831 ;
Begin
  if not IsDossierPCL then
  begin
    // NR
    ExecuteSQLContOnExcept('Update PIECEADRESSE SET GPA_TYPECONTACT="" Where  GPA_TYPECONTACT is Null');
    ExecuteSQLContOnExcept('Update PIECEADRESSE SET GPA_AUXICONTACT="" Where GPA_AUXICONTACT is Null');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_TYPEPASSCPTAR = GPP_TYPEPASSCPTA, GPP_TYPEPASSACCR = GPP_TYPEPASSACC');
    // JLS
    ExecuteSQLContOnExcept('UPDATE QWRELEVELIG SET QWL_TAUXUNITAIRE=0 WHERE QWL_TAUXUNITAIRE IS NULL');
    // JTR
    ExecuteSQLContOnExcept('UPDATE PARPIECECOMPL SET GPC_TYPEPASSCPTAR = GPC_TYPEPASSCPTA, GPC_TYPEPASSACCR = GPC_TYPEPASSACC');
  end;

  // MD
  If IsMonoOuCommune = TRUE then
  begin
    ExecuteSQL ('UPDATE UTILISAT SET US_TEL1="",US_TEL2="",US_TEL3="" ')
  end;
  // CA
  ExecuteSQLContOnExcept('UPDATE SECTION SET S_TRANCHEGENEDE="", S_TRANCHEGENEA="", S_INDIRECTE="-", S_UO=0, S_UOLIBELLE=""');
  ExecuteSQLContOnExcept('UPDATE SECTIONREF SET SRE_TRANCHEGENEDE="", SRE_TRANCHEGENEA="", SRE_INDIRECTE="-", SRE_UO=0, SRE_UOLIBELLE=""');

  //  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE = "FEMPRUNT"');

  // VG
  ExecuteSQLContOnExcept('update DECLARATIONS SET PDT_ALLERETOUR = "-", PDT_LOCALITEACC = "", PDT_TEMOIN2 = "", PDT_TEMOIN2ADR1 = "", PDT_TEMOIN2ADR2 = "", PDT_TEMOIN3 = "", PDT_TEMOIN3ADR1 = "", PDT_TEMOIN3ADR2 = "", PDT_CIRCACC7 = ""');
  ExecuteSQLContOnExcept('update CONTRATTRAVAIL SET PCI_FONCTIONSAL=""');
  ExecuteSQLContOnExcept('update REMUNERATION SET PRM_PRIMEASSEDIC="-"');
  ExecuteSQLContOnExcept('update SALARIES SET PSA_FONCTIONSAL="", PSA_REGIMEAT="", PSA_REGIMEMAL="", PSA_REGIMEVIP="", PSA_REGIMEVIS="", PSA_TYPEREGIME="-"');
  ExecuteSQLContOnExcept('UPDATE SESSIONSTAGE SET PSS_NODOSSIER="000000", PSS_PREDEFINI="STD", PSS_PGTYPESESSION="AUC", PSS_IDSESSIONFOR=RTRIM(PSS_CODESTAGE)||LTRIM(STR(PSS_ORDRE)), PSS_NOSESSIONMULTI=-1, PSS_PROGRAMME="-", PSS_NBSTAMIN="0"');
  ExecuteSQLContOnExcept('UPDATE FORMATIONS SET PFO_NODOSSIER="000000", PFO_PREDEFINI="STD", PFO_PGTYPESESSION="AUC", PFO_IDSESSIONFOR=RTRIM(PFO_CODESTAGE)||LTRIM(STR(PFO_ORDRE)), PFO_NOSESSIONMULTI=-1, PFO_CURSUS=""');
  ExecuteSQLContOnExcept('UPDATE SESSIONANIMAT SET PAN_NODOSSIER="000000", PAN_PREDEFINI="STD", PAN_PGTYPESESSION="AUC", PAN_IDSESSIONFOR=RTRIM(PAN_CODESTAGE)||LTRIM(STR(PAN_ORDRE)), PAN_NOSESSIONMULTI=-1');
  ExecuteSQLContOnExcept('UPDATE FRAISSALFORM SET PFS_NODOSSIER="000000", PFS_PREDEFINI="STD", PFS_PGTYPESESSION="AUC", PFS_IDSESSIONFOR=RTRIM(PFS_CODESTAGE)||LTRIM(STR(PFS_ORDRE)), PFS_NOSESSIONMULTI=-1, PFS_CODEPOP="", PFS_POPULATION=""');
  ExecuteSQLContOnExcept('UPDATE STAGE SET PST_NODOSSIER="000000", PST_PREDEFINI="STD"');
  ExecuteSQLContOnExcept('UPDATE INTERIMAIRES SET PSI_NODOSSIER="000000", PSI_PREDEFINI="STD"');
  ExecuteSQLContOnExcept('UPDATE LIEUFORMATION SET PLF_VILLE="" WHERE PLF_VILLE IS NULL');
  ExecuteSQLContOnExcept('UPDATE INSCFORMATION SET PFI_NODOSSIER="000000", PFI_PREDEFINI="STD"');
  ExecuteSQLContOnExcept('UPDATE CURSUS SET PCU_NODOSSIER="000000", PCU_PREDEFINI="STD"');
  ExecuteSQLContOnExcept('UPDATE LIEUFORMATION SET PLF_ADRESSE1="", PLF_ADRESSE2="", PLF_ADRESSE3="", PLF_VILLE=(SELECT O_VILLE FROM CODEPOST WHERE O_CODEPOSTAL=PLF_CODEPOSTAL)');
  ExecuteSQLContOnExcept('UPDATE FRAISSALPLAF SET PFP_CODEPOP="", PFP_POPULATION=""');
  ExecuteSQLContOnExcept('UPDATE FORFAITFORM SET PFF_CODEPOP="---", PFF_POPULATION="---"');
  ExecuteSQLContOnExcept('update ETABCOMPL SET ETB_PERIODCT="-"');
  ExecuteSQLContOnExcept('UPDATE SALARIES SET PSA_PERIODCT="-", PSA_TYPPERIODCT="ETB" ');
  ExecuteSQLContOnExcept('UPDATE CRITMAINTIEN SET PCM_VALCATEGORIE2="", PCM_VALCATEGORIE3="", PCM_VALCATEGORIE4="", PCM_TYPEABSENCES=PCM_TYPEABS');

  if not IsDossierPCL then
  begin
      // PL
      ExecuteSQLContOnExcept('update CIBLAGE SET RCB_RCBTABLELIBRE1="", RCB_RCBTABLELIBRE2="", RCB_RCBTABLELIBRE3="",'
      +'RCB_RCBTABLELIBRE4="", RCB_RCBTABLELIBRE5="", RCB_RCBTEXTELIBRE1="", RCB_RCBTEXTELIBRE2="",'
      +'RCB_RCBTEXTELIBRE3="", RCB_RCBTEXTELIBRE4="", RCB_RCBTEXTELIBRE5="",'
      +'RCB_RCBDATELIBRE1="'+USDATETime(iDate1900)+'",'
      +'RCB_RCBDATELIBRE2="'+USDATETime(iDate1900)+'",'
      +'RCB_RCBDATELIBRE3="'+USDATETime(iDate1900)+'",'
      +'RCB_RCBDATELIBRE4="'+USDATETime(iDate1900)+'",'
      +'RCB_RCBDATELIBRE5="'+USDATETime(iDate1900)+'"');

      ExecuteSQLContOnExcept('update SUSPECTSCOMPL SET RSC_6RSCLIBTABLE26="", RSC_6RSCLIBTABLE27="",'
      +'RSC_6RSCLIBTABLE28="", RSC_6RSCLIBTABLE29="", RSC_6RSCLIBTABLE30="",'
      +'RSC_6RSCLIBTABLE31="", RSC_6RSCLIBTABLE32="", RSC_6RSCLIBTABLE33="",'
      +'RSC_6RSCLIBTABLE34="",'
      +'RSC_6RSCLIBTABLE35="", RSC_RSCLIBTEXTE10="", RSC_RSCLIBTEXTE11="",'
      +'RSC_RSCLIBTEXTE12="", RSC_RSCLIBTEXTE13="", RSC_RSCLIBTEXTE14="",'
      +'RSC_RSCLIBTEXTE15="", RSC_RSCLIBTEXTE16="", RSC_RSCLIBTEXTE17="",'
      +'RSC_RSCLIBTEXTE18="", RSC_RSCLIBTEXTE19="",'
      +'RSC_RSCLIBTEXTE20="", RSC_RSCLIBTEXTE21="", RSC_RSCLIBTEXTE22="",'
      +'RSC_RSCLIBTEXTE23="", RSC_RSCLIBTEXTE24="", RSC_RSCLIBTEXTE25="",'
      +'RSC_RSCLIBTEXTE26="", RSC_RSCLIBTEXTE27="", RSC_RSCLIBTEXTE28="",'
      +'RSC_RSCLIBTEXTE29="",'
    +'RSC_RSCLIBMUL10="", RSC_RSCLIBMUL11="", RSC_RSCLIBMUL12="",'
    +'RSC_RSCLIBMUL13="", RSC_RSCLIBMUL14="", RSC_RSCLIBMUL15="",'
    +'RSC_RSCLIBMUL16="", RSC_RSCLIBMUL17="", RSC_RSCLIBMUL18="",'
    +'RSC_RSCLIBMUL19="", RSC_RSCLIBBOOL10="-",'
    +'RSC_RSCLIBBOOL11="-", RSC_RSCLIBBOOL12="-", RSC_RSCLIBBOOL13="-",'
    +'RSC_RSCLIBBOOL14="-", RSC_RSCLIBBOOL15="-", RSC_RSCLIBBOOL16="-",'
    +'RSC_RSCLIBBOOL17="-", RSC_RSCLIBBOOL18="-", RSC_RSCLIBBOOL19="-",'
    +'RSC_RSCLIBVAL10=0, RSC_RSCLIBVAL11=0,'
    +'RSC_RSCLIBVAL12=0, RSC_RSCLIBVAL13=0, RSC_RSCLIBVAL14=0, '
    +'RSC_RSCLIBVAL15=0, RSC_RSCLIBVAL16=0, RSC_RSCLIBVAL17=0, '
    +'RSC_RSCLIBVAL18=0, RSC_RSCLIBVAL19=0, '
    +'RSC_RSCLIBDATE10="'+USDATETime(iDate1900)+'",'
    +'RSC_RSCLIBDATE11="'+USDATETime(iDate1900)+'",'
    +'RSC_RSCLIBDATE12="'+USDATETime(iDate1900)+'",'
    +'RSC_RSCLIBDATE13="'+USDATETime(iDate1900)+'",'
    +'RSC_RSCLIBDATE14="'+USDATETime(iDate1900)+'",'
    +'RSC_RSCLIBDATE15="'+USDATETime(iDate1900)+'",'
    +'RSC_RSCLIBDATE16="'+USDATETime(iDate1900)+'",'
    +'RSC_RSCLIBDATE17="'+USDATETime(iDate1900)+'",'
    +'RSC_RSCLIBDATE18="'+USDATETime(iDate1900)+'",'
    +'RSC_RSCLIBDATE19="'+USDATETime(iDate1900)+'"');

    ExecuteSQLContOnExcept('update PROSPECTS SET RPR_RPRLIBTEXTE10="",'
    +'RPR_RPRLIBTEXTE11="", RPR_RPRLIBTEXTE12="", RPR_RPRLIBTEXTE13="",'
    +'RPR_RPRLIBTEXTE14="", RPR_RPRLIBTEXTE15="", RPR_RPRLIBTEXTE16="",'
    +'RPR_RPRLIBTEXTE17="", RPR_RPRLIBTEXTE18="", RPR_RPRLIBTEXTE19="",'
    +'RPR_RPRLIBTEXTE20="", RPR_RPRLIBTEXTE21="", RPR_RPRLIBTEXTE22="",'
    +'RPR_RPRLIBTEXTE23="", RPR_RPRLIBTEXTE24="", RPR_RPRLIBTEXTE25="",'
    +'RPR_RPRLIBTEXTE26="", RPR_RPRLIBTEXTE27="", RPR_RPRLIBTEXTE28="", RPR_RPRLIBTEXTE29="",'
    +'RPR_RPRLIBMUL10="", RPR_RPRLIBMUL11="", RPR_RPRLIBMUL12="", RPR_RPRLIBMUL13="",'
    +'RPR_RPRLIBMUL14="", RPR_RPRLIBMUL15="", RPR_RPRLIBMUL16="",'
    +'RPR_RPRLIBMUL17="", RPR_RPRLIBMUL18="", RPR_RPRLIBMUL19="", RPR_RPRLIBBOOL10="-",'
    +'RPR_RPRLIBBOOL11="-", RPR_RPRLIBBOOL12="-", RPR_RPRLIBBOOL13="-",'
    +'RPR_RPRLIBBOOL14="-", RPR_RPRLIBBOOL15="-", RPR_RPRLIBBOOL16="-",'
    +'RPR_RPRLIBBOOL17="-", RPR_RPRLIBBOOL18="-", RPR_RPRLIBBOOL19="-",'
    +'RPR_RPRLIBVAL10=0, RPR_RPRLIBVAL11=0,'
    +'RPR_RPRLIBVAL12=0, RPR_RPRLIBVAL13=0, RPR_RPRLIBVAL14=0,'
    +'RPR_RPRLIBVAL15=0, RPR_RPRLIBVAL16=0, RPR_RPRLIBVAL17=0, RPR_RPRLIBVAL18=0,'
    +'RPR_RPRLIBVAL19=0,'
    +'RPR_RPRLIBDATE10="'+USDATETime(iDate1900)+'",'
    +'RPR_RPRLIBDATE11="'+USDATETime(iDate1900)+'",'
    +'RPR_RPRLIBDATE12="'+USDATETime(iDate1900)+'",'
    +'RPR_RPRLIBDATE13="'+USDATETime(iDate1900)+'",'
    +'RPR_RPRLIBDATE14="'+USDATETime(iDate1900)+'",'
    +'RPR_RPRLIBDATE15="'+USDATETime(iDate1900)+'",'
    +'RPR_RPRLIBDATE16="'+USDATETime(iDate1900)+'",'
    +'RPR_RPRLIBDATE17="'+USDATETime(iDate1900)+'",'
    +'RPR_RPRLIBDATE18="'+USDATETime(iDate1900)+'",'
    +'RPR_RPRLIBDATE19="'+USDATETime(iDate1900)+'"');
  end;
  
  // GC
  TransfertPlanRevision ;
  TransfertCycleRevision ;


end;

  //Marie-Christine DESSEIGNET Demande N° 1242
Procedure RT_InsertLibellesLibres;
var i: integer;
Begin

  // RTLIBPROJET
  for i := 1 to 5 do InsertChoixCode('RRP', 'TL' + intTostr(i), '.-Projet Table libre ' + intTostr(i), '', '');

  // RTLIBOPERATION
  for i := 1 to 5 do InsertChoixCode('RRO', 'TL' + intTostr(i), '.-Opération Table libre ' + intTostr(i), '', '');

  // RTLIBPERSPECTIVE
  for i := 1 to 5 do InsertChoixCode('RRS', 'TL' + intTostr(i), '.-Proposition Table libre ' + intTostr(i), '', '');

  // RTLIBCIBLAGE
  for i := 1 to 5 do InsertChoixCode('RCI', 'DL' + intTostr(i), '.-Ciblage Date libre ' + intTostr(i), '', '');
  for i := 1 to 5 do InsertChoixCode('RCI', 'TL' + intTostr(i), '.-Ciblage Table libre ' + intTostr(i), '', '');
  for i := 1 to 5 do InsertChoixCode('RCI', 'TX' + intTostr(i), '.-Ciblage Texte libre ' + intTostr(i), '', '');

  // RTLIBCHAMPSLIBRES
  InsertChoixCode('RLZ', 'BLA', '.-Prospect Booléen libre 10', '', '');
  InsertChoixCode('RLZ', 'BLB', '.-Prospect Booléen libre 11', '', '');
  InsertChoixCode('RLZ', 'BLC', '.-Prospect Booléen libre 12', '', '');
  InsertChoixCode('RLZ', 'BLD', '.-Prospect Booléen libre 13', '', '');
  InsertChoixCode('RLZ', 'BLE', '.-Prospect Booléen libre 14', '', '');
  InsertChoixCode('RLZ', 'BLF', '.-Prospect Booléen libre 15', '', '');
  InsertChoixCode('RLZ', 'BLG', '.-Prospect Booléen libre 16', '', '');
  InsertChoixCode('RLZ', 'BLH', '.-Prospect Booléen libre 17', '', '');
  InsertChoixCode('RLZ', 'BLI', '.-Prospect Booléen libre 18', '', '');
  InsertChoixCode('RLZ', 'BLJ', '.-Prospect Booléen libre 19', '', '');
  InsertChoixCode('RLZ', 'DLA', '.-Prospect Date Libre 10', '', '');
  InsertChoixCode('RLZ', 'DLB', '.-Prospect Date Libre 11', '', '');
  InsertChoixCode('RLZ', 'DLC', '.-Prospect Date Libre 12', '', '');
  InsertChoixCode('RLZ', 'DLD', '.-Prospect Date Libre 13', '', '');
  InsertChoixCode('RLZ', 'DLE', '.-Prospect Date Libre 14', '', '');
  InsertChoixCode('RLZ', 'DLF', '.-Prospect Date Libre 15', '', '');
  InsertChoixCode('RLZ', 'DLG', '.-Prospect Date Libre 16', '', '');
  InsertChoixCode('RLZ', 'DLH', '.-Prospect Date Libre 17', '', '');
  InsertChoixCode('RLZ', 'DLI', '.-Prospect Date Libre 18', '', '');
  InsertChoixCode('RLZ', 'DLJ', '.-Prospect Date Libre 19', '', '');
  InsertChoixCode('RLZ', 'MLA', '.-Prospect Multi-choix libre 10', '', '');
  InsertChoixCode('RLZ', 'MLB', '.-Prospect Multi-choix libre 11', '', '');
  InsertChoixCode('RLZ', 'MLC', '.-Prospect Multi-choix libre 12', '', '');
  InsertChoixCode('RLZ', 'MLD', '.-Prospect Multi-choix libre 13', '', '');
  InsertChoixCode('RLZ', 'MLE', '.-Prospect Multi-choix libre 14', '', '');
  InsertChoixCode('RLZ', 'MLF', '.-Prospect Multi-choix libre 15', '', '');
  InsertChoixCode('RLZ', 'MLG', '.-Prospect Multi-choix libre 16', '', '');
  InsertChoixCode('RLZ', 'MLH', '.-Prospect Multi-choix libre 17', '', '');
  InsertChoixCode('RLZ', 'MLI', '.-Prospect Multi-choix libre 18', '', '');
  InsertChoixCode('RLZ', 'MLJ', '.-Prospect Multi-choix libre 19', '', '');
  InsertChoixCode('RLZ', 'TLA', '.-Prospect Texte libre 10', '', '');
  InsertChoixCode('RLZ', 'TLB', '.-Prospect Texte libre 11', '', '');
  InsertChoixCode('RLZ', 'TLC', '.-Prospect Texte libre 12', '', '');
  InsertChoixCode('RLZ', 'TLD', '.-Prospect Texte libre 13', '', '');
  InsertChoixCode('RLZ', 'TLE', '.-Prospect Texte libre 14', '', '');
  InsertChoixCode('RLZ', 'TLF', '.-Prospect Texte libre 15', '', '');
  InsertChoixCode('RLZ', 'TLG', '.-Prospect Texte libre 16', '', '');
  InsertChoixCode('RLZ', 'TLH', '.-Prospect Texte libre 17', '', '');
  InsertChoixCode('RLZ', 'TLI', '.-Prospect Texte libre 18', '', '');
  InsertChoixCode('RLZ', 'TLJ', '.-Prospect Texte libre 19', '', '');
  InsertChoixCode('RLZ', 'TLK', '.-Prospect Texte libre 20', '', '');
  InsertChoixCode('RLZ', 'TLL', '.-Prospect Texte libre 21', '', '');
  InsertChoixCode('RLZ', 'TLM', '.-Prospect Texte libre 22', '', '');
  InsertChoixCode('RLZ', 'TLN', '.-Prospect Texte libre 23', '', '');
  InsertChoixCode('RLZ', 'TLO', '.-Prospect Texte libre 24', '', '');
  InsertChoixCode('RLZ', 'TLP', '.-Prospect Texte libre 25', '', '');
  InsertChoixCode('RLZ', 'TLQ', '.-Prospect Texte libre 26', '', '');
  InsertChoixCode('RLZ', 'TLR', '.-Prospect Texte libre 27', '', '');
  InsertChoixCode('RLZ', 'TLS', '.-Prospect Texte libre 28', '', '');
  InsertChoixCode('RLZ', 'TLT', '.-Prospect Texte libre 29', '', '');
  InsertChoixCode('RLZ', 'VLA', '.-Prospect Valeur libre 10', '', '');
  InsertChoixCode('RLZ', 'VLB', '.-Prospect Valeur libre 11', '', '');
  InsertChoixCode('RLZ', 'VLC', '.-Prospect Valeur libre 12', '', '');
  InsertChoixCode('RLZ', 'VLD', '.-Prospect Valeur libre 13', '', '');
  InsertChoixCode('RLZ', 'VLE', '.-Prospect Valeur libre 14', '', '');
  InsertChoixCode('RLZ', 'VLF', '.-Prospect Valeur libre 15', '', '');
  InsertChoixCode('RLZ', 'VLG', '.-Prospect Valeur libre 16', '', '');
  InsertChoixCode('RLZ', 'VLH', '.-Prospect Valeur libre 17', '', '');
  InsertChoixCode('RLZ', 'VLI', '.-Prospect Valeur libre 18', '', '');
  InsertChoixCode('RLZ', 'VLJ', '.-Prospect Valeur libre 19', '', '');

  // RTLIBCHPLIBSUSPECTS
  InsertChoixCode('RSZ', 'BLA', '.-Booléen libre suspect 11', '', '');
  InsertChoixCode('RSZ', 'BLB', '.-Booléen libre suspect 12', '', '');
  InsertChoixCode('RSZ', 'BLC', '.-Booléen libre suspect 13', '', '');
  InsertChoixCode('RSZ', 'BLD', '.-Booléen libre suspect 14', '', '');
  InsertChoixCode('RSZ', 'BLE', '.-Booléen libre suspect 15', '', '');
  InsertChoixCode('RSZ', 'BLF', '.-Booléen libre suspect 16', '', '');
  InsertChoixCode('RSZ', 'BLG', '.-Booléen libre suspect 17', '', '');
  InsertChoixCode('RSZ', 'BLH', '.-Booléen libre suspect 18', '', '');
  InsertChoixCode('RSZ', 'BLI', '.-Booléen libre suspect 19', '', '');
  InsertChoixCode('RSZ', 'BLJ', '.-Booléen libre suspect 20', '', '');
  InsertChoixCode('RSZ', 'CLA', '.-Table libre suspect 11', '', '');
  InsertChoixCode('RSZ', 'CLB', '.-Table libre suspect 12', '', '');
  InsertChoixCode('RSZ', 'CLC', '.-Table libre suspect 13', '', '');
  InsertChoixCode('RSZ', 'CLD', '.-Table libre suspect 14', '', '');
  InsertChoixCode('RSZ', 'CLE', '.-Table libre suspect 15', '', '');
  InsertChoixCode('RSZ', 'CLF', '.-Table libre suspect 16', '', '');
  InsertChoixCode('RSZ', 'CLG', '.-Table libre suspect 17', '', '');
  InsertChoixCode('RSZ', 'CLH', '.-Table libre suspect 18', '', '');
  InsertChoixCode('RSZ', 'CLI', '.-Table libre suspect 19', '', '');
  InsertChoixCode('RSZ', 'CLJ', '.-Table libre suspect 20', '', '');
  InsertChoixCode('RSZ', 'CLK', '.-Table libre suspect 21', '', '');
  InsertChoixCode('RSZ', 'CLL', '.-Table libre suspect 22', '', '');
  InsertChoixCode('RSZ', 'CLM', '.-Table libre suspect 23', '', '');
  InsertChoixCode('RSZ', 'CLN', '.-Table libre suspect 24', '', '');
  InsertChoixCode('RSZ', 'CLO', '.-Table libre suspect 25', '', '');
  InsertChoixCode('RSZ', 'CLP', '.-Table libre suspect 26 ', '', '');
  InsertChoixCode('RSZ', 'CLQ', '.-Table libre suspect 27 sur 6', '', '');
  InsertChoixCode('RSZ', 'CLR', '.-Table libre suspect 28 sur 6', '', '');
  InsertChoixCode('RSZ', 'CLS', '.-Table libre suspect 29 sur 6', '', '');
  InsertChoixCode('RSZ', 'CLT', '.-Table libre suspect 30 sur 6', '', '');
  InsertChoixCode('RSZ', 'CLU', '.-Table libre suspect 31 sur 6', '', '');
  InsertChoixCode('RSZ', 'CLV', '.-Table libre suspect 32 sur 6', '', '');
  InsertChoixCode('RSZ', 'CLW', '.-Table libre suspect 33 sur 6', '', '');
  InsertChoixCode('RSZ', 'CLX', '.-Table libre suspect 34 sur 6', '', '');
  InsertChoixCode('RSZ', 'CLY', '.-Table libre suspect 35 sur 6', '', '');
  InsertChoixCode('RSZ', 'CLZ', '.-Table libre suspect 36 sur 6', '', '');
  InsertChoixCode('RSZ', 'DLA', '.-Date libre suspect 11', '', '');
  InsertChoixCode('RSZ', 'DLB', '.-Date libre suspect 12', '', '');
  InsertChoixCode('RSZ', 'DLC', '.-Date libre suspect 13', '', '');
  InsertChoixCode('RSZ', 'DLD', '.-Date libre suspect 14', '', '');
  InsertChoixCode('RSZ', 'DLE', '.-Date libre suspect 15', '', '');
  InsertChoixCode('RSZ', 'DLF', '.-Date libre suspect 16', '', '');
  InsertChoixCode('RSZ', 'DLG', '.-Date libre suspect 17', '', '');
  InsertChoixCode('RSZ', 'DLH', '.-Date libre suspect 18', '', '');
  InsertChoixCode('RSZ', 'DLI', '.-Date libre suspect 19', '', '');
  InsertChoixCode('RSZ', 'DLJ', '.-Date libre suspect 20', '', '');
  InsertChoixCode('RSZ', 'ML0', '.-multi-champs libre suspect 1', '', '');
  InsertChoixCode('RSZ', 'ML1', '.-multi-champs libre suspect 2', '', '');
  InsertChoixCode('RSZ', 'ML2', '.-multi-champs libre suspect 3', '', '');
  InsertChoixCode('RSZ', 'ML3', '.-multi-champs libre suspect 4', '', '');
  InsertChoixCode('RSZ', 'ML4', '.-multi-champs libre suspect 5', '', '');
  InsertChoixCode('RSZ', 'ML5', '.-multi-champs libre suspect 6', '', '');
  InsertChoixCode('RSZ', 'ML6', '.-multi-champs libre suspect 7', '', '');
  InsertChoixCode('RSZ', 'ML7', '.-multi-champs libre suspect 8', '', '');
  InsertChoixCode('RSZ', 'ML8', '.-multi-champs libre suspect 9', '', '');
  InsertChoixCode('RSZ', 'ML9', '.-multi-champs libre suspect 10', '', '');
  InsertChoixCode('RSZ', 'MLA', '.-multi-champs libre suspect 11', '', '');
  InsertChoixCode('RSZ', 'MLB', '.-multi-champs libre suspect 12', '', '');
  InsertChoixCode('RSZ', 'MLC', '.-multi-champs libre suspect 13', '', '');
  InsertChoixCode('RSZ', 'MLD', '.-multi-champs libre suspect 14', '', '');
  InsertChoixCode('RSZ', 'MLE', '.-multi-champs libre suspect 15', '', '');
  InsertChoixCode('RSZ', 'MLF', '.-multi-champs libre suspect 16', '', '');
  InsertChoixCode('RSZ', 'MLG', '.-multi-champs libre suspect 17', '', '');
  InsertChoixCode('RSZ', 'MLH', '.-multi-champs libre suspect 18', '', '');
  InsertChoixCode('RSZ', 'MLI', '.-multi-champs libre suspect 19', '', '');
  InsertChoixCode('RSZ', 'MLJ', '.-multi-champs libre suspect 20', '', '');
  InsertChoixCode('RSZ', 'TLA', '.-Texte libre suspect 11', '', '');
  InsertChoixCode('RSZ', 'TLB', '.-Texte libre suspect 12', '', '');
  InsertChoixCode('RSZ', 'TLC', '.-Texte libre suspect 13', '', '');
  InsertChoixCode('RSZ', 'TLD', '.-Texte libre suspect 14', '', '');
  InsertChoixCode('RSZ', 'TLE', '.-Texte libre suspect 15', '', '');
  InsertChoixCode('RSZ', 'TLF', '.-Texte libre suspect 16', '', '');
  InsertChoixCode('RSZ', 'TLG', '.-Texte libre suspect 17', '', '');
  InsertChoixCode('RSZ', 'TLH', '.-Texte libre suspect 18', '', '');
  InsertChoixCode('RSZ', 'TLI', '.-Texte libre suspect 19', '', '');
  InsertChoixCode('RSZ', 'TLJ', '.-Texte libre suspect 20', '', '');
  InsertChoixCode('RSZ', 'TLK', '.-Texte libre suspect 21', '', '');
  InsertChoixCode('RSZ', 'TLL', '.-Texte libre suspect 22', '', '');
  InsertChoixCode('RSZ', 'TLM', '.-Texte libre suspect 23', '', '');
  InsertChoixCode('RSZ', 'TLN', '.-Texte libre suspect 24', '', '');
  InsertChoixCode('RSZ', 'TLO', '.-Texte libre suspect 25', '', '');
  InsertChoixCode('RSZ', 'TLP', '.-Texte libre suspect 26', '', '');
  InsertChoixCode('RSZ', 'TLQ', '.-Texte libre suspect 27', '', '');
  InsertChoixCode('RSZ', 'TLR', '.-Texte libre suspect 28', '', '');
  InsertChoixCode('RSZ', 'TLS', '.-Texte libre suspect 29', '', '');
  InsertChoixCode('RSZ', 'TLT', '.-Texte libre suspect 30', '', '');
  InsertChoixCode('RSZ', 'VLA', '.-montant libre suspect 11', '', '');
  InsertChoixCode('RSZ', 'VLB', '.-montant libre suspect 12', '', '');
  InsertChoixCode('RSZ', 'VLC', '.-montant libre suspect 13', '', '');
  InsertChoixCode('RSZ', 'VLD', '.-montant libre suspect 14', '', '');
  InsertChoixCode('RSZ', 'VLE', '.-montant libre suspect 15', '', '');
  InsertChoixCode('RSZ', 'VLF', '.-montant libre suspect 16', '', '');
  InsertChoixCode('RSZ', 'VLG', '.-montant libre suspect 17', '', '');
  InsertChoixCode('RSZ', 'VLH', '.-montant libre suspect 18', '', '');
  InsertChoixCode('RSZ', 'VLI', '.-montant libre suspect 19', '', '');
  InsertChoixCode('RSZ', 'VLJ', '.-montant libre suspect 20', '', '');

End;

Procedure RTMoveRessources;
Var Q: TQuery;
    StListe, stRessource, Sql, stIntervenant, ListeOut, ListeRech : string;
Begin
  {
    Laurent Abélard le 27 Septembre 2007
    Comme ce traitement dure plusieurs dizaines d'heures sur la base de SIC (3627117 lignes dans ACTIONS)
    On découpe le traitement en deux passes.
    1) Traiter les champs avec une seul intervention en une seule requête INSERT INTO SELECT. (2830919)
    2) Traiter les autres dans la boucle. (9095)
    ATTENTION de bien ignorer les records vides. (787103)
    Je supprime aussi la transaction qui rempli violement les logs alors que si cela plante, il faudra
    de toute façons restaurer la Base.
  }
  //BeginTrans;

  Sql := 'INSERT INTO ACTIONINTERVENANT (RAI_AUXILIAIRE,RAI_NUMACTION,RAI_RESSOURCE,RAI_GUID)';
  Sql := Sql + ' SELECT RAC_AUXILIAIRE,RAC_NUMACTION, LEFT(RAC_INTERVINT,LEN(RAC_INTERVINT)-1),PGIGUID FROM ACTIONS';
  Sql := Sql + ' INNER JOIN RESSOURCE ON ARS_RESSOURCE = LEFT(RAC_INTERVINT,LEN(RAC_INTERVINT)-1)';
  Sql := Sql + ' WHERE (RAC_INTERVINT LIKE "%;" AND RAC_INTERVINT NOT LIKE "%;%;")';

  ExecuteSQLContOnExcept(Sql);

  ExecuteSQLContOnExcept('UPDATE ACTIONS SET RAC_INTERVINT="" WHERE (RAC_INTERVINT LIKE "%;" AND RAC_INTERVINT NOT LIKE "%;%;")'+
  ' AND EXISTS (SELECT ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE = LEFT(RAC_INTERVINT,LEN(RAC_INTERVINT)-1))');

  Sql:='';
  Q := OpenSQL('SELECT RAC_INTERVINT,RAC_AUXILIAIRE,RAC_NUMACTION FROM ACTIONS WHERE RAC_INTERVINT LIKE "%;%;%"', True);
  While Not Q.EOF Do
  Begin
    StListe := Q.FindField('RAC_INTERVINT').AsString;
    If StListe <> '' Then
    Begin

      //Claude DUMAS Le 21/01/2008 Version 8.1.850.18 Demande N° 2075
      // CD le 14/12/07 pour supprimer les ressources figurant plusieurs fois ds la liste
      ListeOut := '';
      Repeat
        stintervenant := ReadTokenSt (StListe);
        If stintervenant <> '' Then
        Begin
          ListeRech := ';' + StListe;
          If Pos((';'+stIntervenant+';'),ListeRech) = 0 Then
            ListeOut := ListeOut + stIntervenant + ';';
        End;
      Until StListe = '';
      StListe := ListeOut;
      // CD le 14/12/07 pour supprimer les ressources figurant plusieurs fois ds la liste

      Repeat
        StRessource := Trim(ReadTokenSt(StListe));
        If StRessource <> '' Then
        Begin
          If ExisteSql ('SELECT ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE="'+StRessource+'"') Then
          Begin
            Sql := 'INSERT INTO ACTIONINTERVENANT (RAI_AUXILIAIRE,RAI_NUMACTION,RAI_RESSOURCE,RAI_GUID) VALUES ("';
            Sql := Sql + Q.FindField('RAC_AUXILIAIRE').AsString + '",' ;
            Sql := Sql + IntToStr(Q.FindField('RAC_NUMACTION').AsInteger) + ',"';
            Sql := Sql + StRessource + '","';
            Sql := Sql + AglGetGuid() + '")';
            ExecuteSQLContOnExcept(Sql);
          End;
        End;
      Until StRessource = '';
    Q.Next;
    End;
  End;
  Ferme(Q);

  ExecuteSQLContOnExcept('Update ACTIONS set RAC_INTERVINT="" where RAC_INTERVINT <> ""');

  //CommitTrans;
End;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : MF
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. : Mise à jour PDP_CODIFALSACE (DUCSPARAM) et
Suite ........ : PDF_CODIFALSACE(DUCSAFFECT) pour les codifications
Suite ........ : "dossier" et "STD".
Suite ........ : Ces champs prennent la valeur de PDP_CODIFALSACE
Suite ........ : de la codification "CEG" identique
Mots clefs ... :
*****************************************************************}
procedure majAlsaceMoselle();
var
  QCEG                        : TQuery;
  StSql                       : string;
  Codification,CodifAlsace    : string;

begin
  QCEG := nil;
  try

    StSql :=  'SELECT PDP_CODIFICATION, PDP_CODIFALSACE FROM DUCSPARAM WHERE '+
              'PDP_PREDEFINI="CEG"';
    QCEG := OpenSql (Stsql,TRUE);
    if QCEG.Eof then Exit;

    while not QCEG.Eof do
    begin
      Codification   := QCEG.FindField('PDP_CODIFICATION').AsString;
      CodifAlsace := QCEG.FindField('PDP_CODIFALSACE').AsString;

      ExecuteSQLContOnExcept('UPDATE DUCSPARAM SET PDP_CODIFALSACE="'+CodifAlsace+'" '+
                'WHERE PDP_CODIFICATION = "'+Codification+'" AND '+
                '(PDP_PREDEFINI="STD" OR PDP_PREDEFINI="DOS")');
      ExecuteSQLContOnExcept('UPDATE DUCSAFFECT SET PDF_CODIFALSACE="'+CodifAlsace+'" '+
                'WHERE PDF_CODIFICATION = "'+Codification+'" AND '+
                '(PDF_PREDEFINI="STD" OR PDF_PREDEFINI="DOS")');
      QCEG.Next;
    end;
  finally
    Ferme( QCEG );
  end;
end;

Procedure MajVer832 ;
Begin
  if not IsDossierPCL then
  begin
    // MCD
    {
      Laurent Abélard le 27 Septembre 2007, Il apparaît que plusieurs UPDATE sur LIGNECOMPL sont très long
      et nécessitent de regrouper tout ces UPDATE en un seul (Voir en Fin de MajVer822()).
    ExecuteSQLContOnExcept('UPDATE LIGNECOMPL  SET GLC_COEFFREVALO=0,'
    +'GLC_DATEDEBUTFAC="'+USDATETime(iDate1900)+'",'
    +'GLC_DATEFINFAC="'+USDATETime(iDate2099)+'",'
    +'GLC_DATEAUGMENT="'+USDATETime(iDate1900)+'"');
    }
    ExecuteSQLContOnExcept('update PROJETS set RPJ_RPJTABLELIBRE1="", RPJ_RPJTABLELIBRE2="", RPJ_RPJTABLELIBRE3="", RPJ_RPJTABLELIBRE4="", RPJ_RPJTABLELIBRE5="" ');
    ExecuteSQLContOnExcept('update OPERATIONS set ROP_ROPTABLELIBRE1="", ROP_ROPTABLELIBRE2="", ROP_ROPTABLELIBRE3="", ROP_ROPTABLELIBRE4="", ROP_ROPTABLELIBRE5="" ');
    ExecuteSQLContOnExcept('update PERSPECTIVES set RPE_RPETABLELIBRE1="", RPE_RPETABLELIBRE2="", RPE_RPETABLELIBRE3="", RPE_RPETABLELIBRE4="", RPE_RPETABLELIBRE5="" ');
    ExecuteSQLContOnExcept('update PERSPHISTO set RPH_RPETABLELIBRE1="", RPH_RPETABLELIBRE2="", RPH_RPETABLELIBRE3="", RPH_RPETABLELIBRE4="", RPH_RPETABLELIBRE5="" ');
    // PL
    RT_InsertLibellesLibres();
    // MNG
    RTMoveRessources();

    // GM
    ExecuteSQLContOnExcept('UPDATE AFFAIRE  Set  AFF_CODECUTOFFVTE="",AFF_CODECUTOFFACH="",AFF_DATEAUGMENT="'+USDATETime(iDate1900)+'"');
    ExecuteSQLContOnExcept('UPDATE AFCUMUL Set ACU_CODEFORMCUTOFF=""');

    // MCD
    AglNettoieListesPlus('AFMULFACTIERSAFF', 'AFF_DOMAINE;AFF_ETABLISSEMENT',nil,true);
    AglNettoieListesPlus('AFMULFACTAFFAPP', 'AFF_DOMAINE;AFF_ETABLISSEMENT',nil,true);
  end;
  if IsMonoOuCommune then
  begin
    ExecuteSQLContOnExcept('UPDATE DPFISCAL SET DFI_GUIDPERCDISIE="", DFI_GUIDPERCCI=""');
    ExecuteSQLContOnExcept('UPDATE DPDOCUMENT SET DPD_OBJETDOC=""');
  end;
  // VG
  ExecuteSQLContOnExcept('UPDATE COTISATION SET PCT_DADSBASEPREV="-"');
  ExecuteSQLContOnExcept('UPDATE DADSPERIODES SET PDE_DADSCDC="V08R04"');
  ExecuteSQLContOnExcept('UPDATE ABSENCESALARIE set PCN_HDEB = "'+UsDateTime(IDate1900)+'", PCN_HFIN = "'+UsDateTime(IDate1900)+'",PCN_NBHEURESNUIT = 0');
  ExecuteSQLContOnExcept('UPDATE VARIABLEPAIE SET PVA_TYPEVARIABLE ="PAI",PVA_VARPERIODICITE=""');
  ExecuteSQLContOnExcept('UPDATE MOTIFABSENCE set PMA_TYPEMOTIF = "ABS", PMA_CTRLABSEXISTE="-", PMA_CTRLPLAGEH="-",PMA_CTRLPREEXISTE="-",PMA_PGCOLORPRE=""');
  ExecuteSQLContOnExcept('UPDATE ETABCOMPL SET ETB_CODETVT="", ETB_CODETTAT=""');
  ExecuteSQLContOnExcept('UPDATE SALARIES SET PSA_CODETVT="", PSA_CODETTAT="", PSA_TYPCODETVT="ETB", PSA_TYPCODETTAT="ETB"');
  ExecuteSQLContOnExcept('UPDATE REMUNERATION SET PRM_MOIS5="", PRM_MOIS6=""');
  ExecuteSQLContOnExcept('UPDATE CODTAB SET PTI_CODTABLDYN = ""');
  ExecuteSQLContOnExcept('UPDATE METHCALCULSALMOY SET PSM_PRORATISE = "-"');
  ExecuteSQLContOnExcept('UPDATE RESULTSIMUL SET PSR_CODTABSUR = "", PSR_CODTABDROIT = "", PSR_CODTABREVAL = "", PSR_CODTABTUR = "", PSR_QUALIFICATION = "", PSR_CODEEMPLOI = "", PSR_CATBILAN = "", PSR_CATDADS = "", PSR_DATECREATION = "'+UsDateTime(Idate1900)+'", PSR_CREATEUR = "" ');
  ExecuteSQLContOnExcept('UPDATE SALAIREMOYEN SET PSY_SALAIREMOYPER = 0, PSY_SALAIREMOYDERM = 0, PSY_DATEENTREEGRP = "'+UsDateTime(Idate1900)+'", PSY_DATEANCIENNETE = "'+UsDateTime(Idate1900)+'", PSY_NBANNEESERV = 0,PSY_COEFFICIENT = "", PSY_QUALIFICATION = "", PSY_CODEEMPLOI = "", PSY_CATBILAN = "", PSY_CATDADS = "", PSY_DATECREATION = "'+UsDateTime(Idate1900)+'", PSY_CREATEUR = ""');
  ExecuteSQLContOnExcept('UPDATE ENFANTSALARIE SET PEF_TYPEPARENTAL = "001"');
  ExecuteSQLContOnExcept('UPDATE DPTABGENPAIE SET DT1_LIBREMONTANT1=0,DT1_LIBREMONTANT2=0,DT1_LIBREMONTANT3=0,'
  +'DT1_LIBREMONTANT4=0,DT1_LIBREMONTANT5=0,DT1_LIBREMONTANT6=0,DT1_LIBREMONTANT7=0,'
  +'DT1_LIBREMONTANT8=0,DT1_LIBREMONTANT9=0,DT1_LIBREMONTANT10=0,DT1_LIBREMONTANT11=0,DT1_LIBREMONTANT12=0,DT1_LIBREMONTANT13=0,DT1_LIBREMONTANT14=0,DT1_LIBREMONTANT15=0');

  majAlsaceMoselle();
  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE="PGHISTODETAIL"');
  ExecuteSQLContOnExcept('UPDATE SESSIONSTAGE SET PSS_VOIRPORTAIL="-"');
  ExecuteSQLContOnExcept('UPDATE LIEUFORMATION SET PLF_TELEPHONE="",PLF_FAX=""');

  // CA
  ExecuteSQLContOnExcept('DELETE FROM CIB WHERE TCI_BANQUE = "" OR TCI_BANQUE IS NULL');
  if V_PGI.ModePCL <> '1' then
  begin
    if ExisteSQL('SELECT TCI_CODECIB FROM CIB WHERE TCI_BANQUE = "@ID" AND (TCI_PREDEFINI = "" OR TCI_PREDEFINI IS NULL)') then
      ExecuteSQLContOnExcept('DELETE FROM CIB WHERE TCI_PREDEFINI = "CEG"');
    ExecuteSQLContOnExcept('UPDATE CIB SET TCI_PREDEFINI = "STD" WHERE TCI_PREDEFINI<>"CEG"');
  end;
end;



// MB
procedure MiseAJourGRPDonnees();
var Imin, IResult, IResult2 : Integer ;
  sCode,sLib : String ;
  IsOK : Boolean ;
  Q,QR : TQuery;
begin
  // Lecture de choixcod et deversement dans GRPDONNEES
  Imin := 0 ;
  IsOK := True ;
  BEGINTRANS ;
  Q := OpenSQL('Select CC_CODE,CC_LIBELLE from CHOIXCOD WHERE CC_TYPE = "UCO"', True);
  while not Q.EOF do
  begin
    sCode := Q.FindField('CC_CODE').AsString ;
    sLib  := Q.FindField('CC_LIBELLE').AsString ;
    try
        QR := OpenSQL('SELECT GRP_CODE FROM GRPDONNEES WHERE GRP_NOM = "GROUPECONF" AND GRP_CODE="'
        +sCode+'"',true);
        if QR.RecordCount = 0 then
        begin
           try
              IResult := ExecuteSQLContOnExcept('INSERT INTO GRPDONNEES(GRP_NOM,GRP_ID,GRP_IDPERE,GRP_CODE,GRP_LIBELLE,GRP_COMMENTAIRE) '
              +' VALUES("GROUPECONF",'+IntToStr(Imin)+',-1,"'+sCode+'","'+sLib+'","")') ;
              if IResult = 1 then // OK je traite les tables USERCONF et DOSSIERGRP liés
              begin
                  IResult2 := ExecuteSQLContOnExcept('INSERT INTO LIENDONNEES (LND_NOM,LND_GRPID, LND_PROFILID, LND_USERID) '
                  +' SELECT "GROUPECONF",'+inttostr(Imin)+',"",UCO_USER FROM OLDUSERCONF WHERE UCO_GROUPECONF = "'+sCode+'"');
                  if IResult2 >= 0 then
                  begin

                     if TableExiste('OLDDOSSIERGRP') then // cas 750 ou supérieur
                         IResult2 := ExecuteSQLContOnExcept('INSERT INTO LIENDOSGRP (LDO_NOM,LDO_GRPID, LDO_MARK, LDO_NODOSSIER) '
                         +' SELECT "GROUPECONF",'+inttostr(Imin)+',"X",DOG_NODOSSIER FROM OLDDOSSIERGRP WHERE DOG_GROUPECONF = "'+sCode+'"')
                     else // Sinon récupére de la table dossier directement.
                         IResult2 := ExecuteSQLContOnExcept('INSERT INTO LIENDOSGRP (LDO_NOM,LDO_GRPID, LDO_MARK, LDO_NODOSSIER) '
                         +' SELECT "GROUPECONF",'+inttostr(Imin)+',"X",DOS_NODOSSIER FROM DOSSIER WHERE DOS_GROUPECONF = "'+sCode+'"');

                     if IResult2 >= 0 then
                     begin
                       ExecuteSQLContOnExcept('DELETE FROM OLDUSERCONF WHERE UCO_GROUPECONF = "'+sCode+'"');
                       if TableExiste('OLDDOSSIERGRP') then // cas 750 ou supérieur
                         ExecuteSQLContOnExcept('DELETE FROM OLDDOSSIERGRP WHERE DOG_GROUPECONF = "'+sCode+'"');
                     end;

                  end;
              end;
           except
              on e:exception do
              begin
                 PGIBox(e.Message);
                 IResult := -1 ;
                 IsOK := False ;
              end;
           end;
           if IResult = 1 then
              ExecuteSQLContOnExcept('DELETE FROM CHOIXCOD WHERE CC_TYPE = "UCO" AND CC_CODE = "'+sCode+'"') ;
           Inc(Imin) ;
           If not IsOK then break ;
        end;
        Ferme(QR);
    except
      on e:exception do
      begin
         PGIBox(e.Message);
         IsOK := False ;
      end;
    end;
    If not IsOK then break ;
    Q.Next;
  end;
  Ferme(Q);
  //
  if not IsOK then
     ROLLBACK
  else
     COMMITTRANS ;
end;


Procedure MajVer833 ;
Begin
  if not IsDossierPCL then
  begin
    // JLS
    ExecuteSQLContOnExcept('UPDATE PORT SET GPO_MODEGROUPEPORT="" WHERE GPO_MODEGROUPEPORT IS NULL');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_TARIFGENPNPB="-" WHERE GPP_TARIFGENPNPB IS NULL');

    { Initialisation nouveaux champs de la table RESSOURCE }
    ExecuteSQLContOnExcept('UPDATE RESSOURCE SET ARS_FERMEPDC="-",ARS_CAPAJOURPDC=0,ARS_PCTMAXOFPDC=0 WHERE ARS_FERMEPDC IS NULL');

    { Initialisation nouveaux champs de la table WORDREGAMME }
    ExecuteSQLContOnExcept('UPDATE WORDREGAMME SET WOG_DATEDEBPDC="'+UsDateTime(iDate1900)+'", WOG_DATEFINPDC ="'+UsDateTime(iDate1900)+'" WHERE WOG_DATEDEBPDC IS NULL');

    { Initialisation nouveaux champs de la table WORDRERES }
    ExecuteSQLContOnExcept('UPDATE WORDRERES SET WOR_DATEDEBPDC="'+UsDateTime(iDate1900)+'", WOR_DATEFINPDC ="'+UsDateTime(iDate1900)+'" WHERE WOR_DATEDEBPDC IS NULL');

    { Initialisation nouveaux champs de la table WORDRELIG }
    ExecuteSQLContOnExcept('UPDATE WORDRELIG SET WOL_METHCALCULPDC="",WOL_CTXPDC ="" WHERE WOL_METHCALCULPDC IS NULL');

    { Consolidation table WCALCULPARAM à partir de WCBNPARAM/WSCMPARAM    }
    wConsolideCalculParam();

    { Mise à jour de EDIADRESSE }
    ExecuteSQLNoPCL('UPDATE EDIADRESSE SET EDA_SECTEURGEO="", EDA_TRANSPORTEUR="" WHERE EDA_SECTEURGEO IS NULL AND EDA_TRANSPORTEUR IS NULL');

    //Jean-Luc SAUZET Demande N° 1554
    {If Not IsDossierPCL Then
      ExecuteSQLContOnExcept('UPDATE LIGNECOMPL SET LIGNECOMPL.GLC_DATEDEM=LIGNE.GL_DATELIVRAISON FROM LIGNE WHERE LIGNECOMPL.GLC_DATEDEM IS NULL and LIGNECOMPL.GLC_NATUREPIECEG=LIGNE.GL_NATUREPIECEG and LIGNECOMPL.GLC_SOUCHE=LIGNE.GL_SOUCHE and LIGNECOMPL.GLC_NUMERO=LIGNE.GL_NUMERO and LIGNECOMPL.GLC_INDICEG=LIGNE.GL_INDICEG and LIGNECOMPL.GLC_NUMORDRE=LIGNE.GL_NUMORDRE');}
    {
      Laurent Abélard le 27 Septembre 2007, Cette requête dure plus de 13 heures sur les 12 Millions de LIGNECOMPL de
      la base SIC. Il apparaît que plusieurs UPDATE sur LIGNECOMPL sont très long et nécessitent de regrouper tout
      ces UPDATE en un seul (Voir en MajVer822()).
    ExecuteSQLNoPCL( 'UPDATE LIGNECOMPL SET GLC_DATEDEM=ISNULL((SELECT MAX(GL_DATELIVRAISON) FROM LIGNE ' +
                     'WHERE GLC_NATUREPIECEG=GL_NATUREPIECEG AND GLC_SOUCHE=GL_SOUCHE AND GLC_NUMERO=GL_NUMERO ' +
                     'AND GLC_INDICEG=GL_INDICEG AND GLC_NUMORDRE=GL_NUMORDRE), "'
                     + UsDateTime(iDate1900) + '") WHERE GLC_DATEDEM IS NULL');
    }

    { Initialisation nouveaux champs de la table WNOMELIG }
    ExecuteSQLContOnExcept('UPDATE WNOMELIG SET WNL_PHASEAPPROLIB=IIF(WNL_OPEITIAPPRO=WNL_OPEITI,WNL_PHASELIB,"") WHERE WNL_PHASEAPPROLIB IS NULL');
  end;
  if IsMonoOuCommune then
  begin
    // MD
    ExecuteSQLContOnExcept('UPDATE DOSSIER SET DOS_VERSIONBASE=0, DOS_NOMSERVEUR=""');
    ExecuteSQLContOnExcept('UPDATE DOSSIER SET DOS_GUIDDOSSIER="" WHERE DOS_GUIDDOSSIER IS NULL');
  end;
  // MB
  ExecuteSQLContOnExcept('UPDATE ANNUBIS SET ANB_DOUBLON=""');
  ExecuteSQLContOnExcept('UPDATE TIERSCOMPL SET YTC_DOUBLON=""');
  // CD
  {$IFDEF MAJPCL}
    if IsdossierPCL then
    begin
      UpdateDomaine ('RTDOCUMENT','Y');
    end;
  {$ENDIF}
  // JTR
  if not IsDossierPCL then
  begin
    AglNettoieListesPlus('GCGROUPEPIECE', 'GP_ECHEBLOQUE',nil,true);
    AglNettoieListesPlus('GCGROUPEMANACH', 'GP_ECHEBLOQUE',nil,true);
    AglNettoieListesPlus('GCGROUPEMANVTE', 'GP_ECHEBLOQUE',nil,true);
    ExecuteSQLContOnExcept('UPDATE STKMOUVEMENT SET GSM_MONTANTCPTA = GSM_MONTANT WHERE GSM_STKTYPEMVT = "PHY" AND GSM_QUALIFMVT IN (SELECT GSN_QUALIFMVT FROM STKNATURE WHERE GSN_GERECOMPTA="X") AND GSM_COMPTABILISE = "X"');
    AglNettoieListesPlus('GCGROUPEPIECE', 'GP_TAUXDEV;GP_TYPETAUXFIXING;GP_TAUXFIXING',nil,False);
    // MCD
    ExecuteSQLContOnExcept('Update AFPLANNINGPARAM Set APP_MODIFIABLE = "-"'
    +' WHERE APP_CODEPARAM = "ABS" OR APP_CODEPARAM = "ACT"');
    ExecuteSQLContOnExcept('Update AFPLANNINGPARAM Set APP_MODIFIABLE = "X"'
    +' WHERE APP_CODEPARAM = "DEC" OR APP_CODEPARAM = "FAC" OR APP_CODEPARAM = "TER"');
  end;
  // PL
  // Correction de la tablette TRTypeCpte suite à une erreur PL (même préfixe tablette RTTypeCiblage)
  ExecuteSQLContOnExcept('delete from commun where co_type="TCB" and co_code="001"');
  ExecuteSQLContOnExcept('delete from commun where co_type="TCB" and co_code="002"');
  // VG
  ExecuteSQLContOnExcept('UPDATE PGAEM SET PGA_NUMOBJETAEM=""');
  ExecuteSQLContOnExcept('UPDATE CONTRATTRAVAIL SET PCI_NUMOBJETAEM=""');
  ExecuteSQLContOnExcept('UPDATE DEPORTSAL SET PSE_ISLIBEMPLOI=""');
  ExecuteSQLContOnExcept('UPDATE SESSIONSTAGE SET PSS_DATEINSC="'+UsdateTime(IDate1900)+'"');
  ExecuteSQLContOnExcept('UPDATE HISTORETENUE SET PHR_TYPEPAIRETSAL="BUL"');
  ExecuteSQLContOnExcept('UPDATE ORGANISMEPAIE SET POG_PGORGGUID=""');
  //PD
  ExecuteSQLContOnExcept('UPDATE DPSOCIAL SET DSO_NUMADHERSYNDIC="",DSO_NUMADHERMETIER="",DSO_SALARIECONF="-",DSO_MONTPROVCP=0,DSO_PROVCPCOMMENT="",DSO_PERIODRENOUVRP="",DSO_DATEELECTIONRP="'+UsdateTime(IDate1900)+'"');
  ExecuteSQLContOnExcept('UPDATE DPTABGENPAIE SET DT1_LIBREMONTANT16=0, DT1_LIBREMONTANT17=0, DT1_LIBREMONTANT18=0, DT1_LIBREMONTANT19=0, DT1_LIBREMONTANT20=0, DT1_LIBREMONTANT21=0, DT1_LIBREMONTANT22=0, DT1_LIBREMONTANT23=0, DT1_LIBREMONTANT24=0, DT1_LIBREMONTANT25=0');
  ExecuteSQLContOnExcept('UPDATE DPTABGENPAIE SET DT1_LIBREMONTANT26=0, DT1_LIBREMONTANT27=0, DT1_LIBREMONTANT28=0, DT1_LIBREMONTANT29=0, DT1_LIBREMONTANT30=0');
  ExecuteSQLContOnExcept('UPDATE DPTABGENPAIE SET DT1_LIBREMONTANT31=0, DT1_LIBREMONTANT32=0, DT1_LIBREMONTANT33=0, DT1_LIBREMONTANT34=0, DT1_LIBREMONTANT35=0');
  ExecuteSQLContOnExcept('UPDATE DPTABGENPAIE SET DT1_LIBREMONTANT36=0, DT1_LIBREMONTANT37=0, DT1_LIBREMONTANT38=0, DT1_LIBREMONTANT39=0, DT1_LIBREMONTANT40=0, DT1_LIBREMONTANT41=0, DT1_LIBREMONTANT42=0, DT1_LIBREMONTANT43=0, DT1_LIBREMONTANT44=0, DT1_LIBREMONTANT45=0');
  ExecuteSQLContOnExcept('UPDATE DPTABGENPAIE SET DT1_LIBREMONTANT46=0, DT1_LIBREMONTANT47=0, DT1_LIBREMONTANT48=0, DT1_LIBREMONTANT49=0, DT1_LIBREMONTANT50=0');

  // JS
  if not IsDossierPCL then
  begin
    ExecuteSQLContOnExcept('UPDATE ARTICLEQTE SET GAF_DEBITSURLISTE="-" WHERE GAF_DEBITSURLISTE IS NULL');
    ExecuteSQLContOnExcept('UPDATE LIGNEFORMULE SET GLF_DEBITSURLISTE="-" WHERE GLF_DEBITSURLISTE IS NULL');
    ExecuteSQLContOnExcept('UPDATE LIGNEFORMULE SET GLF_QTE=0 WHERE GLF_QTE IS NULL');
    ExecuteSQLContOnExcept('UPDATE LIGNEFORMULE SET GLF_QTERESTE=0 WHERE GLF_QTERESTE IS NULL');
    ExecuteSQLContOnExcept('UPDATE LIGNEFORMULE SET GLF_RESULT=0 WHERE GLF_RESULT IS NULL');
    // JLS
    if not ExisteSQL('select 1 from STKNATURE where GSN_QUALIFMVT="SDA"') then
    begin
      ExecuteSQLContOnExcept('INSERT INTO STKNATURE (GSN_QUALIFMVT,GSN_LIBELLE,GSN_STKTYPEMVT,GSN_QTEPLUS,GSN_QUALIFMVTSUIV,GSN_SIGNEMVT,'
+'GSN_STKFLUX,GSN_GERECOMPTA,GSN_CALLGSL,GSN_CALLGSS,GSN_CONTREMARQUE,GSN_CTRLDISPO,GSN_CTRLPEREMPTION,'
+'GSN_MAJPRIXVALO,GSN_SDISPODISPATCH,GSN_SFLUXDISPATCH,GSN_SDISPOPICKING,GSN_SFLUXPICKING,GSN_GERESERIEGRP)'
+' VALUES ("SDA", "Sortie directe sur affaire", "PHY", "GQ_PHYSIQUE;", "", "MIX",'
+' "STO","-","-","-","-","000", "000",'
+' "", "", "", "", "", "-")');
    end;

  end;
  // CA
  AglNettoieListesPlus('TRMULICC', 'TE_USERCOMPTABLE',nil,False);

  // MB
  MiseAJourGRPDonnees();
  if TableExiste('OLDUSERCONF') then
  	DBDeleteView ( DBSoc, v_pgi.driver, 'OLDUSERCONF');
  if TableExiste('OLDDOSSIERGRP') then     //La table DOSSIERGRP n'a été créée qu'en Socref 714 et a disparu en 750 !!!
    DBDeleteView( DBSoc, v_pgi.driver, 'OLDDOSSIERGRP');

// ALR pas utile ...
//  ExecuteSQLContOnExcept('DELETE FROM DECHAMPS WHERE DH_PREFIXE IN (SELECT DT_PREFIXE FROM DETABLES '
//                  +' WHERE DT_NOMTABLE IN ("OLDUSERCONF","OLDDOSSIERGRP"))');
//  ExecuteSQLContOnExcept('DELETE FROM DETABLES WHERE DT_NOMTABLE IN ("OLDUSERCONF","OLDDOSSIERGRP")');

  // VG
  ExecuteSQLContOnExcept('UPDATE ENVOISOCIAL SET PES_GUID1="", PES_GUID2="", PES_GUID3=""');

end;

Procedure MajVer834 ;
Begin
  // VG
  AglNettoieListes('PGCODIFDUCS', 'PDP_CODIFALSACE',nil);
  AglNettoieListes('PGDUCSAFFECT', 'PDF_CODIFALSACE',nil);
  AglNettoieListesPlus ('PGENVOIDUCS', 'PES_GUID1',nil);
  ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA="'+GetParamSocSecur ('SO_PGFORVALOSALAIRE','')
          +'" WHERE SOC_NOM="SO_PGFORVALOSALAIREPREV"');
  ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA="'+GetParamSocSecur ('SO_PGFORMETHODECALCPREV','')
          +'" WHERE SOC_NOM="SO_PGFORMETHODECALC"');

  if not IsDossierPCL then
  begin
    // MCD
    ExecuteSQLContOnExcept('UPDATE TACHE SET ATA_HEUREDEBUT = "'
      + usDateTime(iDate1900 + getparamsocSecur('SO_AFAMDEBUT','08:00:00'))
      + '",ATA_NUMEVENT=0');
    // JTR
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_INITQTECRE=0');
  end;
  // MD
  // Champs obsolètes à éliminer des listes (car encore présents physiquements)
  AglNettoieListes('DPANNUPERS',       '', nil, 'ANN_CODEPER;DOS_CODEPER;DOS_GROUPECONF;');
  AglNettoieListes('DPANNUPERS2',      '', nil, 'ANN_CODEPER;DOS_CODEPER;DOS_GROUPECONF;');
  AglNettoieListes('DPMULANU3',        '', nil, 'ANN_CODEPER;DOS_CODEPER;DOS_GROUPECONF;');
  AglNettoieListes('DPMULANU3LITE',    '', nil, 'ANN_CODEPER;DOS_CODEPER;DOS_GROUPECONF;');
  AglNettoieListes('YYDOSSIER_SEL',    '', nil, 'ANN_CODEPER;DOS_CODEPER;DOS_GROUPECONF;');
  AglNettoieListes('YYMULANNDOSS',     '', nil, 'ANN_CODEPER;DOS_CODEPER;DOS_GROUPECONF;');
  AglNettoieListes('YYMULANNDOSSLITE', '', nil, 'ANN_CODEPER;DOS_CODEPER;DOS_GROUPECONF;');
  AglNettoieListes('YYMULSELDOSS',     '', nil, 'ANN_CODEPER;DOS_CODEPER;DOS_GROUPECONF;');
  AglNettoieListes('YYMULSELDOSSLITE', '', nil, 'ANN_CODEPER;DOS_CODEPER;DOS_GROUPECONF;');
  // CD
  AglNettoieListes('RTMULACTREPORT', 'RAC_DUREEACTION',nil);
end;

//Joël TRIFILIEFF Demande N° 1414
{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 19/04/2007
Modifié le ... :
Description .. : Mise à jour des 2 nouveaux champs de la table ACOMPTES
                 dont un (GAC_NUMORDRE) fait parti de la clé primaire
*****************************************************************}
function UpdateAcpteCptaDiff : boolean ;
var TobAcomptesUpd, TobTmp : TOB;
    Cpt, Nbre : integer;
    LaRefPiece, Flux, NatureCpta, IsCptaDiff: string;
    Nature, Souche : string;
    NumP, Indice : integer;
    Qry : TQuery;
begin
  { Charge les acomptes en TOB}
  TobAcomptesUPD := TOB.Create('ACOMPTES', nil, -1);
  Qry := OpenSQL('SELECT ACOMPTES.*, PIECE.GP_VENTEACHAT AS FLUX FROM ACOMPTES '
               +' LEFT JOIN PIECE ON GP_NATUREPIECEG=GAC_NATUREPIECEG AND GP_SOUCHE=GAC_SOUCHE AND GP_NUMERO=GAC_NUMERO AND GP_INDICEG=GAC_INDICEG'
               //Joël TRIFILIEFF Demande N° 1650
               {+' WHERE GAC_NUMORDRE = 0', true);}
               +' WHERE ISNULL(GAC_NUMORDRE, 0) = 0', true);

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
    { Nouvelle pièce }
    if ((Flux <> TobTmp.GetString('FLUX'))
        or (Nature <> TobTmp.GetString('GAC_NATUREPIECEG'))
        or (Souche <> TobTmp.GetString('GAC_SOUCHE'))
        or (NumP <> TobTmp.GetInteger('GAC_NUMERO')) ) then
    begin
      Flux := TobTmp.GetString('FLUX');
      Nature := TobTmp.GetString('GAC_NATUREPIECEG');
      Souche := TobTmp.GetString('GAC_SOUCHE');
      NumP := TobTmp.GetInteger('GAC_NUMERO');
      Indice := TobTmp.GetInteger('GAC_INDICEG');
      Nbre := 1;
    end else
      Nbre := Nbre + 1;
    { Mise à jour de l'acompte }
    ExecuteSQL('UPDATE ACOMPTES SET GAC_NUMORDRE = ' + IntToStr(Nbre) + ', GAC_CPTADIFF = "-"' 
             + ' WHERE GAC_NATUREPIECEG = "' + Nature + '" AND GAC_SOUCHE = "' + Souche + '" AND GAC_NUMERO = ' + IntToStr(NumP)
             + ' AND GAC_INDICEG = ' + IntToStr(Indice) + ' AND GAC_JALECR = "' + TobTmp.GetString('GAC_JALECR') + '"'
             + ' AND GAC_NUMECR = ' + IntToStr(TobTmp.GetInteger('GAC_NUMECR')));
  end;
end;

Procedure MajVer835 ;
Begin
  // DW
  ExecuteSQLContOnExcept('update fpbiens set fpb_guidcode = pgiguid');
  ExecuteSQLContOnExcept('update fpbiens_ve set fve_guidcode = '
    +'(select fpb_guidcode from fpbiens where fpb_n01200=fve_ncode)');

  if not IsDossierPCL then
  begin

    //Marie-Noël GARNIER Le 29/05/2008 Version 8.1.850.86 Demande N° 2179
    {
    // MNG
    if (ExisteSQL('SELECT ##TOP 1## RPR_AUXILIAIRE FROM PROSPECTS')) then
      ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA = "X" WHERE SOC_NOM LIKE "SO_AVECCRM"');
    }

    // JLS
    ExecuteSQLContOnExcept('Update WordreRes Set WOR_TAUXUNITRES=0 WHERE WOR_TAUXUNITRES IS NULL');
    AGLNettoieListes('AFMULAVTAFFAIRE', ' AFF_TYPEAFFAIRE');
    ExecuteSQLContOnExcept('DELETE from LISTE where LI_LISTE IN ("WAFFLIGNES","WAFREVISIONBUDGET")');

    // Comportement des frais en génération de pièce avec ou sans regroupement
    // Annuler par Jean-Luc SAUZET Demande N° 1490
    {ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_MODEGROUPEPORT=iif(GPP_RECHTARIF501="X","REC","CHA") WHERE GPP_MODEGROUPEPORT IN ("CUM","FUS","INI")');}

    ExecuteSQLContOnExcept('UPDATE PORT SET GPO_MODEGROUPEPORT="" WHERE GPO_MODEGROUPEPORT IS NULL');
    ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE = "WORDRELAS"');

    // JTR
    UpdateAcpteCptaDiff();
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_INFOPIECEPREC = "", GPP_INFOPCEPRECCHX = ""');
    ExecuteSQLContOnExcept('UPDATE TIERSPIECE SET GTP_INFOPIECEPREC = "", GTP_INFOPCEPRECCHX = ""');

    //MCD
    InsertChoixCode('ATU', '008', 'Ressource commune', '', 'Assistant commun');

    // RH
    ExecuteSQLContOnExcept ('UPDATE HRCAISSE SET HRC_CLOTURESERVICE="-" WHERE HRC_CLOTURESERVICE IS NULL') ;
    ExecuteSQLContOnExcept ('UPDATE HRTYPEPREPA SET HTP_POINTINFO="" WHERE HTP_POINTINFO IS NULL') ;
    ExecuteSQLContOnExcept ('UPDATE JOURSCAISSE SET GJC_HRPERIODEDEBUT="" WHERE GJC_HRPERIODEDEBUT IS NULL');
    ExecuteSQLContOnExcept ('UPDATE MODEPAIE SET MP_GEREQTE="-" WHERE MP_GEREQTE IS NULL');
    ExecuteSQLContOnExcept ('UPDATE PIEDECHE SET GPE_DATEMODIF=GPE_DATECREATION, GPE_UTILISATEUR=GPE_CREATEUR, GPE_QUANTITE=0, GPE_VFACIALE=0 WHERE GPE_QUANTITE IS NULL');

  end;
end;


Procedure MajVer836 ;
Begin

  // ALR plante sur base dossier PCL
  // mais utilisation en multi dossier non pcl ???
  if not tablesurAutreBase('FILTREDONNEES') then
  // MB
		if not (ExisteSQL('SELECT 1 FROM FILTREDONNEES WHERE FTD_CODE="GROUPECONF"')) then
    begin
      ExecuteSQLContOnExcept( 'INSERT INTO FILTREDONNEES (FTD_CODE,FTD_TABLELIEN,FTD_TABLELIEE,FTD_CLETABLELIEN,FTD_CLETABLELIEE) '
                          + 'VALUES ("GROUPECONF","LIENDOSGRP","DOSSIER","LDO_NODOSSIER","DOS_NODOSSIER")' );
    end;
  //Laurent Abélard Le 11 Mai 2007
  //La conversion implicite de Chaîne de caractères en BLOB sous DB2 génére ce message d'erreur :
  //SQL0301 - Variable d'entrée PMD_SELECT ou argument 3 non admis
  //Je corrige en faisant un CAST explicite sachant que le CAST explicite ne fonctionne pas
  //sous ORACLE 8i ce qui interdit de modifier le ChangeSQL().
  If isDb2 Then
  Begin
    ExecuteSQLContOnExcept( 'INSERT INTO PARAMDONNEES (PMD_GRPNOM,PMD_NOMTABLE,PMD_SELECT,PMD_ORDERBY,PMD_LIBELLE) '
                          + 'VALUES ("GROUPECONF","DOSSIER",CAST( "DOS_NODOSSIER,DOS_LIBELLE" AS BLOB),"DOS_NODOSSIER ASC","Dossier")' );
    ExecuteSQLContOnExcept( 'INSERT INTO PARAMDONNEES (PMD_GRPNOM,PMD_NOMTABLE,PMD_SELECT,PMD_ORDERBY,PMD_LIBELLE) '
                          + 'VALUES ("GROUPECONF","UTILISAT",CAST( "US_UTILISATEUR,US_LIBELLE,US_ABREGE,US_FONCTION,US_GROUPE" AS BLOB),"US_UTILISATEUR ASC","Dossier")' );
  End
  Else
  Begin
		if not (ExisteSQL('SELECT 1 FROM PARAMDONNEES WHERE PMD_GRPNOM="GROUPECONF"')) then
    begin
      ExecuteSQLContOnExcept( 'INSERT INTO PARAMDONNEES (PMD_GRPNOM,PMD_NOMTABLE,PMD_SELECT,PMD_ORDERBY,PMD_LIBELLE) '
                            + 'VALUES ("GROUPECONF","DOSSIER","DOS_NODOSSIER,DOS_LIBELLE","DOS_NODOSSIER ASC","Dossier")' );
      ExecuteSQLContOnExcept( 'INSERT INTO PARAMDONNEES (PMD_GRPNOM,PMD_NOMTABLE,PMD_SELECT,PMD_ORDERBY,PMD_LIBELLE) '
                            + 'VALUES ("GROUPECONF","UTILISAT","US_UTILISATEUR,US_LIBELLE,US_ABREGE,US_FONCTION,US_GROUPE","US_UTILISATEUR ASC","Dossier")' );
    end;
  End;

  if not IsDossierPCL then
  begin
    // MNG
    AglNettoieListes('RTMULACTIONS', 'RAC_NUMCHAINAGE',nil);
    // JTR
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_TYPEPASSCPTA = "ZDI" WHERE GPP_TYPEPASSCPTA = "DIF"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_TYPEPASSCPTAR = "ZDI" WHERE GPP_TYPEPASSCPTAR = "DIF"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_TYPEPASSACC = "ZDI" WHERE GPP_TYPEPASSACC = "DIF"');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_TYPEPASSACCR = "ZDI" WHERE GPP_TYPEPASSACCR = "DIF"');
    ExecuteSQLContOnExcept('DELETE FROM COMMUN WHERE CO_TYPE = "GLC" AND CO_CODE = "DIF"');

    //Joël TRIFILIEFF Demande N° 1414
    {ExecuteSQLContOnExcept('UPDATE PIECE SET GP_REFCOMPTABLE = "ZDI" WHERE GP_REFCOMPTABLE = "DIFF"');
    ExecuteSQLContOnExcept('DELETE FROM COMPTADIFFEREE WHERE GCD_NATURECOMPTA IN ("FC", "FF", "AC", "AF")');
    UpdateAcpteCptaDiff2();}

    // GJ
    ExecuteSQLContOnExcept('UPDATE QWHISTORES SET QWH_REFERENCESAIS="" WHERE QWH_REFERENCESAIS IS NULL');
    // MCD
    AglNettoieListesPlus('AFMULFACTIERSAFF', 'AFF_PERIODICITE;AFF_METHECHEANCE;AFF_INTERVALGENER',nil,true);
    ExecuteSQLContOnExcept('update PIECE  set GP_PAYS="",GP_TYPEPAYS=""');

  end;

  if IsMonoOuCommune then
  begin
    // JTR
    ExecuteSQLContOnExcept ('UPDATE NETATCONTROLE SET NEC_EXTENSION=0');
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_INITQTECRE = 1 WHERE GPP_INITQTECRE = 0')
  end;
  // MVM
  ExecuteSQLContOnExcept('UPDATE immo SET i_reprisefec=0,i_reprisefeccedee=0,i_remplacee="",i_datecde="'
    + UsDateTime(iDate1900)+ '",i_dateliv="' + UsDateTime(iDate1900)
    +'", i_datejust="' + UsDateTime(iDate1900)+'",i_datederinv="'
    + UsDateTime(iDate1900) +'",i_derinv="",i_opeprime="-",i_opereducpri="-",i_opesbv="-",i_opereducsbv="'
    + UsDateTime(iDate1900) + '",i_operempl="-",i_remplace="",i_comptedep=""');

  ExecuteSQLContOnExcept ('UPDATE immo SET i_string1="",i_string2="",i_string3="",i_double1=0,i_double2=0,i_double3=0,i_version="",i_typer="",i_date1="' + UsDateTime(iDate1900) + '",i_date2="' + UsDateTime(iDate1900) + '",i_date3="' + UsDateTime(iDate1900) + '",'
  + 'i_opeard="-",i_ard="-",i_reparddeficit=0,i_reparddefcedee=0,i_dotmin=0,i_dotmax=0,i_dottheorique=0');

  ExecuteSQLContOnExcept('UPDATE immolog SET il_string1="", il_string2="", il_string3="", il_double1=0, il_double2=0, il_double3=0, il_date1="' + UsDateTime(iDate1900) + '", il_date2= "' +UsDateTime(iDate1900) + '", il_date3= "' + UsDateTime(iDate1900) +'"');
  ExecuteSQLContOnExcept('UPDATE immoamor SET ia_string1="", ia_string2="", ia_string3="", ia_double1=0, ia_double2=0, ia_double3=0, ia_date1="' + UsDateTime(iDate1900) +'", ia_date2="' + UsDateTime(iDate1900) +'", ia_date3="' + UsDateTime(iDate1900) + '", ia_cessionderog=0');

  // VG
  AglNettoieListes('PGELNAT', 'PEL_CONVENTION',nil);
  ExecuteSQLContOnExcept('UPDATE FRAISSALFORM SET PFS_TYPEPLANPREV='
    +'(SELECT PFO_TYPEPLANPREV FROM FORMATIONS '
    +' WHERE PFO_SALARIE=PFS_SALARIE AND PFO_CODESTAGE=PFS_CODESTAGE '
    +' AND PFO_ORDRE=PFS_ORDRE AND PFO_MILLESIME=PFS_MILLESIME)');

end;

Procedure MajVer837 ;
var i : integer;
Begin

  //Jean-Luc SAUZET Demande N° 1136
  If Not IsDossierPCL Then
  Begin
    ExecuteSQL('UPDATE CATALOGU SET GCA_DATEDEB="'+UsDateTime(iDate1900)+'", GCA_DATEFIN ="'+UsDateTime(iDate2099)+'" WHERE GCA_DATEDEB IS NULL');
  End;

  //Christophe AYEL Demande N° 1142
  AGLNettoieListes( 'CPEEXBQ', 'EE_DEVISE;' );
  AGLNettoieListes( 'CPEEXBQ2', 'EE_DEVISE;' );
  AGLNettoieListes( 'CPRUBRIQUEBUD', 'RB_PREDEFINI;' );

  //JUGDE Gérard Demande N° 1149
  For i := 1 To 3 Do
  Begin
    InsertChoixExt( 'Q2X', 'TQWH_DATELIBRE' + IntToStr(i), 'Date libre '     + IntToStr(i), 'Date libre '     + IntToStr(i), '' );
    InsertChoixExt( 'Q2W', 'TQWH_VALLIBRE'  + IntToStr(i), 'Valeur libre '   + IntToStr(i), 'Valeur libre '   + IntToStr(i), '' );
    InsertChoixExt( 'Q2Y', 'TQWH_BOOLLIBRE' + IntToStr(i), 'Décision libre ' + IntToStr(i), 'Décision libre ' + IntToStr(i), '' );
    InsertChoixExt( 'Q2V', 'TQWH_LIBREQWH'  + IntToStr(i), 'Table libre '    + IntToStr(i), 'Table libre '    + IntToStr(i), '' );
    InsertChoixExt( 'Q3X', 'TQWG_DATELIBRE' + IntToStr(i), 'Date libre '     + IntToStr(i), 'Date libre '     + IntToStr(i), '' );
    InsertChoixExt( 'Q3W', 'TQWG_VALLIBRE'  + IntToStr(i), 'Valeur libre '   + IntToStr(i), 'Valeur libre '   + IntToStr(i), '' );
    InsertChoixExt( 'Q3Y', 'TQWG_BOOLLIBRE' + IntToStr(i), 'Décision libre ' + IntToStr(i), 'Décision libre ' + IntToStr(i), '' );
    InsertChoixExt( 'Q3V', 'TQWG_LIBREQWG'  + IntToStr(i), 'Table libre '    + IntToStr(i), 'Table libre '    + IntToStr(i), '' );
  End;

  //JUGDE Gérard Demande N° 1151
  If Not IsDossierPCL Then
  Begin
    ExecuteSql( 'UPDATE QWRELEVELIG SET QWL_LIBREQWH1="",QWL_LIBREQWH2="",QWL_LIBREQWH3="",'
              + 'QWL_VALLIBRE1=0,QWL_VALLIBRE2=0,QWL_VALLIBRE3=0,'
              + 'QWL_DATELIBRE1="' + UsDateTime(iDate1900) + '",QWL_DATELIBRE2="' + UsDateTime(iDate1900) + '",'
              + 'QWL_DATELIBRE3="' + UsDateTime(iDate1900) + '",QWL_BOOLLIBRE1="-",QWL_BOOLLIBRE2="-",'
              + 'QWL_BOOLLIBRE3="-" '
              + 'WHERE QWL_LIBREQWH1 IS NULL' );
  End;

  //GALLIOT Vincent Demande N° 1153
  AGLNettoieListes( 'PGMULSALASSEDIC', 'PSA_ASSEDIC', Nil );

  //GALLIOT Vincent Demande N° 1154
  ExecuteSQL( 'UPDATE TAUXAT SET PAT_CODEELT=""' );
  ExecuteSQL( 'UPDATE RECAPSALARIES SET PRS_PRIVALENCOURS=0, PRS_PRIATTENCOURS=0, PRS_COMPTEUR1=0, PRS_COMPTEUR2=0, PRS_COMPTEUR3=0, PRS_COMPTEUR4=0' );

End;

Procedure MajVer838;
Begin

  //BICHERAY Bernard Demande N° 1181
  If Not IsDossierPCL Then
    ExecuteSQL( 'UPDATE CATALOGU SET GCA_CNS="-", GCA_RETOURHT=0, GCA_RETOURTTC=0' )

End;

Procedure MajVer839;
Begin

  //Marie-Christine DESSEIGNET Demande N° 1196
  SetParamSoc( 'SO_YPLSURBOOKING', afSurbooking );

  //Christophe AYEL Demande N° 1208
  AGLNettoieListes( 'CPDECVIRIN', 'E_CREDITDEV;E_DEBITDEV;' );
  AGLNettoieListes( 'CPDECVIRINEDNC', 'E_CREDITDEV;E_DEBITDEV;' );
  AGLNettoieListes( 'CPDECVIRINEDT', 'E_CREDITDEV;E_DEBITDEV;' );
  AGLNettoieListes( 'CPDECVIRINSV', 'E_CREDITDEV;E_DEBITDEV;' );
  AGLNettoieListes( 'CPDECVIRINEDNCSV', 'E_CREDITDEV;E_DEBITDEV;' );
  AGLNettoieListes( 'CPDECVIRINEDTSV', 'E_CREDITDEV;E_DEBITDEV;' );

  //Joel TRIFILIEFF Demande N° 1211
  //Suppresion de l'ancien paramétrage compta différée dans la tablette GCLIENCOMPTABLE
  ExecuteSQLContOnExcept( 'DELETE FROM COMMUN WHERE CO_TYPE="GLC" AND CO_CODE="DIF"' );

  If Not IsDossierPCL Then
  Begin

    //pour isofonctionnalité sur gestion ligne à 0 dans les pièces (suite modif GPAO)
    //Suite de la Demande N° 1158 corrigée par mail après concertation entre DESSEIGNET Marie-Christine et MERIEUX Gisèle
    //Correction dans la Demande N° 1196
    ExecuteSQLContOnExcept( 'UPDATE PARPIECE SET GPP_PRIXNULOK="X"' );

    //Joel TRIFILIEFF Demande N° 1213
    //eQualité GC n°13403
    ExecuteSQLContOnExcept( 'UPDATE PARPIECE SET GPP_CONTROLEMARGE="DEF" WHERE GPP_CONTROLEMARGE=""' );
  End;

End;

Procedure MajVer840;
Begin

  //Marie-Christine DESSEIGNET Demande N° 1227
  If Not IsDossierPCL Then
    ExecuteSQLContOnExcept( 'UPDATE AFFAIRE SET AFF_DATESIGNE = "' + UsDateTime(IDate1900) + '" where aff_datesigne < "' + UsDateTime(IDate1900) + '"' );

  InsertChoixCode( 'RSZ', 'CLZ', 'Table libre suspect 36 sur 6', '', '' );

  //Jean-Luc SAUZET Demande N° 1238
  //Mise à jour de EDIFAMILLEEMG
  ExecuteSQLNoPCL( 'UPDATE EDIFAMILLEEMG SET EFM_NATURESPIECE=EFM_NATURESPIECE||",DE" WHERE EFM_CODEFAMILLE="CDE"' );

  //Christophe AYEL Demande N° 1241
  AglNettoieListes( 'PGHONORDADS2', 'PDH_HONORAIRE' );

  //Christophe AYEL Demande N° 1243
  //Annulé par la demande N° 1318
  {
  If ( V_PGI.ModePCL <> '1' ) Then
  Begin
    SetParamSoc( 'SO_CPANAOUVRIR', True );
    SetParamSoc( 'SO_CPECHEOUVRIR', True );
  End ;
  }
  ExecuteSQLContOnExcept( 'UPDATE CREVCYCLE SET CCY_PREDEFINI = "CEG" WHERE CCY_PREDEFINI = "---"' );

  //Marie-Christine DESSEIGNET Demande N° 1245
  //GigA supprimé puis remis pour les mettre en visible à false
  AglNettoieListesPlus( 'AFTACHE_MUL', '', nil, false, 'ATA_ACTIVITEREPRIS' );
  AglNettoieListesPlus( 'AFTACHE_MUL', '', nil, false, 'ATA_IDENTLIGNE' );
  AglNettoieListesPlus( 'AFTACHE_MUL', 'ATA_TYPEPLANIF;ATA_ACTIVITEREPRIS;ATA_IDENTLIGNE', nil, false );

End;

Procedure MajVer841;
Var
  vSt    : String;
  vStDeb : String;
  vStFin : String;

Begin

  //Marie-Christine DESSEIGNET Demande N° 1284
  If ( isOracle ) Then
  Begin
    AGLNettoieListesPlus( 'GCMULLIGNEACH', '', nil, False, 'GL_NATUREPIECEG' );
    AGLNettoieListesPlus( 'GCMULLIGNEACH', 'GL_NATUREPIECEG', nil, False );
  End;

  //Marie-Christine DESSEIGNET Demande N° 1316
  If isMssql Then
  Begin
    DateTimeToString( vStDeb, 'hh:nn:ss', GetParamsocSecur('SO_AFPMDEBUT', '14:00:00') );
    DateTimeToString( vStFin, 'hh:nn:ss', GetParamsocSecur('SO_AFPMFIN', '18:00:00') );
    vStDeb := '"' + vStDeb + '"';
    vStFin := '"' + vStFin + '"';
  End
  Else If isOracle Then
  Begin
    DateTimeToString( vStDeb, 'hh', GetParamsocSecur('SO_AFPMDEBUT', '14:00:00') );
    DateTimeToString( vStFin, 'hh', GetParamsocSecur('SO_AFPMFIN', '18:00:00') );
    vStDeb := vStDeb + '/24';
    vStFin := vStFin + '/24';
  End
  Else If isDB2 Then
  Begin
    DateTimeToString( vStDeb, 'hh', GetParamsocSecur('SO_AFPMDEBUT', '14:00:00') );
    DateTimeToString( vStFin, 'hh', GetParamsocSecur('SO_AFPMFIN', '18:00:00') );
    vStDeb := vStDeb + ' hours ';
    vStFin := vStFin + ' hours ';
  End;

  vSt := 'UPDATE AFPLANNING SET apl_heurefin_pla = apl_datedebpla + ' + vStFin + ' where apl_heurefin_pla = apl_datedebpla + ' + vStDeb;
  ExecuteSQLNoPCL( vSt );
  vSt := 'UPDATE AFPLANNING SET apl_heurefin_real = apl_datedebreal + ' + vStFin + ' where apl_heurefin_real = apl_datedebreal + ' + vStDeb;
  ExecuteSQLNoPCL( vSt );
  If isMssql Then
  Begin
    DateTimeToString( vStDeb, 'hh:nn:ss', GetParamsocSecur('SO_AFAMDEBUT', '8:00:00') );
    DateTimeToString( vStFin, 'hh:nn:ss', GetParamsocSecur('SO_AFAMFIN', '12:00:00') );
    vStDeb := '"' + vStDeb + '"';
    vStFin := '"' + vStFin + '"';
  End
  Else If isOracle Then
  Begin
    DateTimeToString( vStDeb, 'hh', GetParamsocSecur('SO_AFPMDEBUT', '8:00:00') );
    DateTimeToString( vStFin, 'hh', GetParamsocSecur('SO_AFPMFIN', '12:00:00') );
    vStDeb := vStDeb + '/24';
    vStFin := vStFin + '/24';
  End
  Else If isDB2 Then
  Begin
    DateTimeToString( vStDeb, 'hh', GetParamsocSecur('SO_AFPMDEBUT', '8:00:00') );
    DateTimeToString( vStFin, 'hh', GetParamsocSecur('SO_AFPMFIN', '12:00:00') );
    vStDeb := vStDeb + ' hours ';
    vStFin := vStFin + ' hours ';
  End;

  ExecuteSQLNoPCL( 'UPDATE AFPLANNING SET apl_heurefin_pla = apl_datedebpla + ' + vStFin + 'where apl_heurefin_pla = apl_datedebpla + ' + vStDeb );
  ExecuteSQLNoPCL( 'UPDATE AFPLANNING SET apl_heurefin_real = apl_datedebreal + ' + vStFin + 'where apl_heurefin_real = apl_datedebreal + ' + vStDeb );

  //Vincent GALLIOT Demande N° 1324
  ExecuteSQLContOnExcept( 'UPDATE CHOIXCOD SET CC_LIBRE="SC1" WHERE CC_TYPE="PSC" AND CC_LIBRE="001"' );
  ExecuteSQLContOnExcept( 'UPDATE CHOIXCOD SET CC_LIBRE="SC2" WHERE CC_TYPE="PSC" AND CC_LIBRE="002"' );
  ExecuteSQLContOnExcept( 'UPDATE CHOIXCOD SET CC_LIBRE="SC3" WHERE CC_TYPE="PSC" AND CC_LIBRE="003"' );
  ExecuteSQLContOnExcept( 'UPDATE CHOIXCOD SET CC_LIBRE="SC4" WHERE CC_TYPE="PSC" AND CC_LIBRE="004"' );
  ExecuteSQLContOnExcept( 'UPDATE CHOIXCOD SET CC_LIBRE="SC5" WHERE CC_TYPE="PSC" AND CC_LIBRE="005"' );

End;

Procedure MajVer842;
Begin

  //Joël TRIFILIEFF Demande N° 1325
  ExecuteSQLContOnExcept( 'UPDATE DECHAMPS SET DH_CONTROLE = "LD" WHERE DH_NOMCHAMP="GP_ECHEBLOQUE"' );

  //Vincent GALLIOT Demande N° 1373
  AGLNettoieListes( 'PGSESSIONSTAGE','',Nil,'ANN_NOM1' );
  AGLNettoieListes( 'PGHISTORETENUE','',Nil,'PHR_TYPERGT' );
  AGLNettoieListes( 'PGTYPERETENUE','',Nil,'PTR_TYPERETENUE' );

  //Marie-Christine DESSEIGNET Demande N° 1374
  ExecuteSQLContOnExcept( 'UPDATE CHOIXCOD SET CC_LIBRE="de missions" WHERE CC_TYPE="atu" AND CC_CODE="003"' );

  //Christophe AYEL Demande N° 1378
  ExecuteSQLContOnExcept( 'UPDATE BANQUECP SET BQ_NATURECPTE = "BQE" WHERE BQ_NATURECPTE <> "COU" OR BQ_NATURECPTE IS NULL' );

End;

Procedure MajVer843;
Begin

  //Jean-Luc SAUZET Demande N° 1391
  //Cette requête annule et remplace la 1er requête envoyée lors de la création anticipée de la structure afin le figeage du MCD
  If Not IsDossierPCL Then
  Begin
    ExecuteSQLContOnExcept( 'UPDATE WORDREBES SET WOB_QUIFOURNI="STR" WHERE WOB_QUIFOURNI IS NULL AND WOB_MODECONSO="NON"'
                          + ' AND ISNULL((SELECT WOL_TYPEORDRE FROM WORDRELIG WHERE WOL_NATURETRAVAIL=WOB_NATURETRAVAIL AND WOL_LIGNEORDRE=WOB_LIGNEORDRE),"")="STA"' );

    //Joël TRIFILIEFF Demande N° 1414
    ExecuteSQLContOnExcept( 'UPDATE PARPIECECOMPL SET GPC_TYPEPASSCPTA="ZDI" WHERE GPC_TYPEPASSCPTA = "DIF"' );
    ExecuteSQLContOnExcept( 'UPDATE PARPIECECOMPL SET GPC_TYPEPASSCPTAR="ZDI" WHERE GPC_TYPEPASSCPTAR = "DIF"' );
    ExecuteSQLContOnExcept( 'UPDATE PARPIECECOMPL SET GPC_TYPEPASSACC="ZDI" WHERE GPC_TYPEPASSACC = "DIF"' );
    ExecuteSQLContOnExcept( 'UPDATE PARPIECECOMPL SET GPC_TYPEPASSACCR="ZDI" WHERE GPC_TYPEPASSACCR = "DIF"' );

    //Dominique BROSSET Demande N° 1385
    ExecuteSQLContOnExcept( 'UPDATE ARTICLECOMPL SET GA2_MONTANTTX1=0, GA2_MONTANTTX3=0, GA2_MONTANTTX4=0, GA2_MONTANTTX5=0, GA2_MODELETAXE=""' );
    UpdateDecoupePiece( 'GP_REGIMETAXE2="", GP_REGIMETAXE3="", GP_REGIMETAXE4="", GP_REGIMETAXE5=""' );
    UpdateDecoupePiece(' GP_COMMISSIONR=0, GP_COMMISSIONR2=0, GP_COMMISSIONR3=0' );
    {ExecuteSQLContOnExcept( 'UPDATE LIGNECOMPL SET GLC_MONTANTTX1=0, GLC_MONTANTTX3=0, GLC_MONTANTTX4=0, GLC_MONTANTTX5=0, GLC_MODELETAXE=""' );}

  End;

  //Dominique BROSSET Demande N° 1385
  ExecuteSQLContOnExcept( 'UPDATE TIERSCOMPL SET YTC_REGIMETX2="", YTC_REGIMETX3="", YTC_REGIMETX4="", YTC_REGIMETX5="", YTC_MODELETAXE=""' );

  //Ludovic MONTAVON Demande N° 1418
  BasculeExpertEtConseil;

  //Vincent GALLIOT Demande N° 1423
  if not (ExisteSQL('SELECT 1 FROM YDATATYPELINKS WHERE YDL_CODEHDTLINK="PAIEGED12"')) then
  begin
  	ExecuteSQLContOnExcept( 'INSERT INTO YDATATYPELINKS (YDL_CODEHDTLINK, YDL_LIBELLE, YDL_MDATATYPE, YDL_SDATATYPE, YDL_TYPELINK, YDL_PREDEFINI, YDL_LCOMMUNE ) VALUES ("PAIEGED12","Liaison des tables libres PGLIBGED 1 - 2", "PGLIBGED1", "PGLIBGED2", "NOR", "DOS", "-" )' );
  END;
  if not (ExisteSQL('SELECT 1 FROM YDATATYPELINKS WHERE YDL_CODEHDTLINK="PAIEGED23"')) then
  begin
  	ExecuteSQLContOnExcept( 'INSERT INTO YDATATYPELINKS (YDL_CODEHDTLINK, YDL_LIBELLE, YDL_MDATATYPE, YDL_SDATATYPE, YDL_TYPELINK, YDL_PREDEFINI, YDL_LCOMMUNE ) VALUES ("PAIEGED23","Liaison des tables libres PGLIBGED 2 - 3", "PGLIBGED2", "PGLIBGED3", "NOR", "DOS", "-" )' );
  end;

End;

Procedure MajVer844;
Begin

  If Not IsDossierPCL Then
  Begin

    //Dominique BROSSET Demande N° 1436
    {
      Laurent Abélard le 27 Septembre 2007, Il apparaît que plusieurs UPDATE sur LIGNECOMPL sont très long
      et nécessitent de regrouper tout ces UPDATE en un seul (Voir en Fin de MajVer822()).
    ExecuteSQLContOnExcept( 'UPDATE LIGNECOMPL SET GLC_MONTANTTX1=0, GLC_MONTANTTX3=0, GLC_MONTANTTX4=0, GLC_MONTANTTX5=0' );
    }

    //Joël TRIFILIEFF Demande N° 1480
    ExecuteSQLContOnExcept( 'UPDATE TIERSIMPACTPIECE SET GTI_TYPENATTIERS = "FAC" WHERE GTI_ELEMENTFC = "TXN" AND GTI_TYPENATTIERS = ""' )

  End;

  //Jean-Luc SAUZET Demande N° 1443
  AGLNettoieListesPlus( 'WPF_AUTOWOL', 'WPF_CODITI ;WPF_CIRCUIT', nil, True );
  { Mise à jour de MENU }
  ExecuteSQLNoPCL( 'DELETE FROM MENU WHERE MN_1=215 AND MN_2=7 AND MN_3=11 AND MN_4>0' );
  ExecuteSQLContOnExcept( 'DELETE FROM PARAMSOC WHERE SOC_NOM = "SCO_PDRGENEREWPL"' );

  //Christophe AYEL Demande N° 1461
  ExecuteSQLContOnExcept( 'DELETE FROM PARAMSOC WHERE SOC_NOM = "SO_CPANAOUVRIR"' );
  ExecuteSQLContOnExcept( 'DELETE FROM PARAMSOC WHERE SOC_NOM = "SO_CPECHEOUVRIR"' );


End;

Procedure MajVer845;
Begin

  If Not IsDossierPCL Then
  Begin

    //Jean-Luc SAUZET Demande N° 1490
    // Comportement des frais en génération de pièce avec ou sans regroupement
    {ExecuteSQLContOnExcept( 'UPDATE PARPIECE SET GPP_MODEGROUPEPORT=iif(GPP_MODEGROUPEPORT="FUS","REG",iif(GPP_RECHTARIF501="X","REC","CHA")) WHERE GPP_MODEGROUPEPORT IN ("CUM", "FUS", "INI")' );}

    //Jean-Luc SAUZET Demande N° 1554
    ExecuteSQLContOnExcept( 'UPDATE PARPIECE SET GPP_MODEGROUPEPORT=iif(GPP_RECHTARIF501="X","REC","CHA") WHERE GPP_MODEGROUPEPORT IN ("CUM","FUS","INI")');

    //Joël SICH Demande N° 1489
    ExecuteSQLContOnExcept( 'UPDATE PARPIECE SET GPP_PRIXNULOK="X" WHERE GPP_NATUREPIECEG="FRA"' );

  End;

  //Joël TRIFILIEFF Demande N° 1491
  ExecuteSQLContOnExcept( 'UPDATE DECHAMPS SET DH_CONTROLE = "LD" WHERE DH_NOMCHAMP = "GP_ECHEBLOQUE"' );

  //Claude DUMAS Demande N° 1501
  AglNettoieListes( 'RTMULLIGNMAILING','T_JURIDIQUE;T_PRENOM;C_NOM;C_PRENOM;C_CIVILITE;'
                  + 'T_ADRESSE1;T_ADRESSE2;T_ADRESSE3;T_PAYS;T_LIBELLE;T_CODEPOSTAL;T_VILLE', Nil );

  AglNettoieListes( 'RTAFFMAILINGCON','T_JURIDIQUE;T_PRENOM;C_NOM;C_PRENOM;C_CIVILITE;'
                  + 'T_ADRESSE1;T_ADRESSE2;T_ADRESSE3;T_PAYS;T_LIBELLE;T_CODEPOSTAL;T_VILLE', Nil );

  AglNettoieListes( 'RTMULCONTMAILING','T_JURIDIQUE;T_PRENOM;C_NOM;C_PRENOM;C_CIVILITE;'
                  + 'T_ADRESSE1;T_ADRESSE2;T_ADRESSE3;T_PAYS;T_LIBELLE;T_CODEPOSTAL;T_VILLE', Nil );

  AglNettoieListes( 'RTMULTIERSMAILING','T_JURIDIQUE;T_PRENOM;C_NOM;C_PRENOM;C_CIVILITE;'
                  + 'T_ADRESSE1;T_ADRESSE2;T_ADRESSE3;T_PAYS;T_LIBELLE;T_CODEPOSTAL;T_VILLE', Nil );

  AglNettoieListes( 'RTSUSPMAILING','RSU_JURIDIQUE;RSU_CONTACTNOM;RSU_CONTACTPRENOM;RSU_CONTACTCIVILITE;'
                  + 'RSU_ADRESSE1;RSU_ADRESSE2;RSU_ADRESSE3;RSU_PAYS;RSU_LIBELLE;RSU_CODEPOSTAL;RSU_VILLE', Nil );

  //Joël SICH Demande N° 1509
  AGLNettoieListes( 'GCADRESSES','ADR_TYPEADRESSE;ADR_NATUREAUXI;ADR_REFCODE', Nil );

  //Vincent GALLIOT Demande N° 1519
  AglNettoieListes( 'PGTYPERETENUE', 'PTR_RUBRIQUE', Nil );

End;

Procedure MajVer846;
Begin

  //Christophe AYEL Demande N° 1520
  AglNettoieListes( 'CPDEVALIDEECR', 'E_QUALIFORIGINE;' );
  AglNettoieListes( 'CPVALIDEECR', 'E_QUALIFORIGINE;' );
  AglNettoieListes( 'CPCONSECR', 'E_QUALIFORIGINE;E_REFGESCOM;' );
  AglNettoieListes( 'REIMPUTATION', 'E_QUALIFORIGINE;' );

  //Christophe AYEL Demande N° 1524
  AglNettoieListes( 'RELEVECOMPTE', 'E_MODEPAIE;' );

  //Pascal BASSET Demande N° 1566
  //MAJ CONCEPT DECLARATION FISCALE : PB le 10/07/2007
  ExecuteSQLContOnExcept( 'UPDATE MENU SET MN_ACCESGRP = ' +
                          '"0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"' +
                          ',MN_VERSIONDEV="-" WHERE MN_1=26 AND MN_2=6' );

  //Vincent GALLIOT Demande N° 1569
  ExecuteSQLContOnExcept( 'UPDATE DETABLES SET DT_COMMUN="S" WHERE DT_NOMTABLE="ELTNIVEAUREQUIS"' );
  ExecuteSQLContOnExcept( 'UPDATE DECHAMPS SET DH_LIBELLE="Population" WHERE DH_PREFIXE="PPU" AND DH_NOMCHAMP="PPU_USEREDIT"' );
  AglNettoieListes( 'PGPARAMSALARIE', 'PPP_PGTYPEDONNE', nil );

  //Vincent GALLIOT Demande N° 1583
  {AglNettoieListes( 'PGEMPLOIINTERIM', 'PEI_ORDRE', nil );}

  //Christophe AYEL Demande N° 1575
  If V_PGI.ModePCL='1' Then
  Begin
    ExecuteSQLContOnExcept( 'UPDATE MENU SET MN_ACCESGRP = ' +
                            '"0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"' +
                            ',MN_VERSIONDEV="-" WHERE MN_TAG=9105 AND MN_1=52' );

    ExecuteSQLContOnExcept( 'UPDATE MENU SET MN_ACCESGRP = ' +
                            '"0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"' +
                            ',MN_VERSIONDEV="-" WHERE MN_TAG=2552 AND MN_1=2' );
  End;

End;

//Laurent Abélard le 16 Juillet 2007, Je saute volontairement de la 846 à 850 pour la dernière SOCREF.
Procedure MajVer850;
Begin
	// un truc bizare...autant qu'etrange
  ExecuteSQL('UPDATE CODECPTA SET GCP_FAMILLETAXE="" WHERE GCP_FAMILLETAXE IS NULL');

  //Etienne VIENOT Demande N° 1637
  ExecuteSQLContOnExcept( 'UPDATE DECHAMPS SET DH_LIBELLE="Membre de l''UE" WHERE  DH_PREFIXE="PY" AND DH_NOMCHAMP="PY_MEMBRECEE"' );

  //Monique FAUDEL Demande N° 1648
  AglNettoieListesPlus( 'PGENVOISOCIAL', 'PES_GUID1', nil, False );
  AglNettoieListesPlus( 'PGENVOITDB', 'PES_GUID1', nil, False );

  //Joël TRIFILIEFF Demande N° 1657
  If Not IsDossierPCL Then
  Begin
    ExecuteSQLContOnExcept( 'UPDATE PARPIECE SET GPP_TYPEPASSCPTA="AUC", GPP_TYPEPASSCPTAR ="AUC", GPP_TYPEPASSACC="AUC", GPP_TYPEPASSACCR="AUC" WHERE GPP_TYPEECRCPTA = "RIE"' );
    ExecuteSQLContOnExcept( 'UPDATE PARPIECECOMPL SET GPC_TYPEPASSCPTA="AUC", GPC_TYPEPASSCPTAR ="AUC", GPC_TYPEPASSACC="AUC", GPC_TYPEPASSACCR="AUC" WHERE GPC_NATUREPIECEG = (SELECT GPP_NATUREPIECEG FROM PARPIECE WHERE GPP_TYPEECRCPTA = "RIE" AND GPP_NATUREPIECEG=GPC_NATUREPIECEG)' );
  End;

  //Christophe AYEL Demande N°1701
  //À la demande de Olivier Lambert, voici 2 modifications à intégrer dans le MajApres
  //pour résoudre le problème de la FQ 14123 :

  ExecuteSQLContOnExcept( 'UPDATE MENU SET MN_ACCESGRP = ' +
                          '"0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"' +
                          ',MN_VERSIONDEV="-" WHERE MN_TAG=20233 AND MN_1=2' );
  ExecuteSQLContOnExcept( 'UPDATE MENU SET MN_ACCESGRP = ' +
                          '"0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"' +
                          ',MN_VERSIONDEV="-" WHERE MN_TAG=20241 AND MN_1=2' );

  //Christophe AYEL Demande N°1784
  ExecuteSQLContOnExcept( 'UPDATE MENU SET MN_VERSIONDEV = "-" WHERE MN_1 = 26 AND MN_2 = 1' );

  //Pierre LENORMAND Demande N°1812
  //CO + PL : correction de demande suivi de socref N° 900
  ExecuteSQLContOnExcept( 'UPDATE DESHARE SET DS_NOMTABLE="YPARAMEDITION" WHERE DS_NOMTABLE="YHISTOPRINT"' );

  //Pascal BASSET Demande N°1667
  //MAJ MENU SCI,SCM,IFU,TVA et BIF des droits
  ExecuteSQLContOnExcept( 'UPDATE MENU SET MN_ACCESGRP = ' +
                          '"0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"' +
                          ',MN_VERSIONDEV="-" WHERE MN_1 IN (172,173,174,175,176)' );

  //Dominique BROSSET Demande N°1713
  If Not IsDossierPCL Then
    ExecuteSQLContOnExcept ('UPDATE PARPIECE SET GPP_VENTEACHAT="STO" WHERE GPP_NATUREPIECEG IN ("EEX", "SEX")' );

  //Joël SICH Demande N° 1745
  ExecuteSQLContOnExcept('UPDATE COMMUN SET CO_LIBELLE="GCGRILLELIBREDIM1" WHERE CO_TYPE="YEB" AND CO_CODE="G94" AND CO_LIBELLE="GCLIBREGRILLEDIM1"' );
  ExecuteSQLContOnExcept('UPDATE COMMUN SET CO_LIBELLE="GCGRILLELIBREDIM2" WHERE CO_TYPE="YEB" AND CO_CODE="G95" AND CO_LIBELLE="GCLIBREGRILLEDIM2"' );
  ExecuteSQLContOnExcept('UPDATE COMMUN SET CO_LIBELLE="GCGRILLELIBREDIM3" WHERE CO_TYPE="YEB" AND CO_CODE="G96" AND CO_LIBELLE="GCLIBREGRILLEDIM3"' );
  ExecuteSQLContOnExcept('UPDATE COMMUN SET CO_LIBELLE="GCGRILLELIBREDIM4" WHERE CO_TYPE="YEB" AND CO_CODE="G97" AND CO_LIBELLE="GCLIBREGRILLEDIM4"' );
  ExecuteSQLContOnExcept('UPDATE COMMUN SET CO_LIBELLE="GCGRILLELIBREDIM5" WHERE CO_TYPE="YEB" AND CO_CODE="G98" AND CO_LIBELLE="GCLIBREGRILLEDIM5"' );

  If EstBaseMultiSoc And TablePartagee ('DIMENSION') Then
  Begin
    ExecuteSQLNoPCL('UPDATE DESHARE SET DS_NOMTABLE="GCGRILLELIBREDIM1" WHERE DS_NOMTABLE="GCLIBREGRILLEDIM1"' );
    ExecuteSQLNoPCL('UPDATE DESHARE SET DS_NOMTABLE="GCGRILLELIBREDIM2" WHERE DS_NOMTABLE="GCLIBREGRILLEDIM2"' );
    ExecuteSQLNoPCL('UPDATE DESHARE SET DS_NOMTABLE="GCGRILLELIBREDIM3" WHERE DS_NOMTABLE="GCLIBREGRILLEDIM3"' );
    ExecuteSQLNoPCL('UPDATE DESHARE SET DS_NOMTABLE="GCGRILLELIBREDIM4" WHERE DS_NOMTABLE="GCLIBREGRILLEDIM4"' );
    ExecuteSQLNoPCL('UPDATE DESHARE SET DS_NOMTABLE="GCGRILLELIBREDIM5" WHERE DS_NOMTABLE="GCLIBREGRILLEDIM5"' );
  End;

{ GERARD PIOT DEMANDE EQUALITE 2501 28052008
  //Chritophe AYEL Demande N° 1837
  ExecuteSQLContOnExcept( 'UPDATE EEXBQLIG SET CEL_DEVISE = (SELECT MAX(EE_DEVISE) FROM EEXBQ ' +
                          'WHERE CEL_REFPOINTAGE = EE_REFPOINTAGE ' +
                          'AND CEL_NUMRELEVE = EE_NUMERO) ' +
                          'WHERE (CEL_DEVISE = "" OR CEL_DEVISE IS NULL)' );
  ExecuteSQLContOnExcept( 'UPDATE EEXBQLIG SET CEL_REFPOINTAGE = "" WHERE (CEL_DATEPOINTAGE = "' + UsDateTime(iDate1900) + '" OR CEL_DATEPOINTAGE IS NULL)' );}

  //Marie-Christine DESSEIGNET Le 25/10/2007 Version 8.1.850.14 Demande N° 1881
  AglNettoieListes( 'AFPLALIGNE_MUL', 'GLC_AFFAIRELIEE;ATA_QTEPLANIFPLA', Nil, 'GLC_AFFAIRELIE' );

  //Claude DUMAS Le 20/11/2007 Version 8.1.850.15 Demande N° 1963
  ExecuteSQLContOnExcept( 'UPDATE MENU SET MN_LIBELLE = "Libellés tables libres propositions sur 6 caractères"' +
                          ' WHERE MN_1 = 65 AND MN_2 = 12 AND MN_3 = 39 AND MN_4 = 0' );

  ExecuteSQLContOnExcept( 'UPDATE MENU SET MN_LIBELLE = "Tables libres propositions sur 6 caractères"' +
                          ' WHERE MN_1 = 65 AND MN_2 = 12 AND MN_3 = 40 AND MN_4 = 0' );

  //Claude DUMAS Le 20/11/2007 Version 8.1.850.15 Demande N° 1964
  //RAB des contrôles du champ RAC_INTERVINT :
  ExecuteSQLContOnExcept( 'UPDATE DECHAMPS SET DH_CONTROLE = "" WHERE DH_NOMCHAMP = "RAC_INTERVINT"' );

  //Jean-Paul LAURENT Le 20/11/2007 Version 8.1.850.15 Demande N° 1998
  //Changement de libellés de menus ( Menu 65 - Paramètres Organisation - Recherche par Responsable)
  ExecuteSQLContOnExcept( 'UPDATE MENU SET MN_LIBELLE="Recherche par responsable" Where Mn_1=65 and MN_2=11 and MN_3=4 and MN_4=0' );

  //Jean-Paul LAURENT Le 20/11/2007 Version 8.1.850.15 Demande N° 2002
  ExecuteSQLContOnExcept( 'UPDATE DECHAMPS SET DH_CONTROLE="LDCZ2" WHERE DH_PREFIXE="T" AND DH_NUMCHAMP = 280 AND DH_NOMCHAMP = "T_ENSEIGNE"' );

  //Jean-Paul LAURENT Le 21/11/2007 Version 8.1.850.15 Demande N° 2011
  ExecuteSQLContOnExcept( 'UPDATE PARAMSOC SET SOC_DATA = "0" WHERE SOC_NOM = "SO_NBJENCOURS" AND SOC_DATA = "-20"' );

  //Nicolas FOURNEL Le 11/12/2007 Version 8.1.850.17 Demande N° 2048
  ExecuteSQLContOnExcept( 'UPDATE PARAMSOC SET SOC_DESIGN = "B;;Module Business Side visible;204;211;363;232;X;-;5;325;15;346;MS Sans Serif;-2147483640;8;0;;0;0;0;X;0;0" WHERE SOC_NOM = "SO_MODULEPGISIDE"' );
  ExecuteSQLContOnExcept( 'UPDATE PARAMSOC SET SOC_DESIGN = "B;;Traduction du libellé de l''article si fournisseur étranger;16;217;289;232;X;-;10;260;95;281;MS Sans Serif;-16777208;8;0;;0;0;0;X;0;0" WHERE SOC_NOM = "SO_TRADUCCONTREMARQUE"' );

  //Marie-Christine DESSEIGNET Le 21/01/2008 Version 8.1.850.18 Demande N° 2069
  ModifPlanningGA();

  //Marie-Noël GARNIER Le 21/01/2008 Version 8.1.850.18 Demande N° 2102
  ExecuteSQLContOnExcept( 'UPDATE DECHAMPS SET DH_CONTROLE = "LDC" WHERE DH_PREFIXE = "WPN" AND DH_NOMCHAMP in ("WPN_AFFAIRE","WPN_ARTICLEPARC",'
                        + '"WPN_ETATPARC","WPN_IDENTPARC","WPN_NATUREPIECEG","WPN_REFLIGNEAFF") ' );
  { wpc_identifiant }
  ExecuteSQLContOnExcept( 'UPDATE DECHAMPS SET DH_CONTROLE = "LDV" WHERE DH_PREFIXE = "WPC" AND DH_NUMCHAMP=1 ' );

  { pour ajouter les 'Z' de la modif en série }
  ExecuteSQLContOnExcept( 'UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Z" WHERE DH_PREFIXE = "WPC" AND ( DH_NOMCHAMP = "WPC_LIBELLE" OR'
                        + ' DH_NOMCHAMP like "%LIBRE%" ) AND DH_CONTROLE NOT LIKE "%Z%"' );

  //Marie-Noël GARNIER Le 21/01/2008 Version 8.1.850.18 Demande N° 2113
  { modification des libellés booléen en décision dans les tablettes infos complémentaire }
  ExecuteSQLContOnExcept( 'UPDATE CHOIXCOD SET CC_LIBELLE = "Décision libre " || SubString(CC_LIBELLE,15,LEN(CC_LIBELLE)) WHERE CC_TYPE LIKE "R%Z" AND CC_LIBELLE LIKE "booléen libre%" ' );

  //Claude DUMAS Le 21/01/2008 Version 8.1.850.18 Demande N° 2120
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Libellé du critère d''évaluation" WHERE DH_PREFIXE = "RQC" AND DH_NOMCHAMP = "RQC_LIBELLE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Nombre d''évaluations enregistrées" WHERE dh_prefixe = "RQT" and DH_NOMCHAMP = "RQT_NBREVAL"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Table libre 26 sur 6" WHERE DH_PREFIXE = "RSC" AND DH_NOMCHAMP = "RSC_6RSCLIBTABLE26"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Table libre 27 sur 6" WHERE DH_PREFIXE = "RSC" AND DH_NOMCHAMP = "RSC_6RSCLIBTABLE27"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Table libre 28 sur 6" WHERE DH_PREFIXE = "RSC" AND DH_NOMCHAMP = "RSC_6RSCLIBTABLE28"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Table libre 29 sur 6" WHERE DH_PREFIXE = "RSC" AND DH_NOMCHAMP = "RSC_6RSCLIBTABLE29"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Table libre 30 sur 6" WHERE DH_PREFIXE = "RSC" AND DH_NOMCHAMP = "RSC_6RSCLIBTABLE30"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Table libre 31 sur 6" WHERE DH_PREFIXE = "RSC" AND DH_NOMCHAMP = "RSC_6RSCLIBTABLE31"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Table libre 32 sur 6" WHERE DH_PREFIXE = "RSC" AND DH_NOMCHAMP = "RSC_6RSCLIBTABLE32"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Table libre 33 sur 6" WHERE DH_PREFIXE = "RSC" AND DH_NOMCHAMP = "RSC_6RSCLIBTABLE33"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Table libre 34 sur 6" WHERE DH_PREFIXE = "RSC" AND DH_NOMCHAMP = "RSC_6RSCLIBTABLE34"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Table libre 35 sur 6" WHERE DH_PREFIXE = "RSC" AND DH_NOMCHAMP = "RSC_6RSCLIBTABLE35"');

  //Jean-Luc SAUZET Le 21/01/2008 Version 8.1.850.18 Demande N° 2121
  AGLNettoieListesPlus( 'WORDRELIG2', 'WOL_QDEMSAIS', Nil );

  //Joël TRIFILIEFF Le 29/05/2008 Version 8.1.850.86 Demande N° 2154
  { JTR le 06/02/2008, Ajout champ dans la liste }
  AglNettoieListesPlus( 'GCGROUPEPIECEACH', 'GP_ECHEBLOQUE', Nil, True );

  //Marie-Noël GARNIER Le 29/05/2008 Version 8.1.850.86 Demande N° 2179
  { maj paramètre crm accompagnement }
  If Not IsDossierPCL Then
    If Not (ExisteSQL('SELECT ##TOP 1## RPR_AUXILIAIRE FROM PROSPECTS')) Then
          ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA = "-" WHERE SOC_NOM LIKE "SO_CRMACCOMPAGNEMENT"' );

  //Claude DUMAS Le 29/05/2008 Version 8.1.850.86 Demande N° 2347
  ExecuteSQLContOnExcept( 'UPDATE DETABLES SET DT_LIBELLE = "Liaison critères et tiers" WHERE DT_PREFIXE = "RQL" AND DT_LIBELLE = "Liaison criteres et tiers"' );

  //Dominique BROSSET Le 29/05/2008 Version 8.1.850.86 Demande N° 2410
  ExecuteSQLContOnExcept( 'UPDATE DECHAMPS SET DH_CONTROLE = "LDZ1234" WHERE DH_NOMCHAMP = "T_LOCALTAX"' );

  //Jean-Luc SAUZET Le 29/05/2008 Version 8.1.850.86 Demande N° 2490
  //JS1 8.1.850.86 ON MET UN FROM CAR LE DELETE SEUL NE PASSE PAS SOUS DB2 !!
  ExecuteSQLContOnExcept( 'DELETE FROM DECOMBOS WHERE DO_COMBO = "TTTARIFALLTIERS"' );

  //Jean-Luc SAUZET Le 29/05/2008 Version 8.1.850.86 Demande N° 2554
  //JS1 8.1.850.86 ON MET UN FROM CAR LE DELETE SEUL NE PASSE PAS SOUS DB2 !!
  ExecuteSQLContOnExcept( 'DELETE FROM MENU WHERE MN_1=160 AND MN_4=1 AND MN_TAG=160100 AND EXISTS (SELECT 1 FROM MENU WHERE MN_1=160 AND MN_4=0 AND MN_TAG=160100)' );
  ExecuteSQLContOnExcept( 'DELETE FROM MENU WHERE MN_1=160 AND MN_4=1 AND MN_TAG=160101 AND EXISTS (SELECT 1 FROM MENU WHERE MN_1=160 AND MN_4=0 AND MN_TAG=160101)' );

  //Gisèle MERIEUX Le 29/05/2008 Version 8.1.850.86 Demande N° 2587
  If Not IsDossierPCL Then
    ExecuteSQLContOnExcept( 'UPDATE AFPLANNING SET APL_ESTFACTURE="X" WHERE APL_ETATLIGNE="FAC"' );

{JS1 18/06/2008 on commente les deux requêtes de la demande 2142
  //Dominique BROSSET Le 29/05/2008 Version 8.1.850.86 Demande N° 2142
  ExecuteSQLContOnExcept( 'UPDATE TIERS SET T_FACTURE=T_AUXILIAIRE WHERE T_FACTURE=""' );
  ExecuteSQLContOnExcept( 'UPDATE TIERS SET T_PAYEUR=T_AUXILIAIRE WHERE T_PAYEUR=""' );}

  //Dominique BROSSET Le 30/05/2008 Version 8.1.850.86 : on supprime la ##top1## qui ne passe pas sous oracle (dans la sous requete update) et
  //on rajoute des AND T_AUXILIAIRE = YTC_AUXILIAIRE pour garantir l'unicité
  ExecuteSQLContOnExcept( 'UPDATE TIERSCOMPL SET YTC_TIERSLIVRE=(SELECT T_AUXILIAIRE FROM TIERS'
                        + ' WHERE T_TIERS=YTC_TIERS AND T_AUXILIAIRE = YTC_AUXILIAIRE) WHERE YTC_TIERSLIVRE="" AND (SELECT T_AUXILIAIRE'
                        + ' FROM TIERS WHERE T_TIERS=YTC_TIERS AND T_AUXILIAIRE = YTC_AUXILIAIRE) IS NOT NULL' );

  //C DUMAS Le 12/06/2008 Version 8.1.850.19
  ExecuteSQLContOnExcept ('update DETABLES set dt_libelle = "Liaison critères et tiers" where dt_prefixe = "RQL" and dt_libelle = "Liaison criteres et tiers"');

  //M MORRETTON Le 18/06/2008 Version 8.1.850.99
  ExecuteSQLNoPCL('UPDATE WPARAMFONCTION SET WPF_VARCHAR02="L" WHERE WPF_CODEFONCTION="DELAIMOYFAB" AND WPF_VARCHAR02=""');
  // BTP
  ExecuteSQLContOnExcept('UPDATE NOMENENT SET GNE_DOMAINE="" WHERE GNE_DOMAINE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE SOCIETE SET SO_PRENOMCONTACT="" WHERE SO_PRENOMCONTACT IS NULL');
  ExecuteSQLContOnExcept ('UPDATE AFFAIRE SET AFF_DETECHEANCE="DME",AFF_NUMEROTACHE=0 WHERE AFF_DETECHEANCE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE TACHE SET ATA_BTETAT="" WHERE ATA_BTETAT IS NULL');
  ExecuteSQLContOnExcept ('UPDATE AFMODELETACHE SET AFM_BTETAT="" WHERE AFM_BTETAT IS NULL');
  UpdateDecoupeLigne('GL_COEFFR=0, GL_COEFFC=0, GL_COEFFG=0, GL_COEFMARG=0, GL_MONTANTFR=0, GL_MONTANTFC=0, '+
                      'GL_MONTANTFG=0, GL_MONTANTPA=0, GL_MONTANTPAFR=0, GL_MONTANTPAFC=0, GL_MONTANTPAFG=0, '+
                      'GL_MONTANTPR=0, GL_PRXACHBASE=0, GL_REMISES="", GL_TPSUNITAIRE = 0',' AND (GL_MONTANTPA IS NULL)');
  UpDateDecoupeLigneCompl ('GLC_NONAPPLICFC=GLC_NONAPPLICFRAIS,GLC_NONAPPLICFG="-"', ' AND GLC_NONAPPLICFG IS NULL');
  UpDateDecoupeLigneOUV ('BLO_NONAPPLICFC=BLO_NONAPPLICFRAIS, BLO_NONAPPLICFG="-", BLO_COEFFR=0, '+
                          'BLO_COEFFC=0, BLO_COEFFG=0, BLO_COEFMARG=0, BLO_MONTANTFR=0, '+
                          'BLO_MONTANTFC=0, BLO_MONTANTFG=0, BLO_MONTANTPA=0, BLO_MONTANTPAFR=0, '+
                          'BLO_MONTANTPAFC=0, BLO_MONTANTPAFG=0, BLO_MONTANTPR=0, BLO_MONTANTHT=0, '+
                          'BLO_MONTANTHTDEV=0, BLO_MONTANTTTC=0, BLO_MONTANTTTCDEV=0, BLO_TPSUNITAIRE=0',' AND BLO_MONTANTPA IS NULL');
  UpDateDecoupeLigneOUV ('BLO_DOMAINE=""', ' AND BLO_DOMAINE IS NULL');
  UpDateDecoupePiece ('GP_APPLICFCST=GP_APPLICFGST,GP_COEFFC=0,GP_COEFFR=0,GP_MONTANTPA=0,GP_MONTANTPAFR=0,GP_MONTANTPAFC=0,GP_MONTANTPAFG=0,GP_MONTANTFG=0,GP_MONTANTFC=0,GP_MONTANTFR=0,GP_MONTANTPR=0,GP_TPSUNITAIRE=0',' AND GP_MONTANTPA IS NULL');
  ExecuteSQLContOnExcept ('UPDATE PIEDBASE SET GPB_DELTA=0,GPB_DELTADEV=0,GPB_BASETAXETTC=0,GPB_BASETTCDEV=0 WHERE GPB_DELTA IS NULL');
  ExecuteSQLContOnExcept ('DELETE FROM BTMAJSTRUCTURES WHERE BTV_TYPEELT <> "VERSION"');
	if not ExisteSQL('SELECT 1 FROM BTMAJSTRUCTURES WHERE BTV_TYPEELT="VERSION"') then
  begin
  	ExecuteSQLContOnExcept('insert into BTMAJSTRUCTURES (BTB_IDBTP,BTV_VERSIONBASEB,BTV_TYPEELT,BTV_NOMELT,BTV_WITHDATA) VALUES (0,"850","VERSION","-")');
  end else
  begin
  	ExecuteSQLContOnExcept ('UPDATE BTMAJSTRUCTURES SET BTV_VERSIONBASEB=850 WHERE BTV_TYPEELT="VERSION"');
	end;
  // Maj FV
	if ExisteSQL ('SELECT 1 FROM PARAMSOC WHERE SOC_DATA="OUI" AND SOC_NOM="SO_FV12120"') then
  begin
  	ExecuteSQLContOnExcept('UPDATE AFFAIRE SET AFF_AFFAIREHT=(SELECT  ##TOP 1## T_FACTUREHT FROM TIERS WHERE T_TIERS=AFF_TIERS AND T_NATUREAUXI="CLI") WHERE AFF_AFFAIRE0="W"');
  	ExecuteSQLContOnExcept('UPDATE PIECE SET GP_FACTUREHT=(SELECT AFF_AFFAIREHT FROM AFFAIRE WHERE AFF_AFFAIRE=GP_AFFAIRE) WHERE SUBSTRING(GP_AFFAIRE,1,1)="W"');
  	ExecuteSQLContOnExcept('UPDATE LIGNE SET GL_FACTUREHT=(SELECT AFF_AFFAIREHT FROM AFFAIRE WHERE AFF_AFFAIRE=GL_AFFAIRE) WHERE SUBSTRING(GL_AFFAIRE,1,1)="W"');
		ExecuteSQL ('UPDATE PARAMSOC SET SOC_DATA="OUI" WHERE SOC_NOM="SO_FV12120"');
  end;
  ExecuteSQLContOnExcept ('UPDATE CODECPTA SET GCP_FAMILLETAXE="" WHERE GCP_FAMILLETAXE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE DECISIONACHLIG SET BAD_PABASE=0 WHERE BAD_PABASE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE AFFAIRE SET AFF_FACTURE=(SELECT  ##TOP 1## T_FACTURE FROM TIERS WHERE T_TIERS=AFF_TIERS AND T_NATUREAUXI="CLI") WHERE AFF_FACTURE IS NULL');

End;

Procedure MajVer900;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //Stéphane BOUSSERT Le 19/03/2008 Version 9.0.900.22 Demande N° 2254
  // Ajout du champs T_RVA dans les listes suivantes :
  AglNettoieListes( 'CPRLANCETRA', 'T_RVA;', nil );
  AglNettoieListes( 'CPRELANCECLIENT', 'T_RVA;', nil );
  AglNettoieListes( 'CPRELANCEDIV', 'T_RVA;', nil );
  //
  // BTP
  ExecuteSQLContOnExcept ('UPDATE CODECPTA SET GCP_FAMILLETAXE="" WHERE GCP_FAMILLETAXE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE DECISIONACHLIG SET BAD_PABASE=0 WHERE BAD_PABASE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE DECISIONACHLIG SET BAD_COEFUAUS=0 WHERE BAD_COEFUAUS IS NULL');
  ExecuteSQLContOnExcept ('UPDATE DECISIONACHLFOU SET BDF_COEFUAUS=0 WHERE BDF_COEFUAUS IS NULL');
  UpdateDecoupeLigneOuvPlat('BOP_MONTANTPR=0, BOP_COEFFG=0, BOP_COEFFR=0, BOP_COEFFC=0',' AND (BOP_COEFFC IS NULL)');
  ExecuteSQLContOnExcept ('UPDATE AFFAIRE SET AFF_FACTURE=(SELECT  ##TOP 1## T_FACTURE FROM TIERS WHERE T_TIERS=AFF_TIERS AND T_NATUREAUXI IN ("CLI","PRO")) WHERE AFF_FACTURE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE FACTAFF SET AFA_DATEDEBUTFAC='+DatetoStr(IDate1900)+',AFA_DATEFINFAC='+DateToStr(Idate1900)+' WHERE AFA_DATEDEBUTFAC IS NULL');
  ExecuteSQLContOnExcept ('UPDATE HRPARAMPLANNING SET HPP_REGCOL1="-", HPP_REGCOL2="-", HPP_REGCOL3="-" WHERE HPP_REGCOL1 IS NULL');
  ExecuteSQLContOnExcept ('UPDATE PARCTIERS SET BP1_DATEFINSERIA="20991231" WHERE BP1_DATEFINSERIA IS NULL');
  //

End;

Procedure MajVer901;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
End;

Procedure MajVer902;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

End;

Procedure MajVer903;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //Joël SICH Le 09/04/2008 Version 9.0.903.26 Demande N° 2328
  InsertChoixCode( 'GLD', '1', 'FRANCE', '', '' );
  InsertChoixCode( 'GLD', '2', 'UE', '', '' );
  InsertChoixCode( 'GLD', '3', 'HORS UE', '', '' );

  //Pascal JOUIN Le 09/04/2008 Version 9.0.903.26 Demande N° 2335
  If Not IsDossierPCL Then
  Begin
//    ExecuteSQLContOnExcept( 'UPDATE CODECPTA SET GCP_DEPOT = "", GCP_NATUREPIECE = "", GCP_NATUREMVTSTK = "", ' +
//                              'GCP_MOTIFMVTSTK = "", GCP_AXEVALOPR = "", GCP_CPTEGENVARCS = ""  WHERE GCP_DEPOT IS NULL');
  End;

End;

Procedure MajVer904;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  ExecuteSQLContOnExcept( 'UPDATE TIERS SET T_LOCALTAX="FRA" WHERE (T_LOCALTAX = "") OR (T_LOCALTAX IS NULL)');
  ExecuteSQLContOnExcept( 'UPDATE ADRESSES SET ADR_LOCALTAX="FRA" WHERE (ADR_LOCALTAX = "") OR (ADR_LOCALTAX IS NULL)');
  ExecuteSQLContOnExcept( 'UPDATE ETABLISS SET ET_LOCALTAX="FRA" WHERE (ET_LOCALTAX = "") OR (ET_LOCALTAX IS NULL)');


  //Marie-Christine DESSEIGNET Le 16/04/2008 Version 9.0.904.32 Demande N° 2370
  ExecuteSQLContOnExcept( 'UPDATE TIERSCOMPL SET YTC_ASEGMENTER="-",YTC_ASEGMENTERGRP="-",YTC_APPORTEUR3="",YTC_APPORTEUR2="" WHERE YTC_ASEGMENTER IS NULL' );

End;

Procedure MajVer905;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;


End;

Procedure MajVer906;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
End;

Procedure MajVer907;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //Béatrice Mériaux  Le 07/05/2008 Version 9.0.907.3 Demande N° 2441
  MajConventions;

  //Martine Vermo-Gauchy Le 07/05/2008 Version 9.0.907.3 Demande N° 2444
  ExecuteSQLContOnExcept( 'UPDATE DPPERSO SET DPP_STATUTSOCIAL="AUT"' );
  ExecuteSQLContOnExcept( 'UPDATE ANNULIEN SET ANL_ESTNONCADRE="-", ANL_EVOLREV=0, ANL_ALSACEMOSELLE="-", ANL_RETRAITE=65' );
  ExecuteSQLContOnExcept( 'UPDATE DPSOCIAL SET DSO_TRANSPORT="-", DSO_ACCIDENT="-", DSO_ART83="-", DSO_PREV="-", DSO_COMPSANTE="-"' );

(*
  //Marie-Christine DESSEIGNET Le 07/05/2008 Version 9.0.907.3 Demande N° 2453
  If Not IsDossierPCL Then
    ExecuteSQLContOnExcept( 'UPDATE PROJETS SET RPJ_TYPEPROJET="" WHERE RPJ_TYPEPROJET IS NULL' );
*)
End;

Procedure MajVer909;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //Martine Vermo-Gauchy Le 21/05/2008 Version 9.0.909.2 Demande N° 2495
  ExecuteSQLContOnExcept( 'UPDATE DPPERSO SET DPP_RETRAITE=65' );

End;

Procedure MajVer910;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //M FAUDEL Le 04/06/2008 Version 9.0.911.16 Demande N° 2534

//  ExecuteSQLContOnExcept( 'Update REMUNERATION Set PRM_VARTESTAPPLIC = "" ');
//  ExecuteSQLContOnExcept( 'Update COTISATION Set PCT_VARTESTAPPLIC = "" ');
  ExecuteSQLContOnExcept( 'Update YPARAMEDITION SET YED_ORDRETRI = SUBSTRING(YED_CODEID, 3, 2) '+
    'WHERE (YED_ORDRETRI is null) and (YED_CODEID is not null) and (YED_CODEID <> "") ');

{ //JS1 le 06/06/2008 : ON NE FAIT JAMAIS CELA DANS LE MAJVER !!! C'EST UNE REQUETE SAL A PASSER SUR LA SOCREF !!  SINON, TOUT BUSINESS A LES TABLES !!
  ExecuteSQLContOnExcept( 'INSERT INTO DESHARE (DS_NOMTABLE, DS_MODEFONC, DS_NOMBASE, DS_TYPTABLE, DS_VUE) VALUES ("PGASSOCETATGED", "DB0", "DB000000", "TAB", "")');}
  ExecuteSQLContOnExcept( 'Update REMUNERATION Set PRM_VARTESTAPPLIC = "", PRM_AJOUTIMPO="-" ');
  ExecuteSQLContOnExcept( 'UPDATE RETENUESALAIRE SET PRE_NONLIMITE10P="-"');
//  ExecuteSQLContOnExcept( 'UPDATE ETABCOMPL SET ETB_PEXOMSA=""');
//  ExecuteSQLContOnExcept( 'UPDATE DEPORTSAL SET PSE_TYPEPEXOMSA="ETB",PSE_PEXOMSA=""');
//  ExecuteSQLContOnExcept( 'UPDATE SALARIES SET PSA_TYPNBJOUTRAV="PER"');
  ExecuteSQLContOnExcept( 'Update tablediment set PTE_SENS1 = PTE_SENSINT WHERE PTE_SENSINT <> "" AND PTE_SENSINT IS NOT NULL');
  ExecuteSQLContOnExcept( 'Update tablediment set PTE_STCRITERE1 = "000"||LTRIM(STR(PTE_CRITERE1)) WHERE PTE_CRITERE1>0 AND PTE_CRITERE1<10');
  ExecuteSQLContOnExcept( 'Update tablediment set PTE_STCRITERE1 = "00"||LTRIM(STR(PTE_CRITERE1)) WHERE PTE_CRITERE1>9 AND PTE_CRITERE1<100');
  ExecuteSQLContOnExcept( 'Update tablediment set PTE_STCRITERE1 = "0"||LTRIM(STR(PTE_CRITERE1)) WHERE PTE_CRITERE1>99 AND PTE_CRITERE1<1000');
  ExecuteSQLContOnExcept( 'Update tablediment set PTE_STCRITERE1 = PTE_CRITERE1 WHERE PTE_CRITERE1>999');
  ExecuteSQLContOnExcept( 'Update tablediment set PTE_STCRITERE1 = "" WHERE PTE_CRITERE1=0');
  ExecuteSQLContOnExcept( 'Update tablediment set PTE_STCRITERE2 = "000"||LTRIM(STR(PTE_CRITERE2)) WHERE PTE_CRITERE2>0 AND PTE_CRITERE2<10');
  ExecuteSQLContOnExcept( 'Update tablediment set PTE_STCRITERE2 = "00"||LTRIM(STR(PTE_CRITERE2)) WHERE PTE_CRITERE2>9 AND PTE_CRITERE2<100');
  ExecuteSQLContOnExcept( 'Update tablediment set PTE_STCRITERE2 = "0"||LTRIM(STR(PTE_CRITERE2)) WHERE PTE_CRITERE2>99 AND PTE_CRITERE2<1000');
  ExecuteSQLContOnExcept( 'Update tablediment set PTE_STCRITERE2 = PTE_CRITERE2 WHERE PTE_CRITERE2>999');
  ExecuteSQLContOnExcept( 'Update tablediment set PTE_STCRITERE2 = "" WHERE PTE_CRITERE2=0');
  ExecuteSQLContOnExcept( 'Update TABLEDIMENT Set PTE_NATURECRITERE1 = PTE_NATURETABLE');
  ExecuteSQLContOnExcept( 'Update TABLEDIMENT Set PTE_NATURECRITERE2 = PTE_NATURETABLE WHERE (PTE_CRITERE2 IS NOT NULL AND PTE_CRITERE2<>"")');

  ExecuteSQLContOnExcept( 'Update TABLEDIMENT Set PTE_NATURECRITERE2 = "" WHERE (PTE_CRITERE2 IS NULL OR PTE_CRITERE2="")');
  ExecuteSQLContOnExcept( 'Update TABLEDIMENT Set PTE_LIBSTCRITERE1 = "",  PTE_LIBSTCRITERE2 = "" WHERE PTE_NATURETABLE = "COD"');
  ExecuteSQLContOnExcept( 'Update TABLEDIMENT Set PTE_LIBSTCRITERE1 = (Select PAI_LIBELLE from PAIEPARIM where PAI_IDENT = PTE_CRITERE1) ,  PTE_LIBSTCRITERE2 = '+
    ' (Select PAI_LIBELLE from PAIEPARIM where PAI_IDENT = PTE_CRITERE2) WHERE PTE_NATURETABLE = "DSA"');
  ExecuteSQLContOnExcept( 'Update TABLEDIMENT Set PTE_LIBSTCRITERE1 = "" where PTE_LIBSTCRITERE1 is null');
  ExecuteSQLContOnExcept( 'Update TABLEDIMENT Set PTE_LIBSTCRITERE2 = "" where PTE_LIBSTCRITERE2 is null');
//VG 13012010 on essaye de filtrer un peu plus pour eviter un plantage
//ExecuteSQLContOnExcept( 'Update TABLEDIMENT Set PTE_LIBSTCRITERE1 =  (Select PVA_LIBELLE from VARIABLEPAIE WHERE PVA_VARIABLE = PTE_CRITERE1) , PTE_LIBSTCRITERE2 = '+
//    '(Select PVA_LIBELLE from VARIABLEPAIE WHERE PVA_VARIABLE = PTE_CRITERE2) where PTE_NATURETABLE = "VAR"');
  ExecuteSQLContOnExcept( 'Update TABLEDIMENT Set PTE_LIBSTCRITERE1 =  (Select PVA_LIBELLE from VARIABLEPAIE WHERE PVA_TYPEVARIABLE="PAI" AND PVA_VARIABLE = PTE_CRITERE1) , PTE_LIBSTCRITERE2 = '+
    '(Select PVA_LIBELLE from VARIABLEPAIE WHERE PVA_TYPEVARIABLE="PAI" AND PVA_VARIABLE = PTE_CRITERE2) where PTE_NATURETABLE = "VAR"');
  ExecuteSQLContOnExcept( 'Update TABLEDIMENT Set PTE_LIBSTCRITERE1 =  (Select max(PEL_LIBELLE) from ELTNATIONAUX WHERE PEL_CODEELT = PTE_CRITERE1) , PTE_LIBSTCRITERE2 = '+
    '(Select max(PEL_LIBELLE) from ELTNATIONAUX WHERE PEL_CODEELT = PTE_CRITERE2) where PTE_NATURETABLE = "ELT"');
//  ExecuteSQLContOnExcept( 'UPDATE PROFILPAIE SET PPI_TYPPERSOURSSAF = "",PPI_TYPPERSOURSSAM = "", PPI_CONVENTION= ""');
  ExecuteSQLContOnExcept( 'UPDATE SALARIES SET PSA_TYPNBJOUTRAV="PER", PSA_TYPEPERSOIRC=""');
//  ExecuteSQLContOnExcept( 'UPDATE PAIEENCOURS SET PPU_PROFILRBS="",PPU_TYPPERSOURSSAF="",PPU_TYPPERSOURSSAM="",PPU_TYPEPERSOIRC=""');
  ExecuteSQLContOnExcept( 'Update COTISATION Set PCT_VARTESTAPPLIC = "", PCT_FACTORISABLE="-"');
  ExecuteSQLContOnExcept( 'UPDATE DUCSPARAM SET PDP_ASSIETTE=SUBSTRING(PDP_CODIFICATION, 7, 1) WHERE PDP_CODIFICATION LIKE "1%" AND '+
    '(SUBSTRING(PDP_CODIFICATION, 7, 1)="D" OR SUBSTRING(PDP_CODIFICATION, 7, 1)="P")');
  ExecuteSQLContOnExcept( 'UPDATE DUCSPARAM SET PDP_ASSIETTE="" WHERE PDP_CODIFICATION LIKE "1%" AND (SUBSTRING(PDP_CODIFICATION, 7, 1)<>"D"'+
    ' AND SUBSTRING(PDP_CODIFICATION, 7, 1)<>"P")');
  ExecuteSQLContOnExcept( 'UPDATE DUCSPARAM SET PDP_ASSIETTE="0"||SUBSTRING(PDP_CODIFICATION, 6, 2)  WHERE (LEFT(PDP_CODIFICATION, 1) >= "3" '+
    'AND LEFT(PDP_CODIFICATION, 1) < "6") OR (LEFT(PDP_CODIFICATION, 1) >= "8" AND PDP_CODIFICATION <= "9ZZZZZZ")');
  ExecuteSQLContOnExcept( 'UPDATE DUCSPARAM SET PDP_ASSIETTE="" WHERE (PDP_CODIFICATION >="2" AND  PDP_CODIFICATION < "3") OR  (PDP_CODIFICATION >="6" AND PDP_CODIFICATION < "8")');
  ExecuteSQLContOnExcept( 'UPDATE DUCSPARAM SET PDP_TYPEPERSO=SUBSTRING(PDP_CODIFICATION, 4, 3), PDP_TYPEPERSOAM=SUBSTRING(PDP_CODIFALSACE, 4, 3) WHERE PDP_CODIFICATION LIKE "1%"');
//  ExecuteSQLContOnExcept( 'UPDATE DUCSPARAM SET PDP_TYPEPERSOAM=SUBSTRING(PDP_CODIFALSACE, 4, 3) WHERE PDP_CODIFICATION LIKE "1%"');
  ExecuteSQLContOnExcept( 'UPDATE DUCSPARAM SET PDP_TYPEPERSO=SUBSTRING(PDP_CODIFICATION, 3,3)  WHERE (LEFT(PDP_CODIFICATION, 1) >= "3" '+
    'AND LEFT(PDP_CODIFICATION, 1) < "6") OR (LEFT(PDP_CODIFICATION, 1) >= "8" AND PDP_CODIFICATION <= "9ZZZZZZ")');
  ExecuteSQLContOnExcept( 'UPDATE DUCSPARAM SET PDP_TYPEPERSOAM=SUBSTRING(PDP_CODIFICATION, 3,3)  WHERE (LEFT(PDP_CODIFICATION, 1) >= "3" '+
    'AND LEFT(PDP_CODIFICATION, 1) < "6") OR (LEFT(PDP_CODIFICATION, 1) >= "8" ANd PDP_CODIFICATION <= "9ZZZZZZ")');
  ExecuteSQLContOnExcept( 'UPDATE DUCSPARAM SET PDP_TYPEPERSO ="", PDP_TYPEPERSOAM ="" WHERE (PDP_CODIFICATION >="2" AND  PDP_CODIFICATION < "3") OR '+
    '(PDP_CODIFICATION >="6" AND PDP_CODIFICATION < "8")');
//  ExecuteSQLContOnExcept( 'UPDATE DUCSPARAM SET PDP_TYPEPERSOAM ="" WHERE (PDP_CODIFICATION >="2" AND  PDP_CODIFICATION < "3") OR '+
//    '(PDP_CODIFICATION >="6" AND PDP_CODIFICATION < "8")');
  ExecuteSQLContOnExcept( 'UPDATE DUCSDETAIL SET PDD_ASSIETTE=SUBSTRING(PDD_CODIFICATION, 7, 1) WHERE PDD_CODIFICATION LIKE "1%" AND '+
    '(SUBSTRING(PDD_CODIFICATION, 7, 1)="D" OR SUBSTRING(PDD_CODIFICATION, 7, 1)="P")');
  ExecuteSQLContOnExcept( 'UPDATE DUCSDETAIL SET PDD_ASSIETTE="" WHERE  PDD_CODIFICATION LIKE "1%" AND (SUBSTRING(PDD_CODIFICATION, 7, 1)<>"D" '+
    'AND SUBSTRING(PDD_CODIFICATION, 7, 1)<>"P")');
  ExecuteSQLContOnExcept( 'UPDATE DUCSDETAIL SET PDD_TYPEPERSO=SUBSTRING(PDD_CODIFICATION, 4, 3) WHERE PDD_CODIFICATION LIKE "1%"');
  ExecuteSQLContOnExcept( 'UPDATE DUCSDETAIL SET PDD_ASSIETTE="0"||SUBSTRING(PDD_CODIFICATION, 6, 2), PDD_TYPEPERSO=SUBSTRING(PDD_CODIFICATION, 3,3) WHERE  (LEFT(PDD_CODIFICATION, 1) >= "3" '+
    'AND  LEFT(PDD_CODIFICATION, 1) < "6") OR (LEFT(PDD_CODIFICATION, 1) >= "8" AND  PDD_CODIFICATION <= "9ZZZZZZ")');
//  ExecuteSQLContOnExcept( 'UPDATE DUCSDETAIL SET PDD_TYPEPERSO=SUBSTRING(PDD_CODIFICATION, 3,3)  WHERE (LEFT(PDD_CODIFICATION, 1) >= "3" '+
//    'AND LEFT(PDD_CODIFICATION, 1) < "6") OR (LEFT(PDD_CODIFICATION, 1) >= "8" AND  PDD_CODIFICATION <= "9ZZZZZZ")');
  ExecuteSQLContOnExcept( 'UPDATE DUCSDETAIL SET PDD_ASSIETTE="", PDD_TYPEPERSO ="" WHERE  (PDD_CODIFICATION >="2" AND  PDD_CODIFICATION < "3") '+
    'OR (PDD_CODIFICATION >="6" AND  PDD_CODIFICATION < "8")');
//  ExecuteSQLContOnExcept( 'UPDATE DUCSDETAIL SET PDD_TYPEPERSO ="" WHERE (PDD_CODIFICATION >="2" AND PDD_CODIFICATION < "3") '+
//    'OR (PDD_CODIFICATION >="6" AND PDD_CODIFICATION < "8")');

  //MC DESSEIGNET Le 04/06/2008 Version 9.0.911.16 Demande N° 2560
//  ExecuteSQLNoPCL( 'update perspectives set rpe_typeprevimiss ="" where rpe_typeprevimiss is null');

  //B MERIAUX Le 04/06/2008 Version 9.0.911.16 Demande N° 2561
  ExecuteSQLContOnExcept('UPDATE JURIDIQUE SET JUR_DATEMISREDJUD = "'+UsDateTime_(iDate1900)+'" WHERE JUR_DATEMISREDJUD is null');

  //M DESGOUTTE Le 04/06/2008 Version 9.0.911.16 Demande N° 2486
  ExecuteSQLContOnExcept('update ANNUAIRE set ANN_PACAGE=""');
  ExecuteSQLContOnExcept('update DPAGRICOLE set DAG_107GEL="-", DAG_108SEMENCE="-"');
{M DESGOUTTE : on ne fait pas l'update car nouvelle table
  ExecuteSQLContOnExcept('update DPPARCELLES set DPC_CULTUREBIO="-",'+
    ' DPC_CULTUREDEROBE="-", DPC_PARCELLEIRRIG="-", DPC_PARCELLEDRAINE="-",'+
    ' DPC_PARCELLEATT="", DPC_DATEDEBEXPL="'+UsDateTime_(iDate1900)+'",' +
    ' DPC_DATEFINEXPL="'+UsDateTime_(iDate2099)+'"');}


End;

Procedure MajVer911;
var iInd : integer;
//    PontTax : IMajPontTax;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //J SICH Le 04/06/2008 Version 9.0.911.16 Demande N° 2572
//  ExecuteSQLNoPCL('UPDATE ARTICLECOMPL SET GA2_DESACTIVEPCO="-", GA2_POURCOMBIEN=0, GA2_TOTALAUTO="", GA2_VARIANTE="", ' +
//                    'GA2_OPTION="-" WHERE GA2_DESACTIVEPCO IS NULL');

//  if not IsDossierPCL then
//    UpDateDecoupeLigne ('GLC_CHAPITRENIV1="", GLC_CHAPITRENIV2="", GLC_CHAPITRENIV3="", GLC_CHAPITRENIV4="", GLC_CHAPITRENIV5="", ' +
//                        'GLC_CHAPITREPERELG=0, GLC_POURCOMBIEN=0, GLC_QLIENCHAPITRE=0, GLC_VARIANTE="", GLC_OPTION="-"',
//                      'AND GLC_CHAPITRENIV1 IS NULL', 'GLC');
//  ExecuteSQLNoPCL('UPDATE LIGNECOMPL SET GLC_CHAPITRENIV1="" WHERE GLC_CHAPITRENIV1 IS NULL');
//  ExecuteSQLNoPCL('UPDATE LIGNECOMPL SET GLC_CHAPITRENIV2="" WHERE GLC_CHAPITRENIV2 IS NULL');
//  ExecuteSQLNoPCL('UPDATE LIGNECOMPL SET GLC_CHAPITRENIV3="" WHERE GLC_CHAPITRENIV3 IS NULL');
//  ExecuteSQLNoPCL('UPDATE LIGNECOMPL SET GLC_CHAPITRENIV4="" WHERE GLC_CHAPITRENIV4 IS NULL');
//  ExecuteSQLNoPCL('UPDATE LIGNECOMPL SET GLC_CHAPITRENIV5="" WHERE GLC_CHAPITRENIV5 IS NULL');
//  ExecuteSQLNoPCL('UPDATE LIGNECOMPL SET GLC_CHAPITREPERELG=0 WHERE GLC_CHAPITREPERELG IS NULL');
//  ExecuteSQLNoPCL('UPDATE LIGNECOMPL SET GLC_POURCOMBIEN=0 WHERE GLC_POURCOMBIEN IS NULL');
//  ExecuteSQLNoPCL('UPDATE LIGNECOMPL SET GLC_QLIENCHAPITRE=0 WHERE GLC_QLIENCHAPITRE IS NULL');
//  ExecuteSQLNoPCL('UPDATE LIGNECOMPL SET GLC_VARIANTE="" WHERE GLC_VARIANTE IS NULL');
//  ExecuteSQLNoPCL('UPDATE LIGNECOMPL SET GLC_OPTION="-" WHERE GLC_OPTION IS NULL');

  //J SICH Le 04/06/2008 Version 9.0.911.16 Demande N° 2575
  (*
  for iInd := 1 to 9 do
    InsertChoixCode('GVR', '001', 'Variante ' + intTostr(iInd), 'Variante ' + intTostr(iInd), '');
  *)
  //N CHAVANNE Le 04/06/2008 Version 9.0.911.16 Demande N° 2608
//  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_CONTROLE = (DH_CONTROLE || "1")  WHERE DH_NOMCHAMP = "GA2_NUMEROSERIEGR" AND DH_CONTROLE NOT LIKE "%1%" ');

{ //JS1 le 06/06/2008 : ON NE FAIT JAMAIS CELA DANS LE MAJVER !!! C'EST UNE REQUETE SAL A PASSER SUR LA SOCREF !!  SINON, TOUT BUSINESS A LES TABLES !!
  //B MERIAUX Le 04/06/2008 Version 9.0.911.16 Demande n° 2614
  ExecuteSQLContOnExcept('insert into deshare (ds_NOMtable, ds_modefonc, ds_nombase, ds_typtable, ds_vue) '+
    'values ("LIENJURGRP", "DB0", "DB000000", "TAB", "")');}

  //J SICH Le 04/06/2008 Version 9.0.911.16 Demande n° 2624
//  ExecuteSQLContOnExcept('delete from modeles where mo_nature="GEA" and mo_code in("EA1","EA2","EAM","EQA")');
//  ExecuteSQLContOnExcept('delete from modedata where md_cle in("EGEAEA1%","EGEAEA2%","EGEAEAM%","EGEAEQA%")');

  //D BROSSET Le 04/06/2008 Version 9.0.911.16 Demande n° 2452

{
	PontTax := CreerObjetTransfertTax (GetParamSocPaysNumerique);
  if not PontTax.TransferTaxOldToNew then
    PgiError ('Transfert des taxes');
}
End;

Procedure MajVer912;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //M FAUDEL le 11/06/2008 Version 9.0.912.3 Demande n°2628
  ExecuteSQLContOnExcept('UPDATE PARAMSALARIE SET PPP_NIVEAUMINI=PPP_TYPENIVEAU, PPP_CONVENTION="000" WHERE PPP_PGTYPEINFOLS="ZLS"');
  ExecuteSQLContOnExcept('UPDATE PARAMSALARIE SET PPP_NIVEAUMINI="", PPP_CONVENTION="" WHERE PPP_PGTYPEINFOLS<>"ZLS"');
  ExecuteSQLContOnExcept('UPDATE PGPARAMETRES SET PGP_PGVALCHAMP="CATDADS" WHERE PGP_PGVALCHAMP="DADSCAT" AND PGP_MODULECHAMP="PAI"');

  //M MORETTON le 11/06/2008 Version 9.0.912.3 Demande n°2654
  (*
  ExecuteSQLNoPCL('UPDATE WPARAMFONCTION SET WPF_CODEPORT="", WPF_TRAITEMENT="" WHERE WPF_TRAITEMENT IS NULL');
  ExecuteSQLNoPCL('UPDATE YTARIFSPARAMETRES SET YFO_OKMVT="------" WHERE YFO_OKMVT IS NULL');
  ExecuteSQLNoPCL('UPDATE YTARIFS SET YTS_STKFLUX="", YTS_QUALIFMVT="", YTS_CASCSTKFLUX="-", YTS_CASCQUALIFMVT="-", YTS_CASCTOUSMVT="-" WHERE YTS_STKFLUX IS NULL');
  ExecuteSQLNoPCL('UPDATE WPARAM SET WPA_COMBO10="" WHERE WPA_COMBO10 IS NULL');
  ExecuteSQLNoPCL('UPDATE WPARAMFONCTION SET WPF_CODEPORT="", WPF_RESSOURCE="", WPF_TYPERESSOURCE="", WPF_TARIFRESSOURCE="" WHERE WPF_CODEPORT IS NULL');
  ExecuteSqlNoPCL('UPDATE PIEDPORT SET '
  +'GPT_TYPEFRAIS=(SELECT GPO_TYPEFRAIS FROM PORT WHERE GPO_CODEPORT=GPT_CODEPORT),'
  +'GPT_MODEGROUPEPORT=(SELECT GPO_MODEGROUPEPORT FROM PORT WHERE GPO_CODEPORT=GPT_CODEPORT),'
  +'GPT_REPARTITION=(SELECT GPO_REPARTITION FROM PORT WHERE GPO_CODEPORT=GPT_CODEPORT),'
  +'GPT_VENTPIECE=(SELECT GPO_VENTPIECE FROM PORT WHERE GPO_CODEPORT=GPT_CODEPORT),'
  +'GPT_TYPEFOURNI=(SELECT GPO_TYPEFOURNI FROM PORT WHERE GPO_CODEPORT=GPT_CODEPORT)'
  +'WHERE GPT_TYPEFRAIS IS NULL');
  ExecuteSQLNoPCL('UPDATE LIGNEFRAIS SET '
  +'LF_TYPEFRAIS=(SELECT GPO_TYPEFRAIS FROM PORT WHERE GPO_CODEPORT=LF_CODEPORT),'
  +'LF_MODEGROUPEPORT=(SELECT GPO_MODEGROUPEPORT FROM PORT WHERE GPO_CODEPORT=LF_CODEPORT),'
  +'LF_REPARTITION=(SELECT GPO_REPARTITION FROM PORT WHERE GPO_CODEPORT=LF_CODEPORT),'
  +'LF_VENTPIECE=(SELECT GPO_VENTPIECE FROM PORT WHERE GPO_CODEPORT=LF_CODEPORT),'
  +'LF_TYPEFOURNI=(SELECT GPO_TYPEFOURNI FROM PORT WHERE GPO_CODEPORT=LF_CODEPORT),'
  +'LF_IDFORF=0, LF_IDFIXE=0, LF_IDREMMT=0, LF_IDCOUTBRUT=0, LF_IDCOUTNET=0, LF_IDPOURCENT=0 '
  +'WHERE LF_TYPEFRAIS IS NULL');
  ExecuteSQLNoPCL('UPDATE TIERSFRAIS SET GTF_NATUREPIECEG="", GTF_METHAPPLIFRAIS="" WHERE GTF_NATUREPIECEG IS NULL');

  if (V_PGI.Driver <> dbORACLE7 ) then
   MajPoidsTarifs;

  *)
  //M C DESSEIGNET le 11/06/2008 Version 9.0.912.3 Demande n°2686
  //A faire dans toutes les bases. cette table fait partie aussi des bases réduites
  ExecuteSQLNoPCL('update econtact set eco_guidperanl ="" where eco_guidperanl is null');
  //S BOUSSERT le 11/06/2008 Version 9.0.912.3 Demande n°2692
  // MAJ champ G_NONTAXABLE sur la table GENERAUX
  ExecuteSQLContOnExcept('UPDATE GENERAUX SET G_NONTAXABLE = "X"');
  // MAJ champ GER_NONTAXABLE sur la table GENERAUXREF
  ExecuteSQLContOnExcept('UPDATE GENERAUXREF SET GER_NONTAXABLE = "X"');

  //C DUMAS le 11/06/2008 Version 9.0.912.3 Demande n°2696
  ExecuteSQLContOnExcept ('update DETABLES set dt_libelle = "Liaison critères et tiers"'+
  ' where dt_prefixe = "RQL" and dt_libelle = "Liaison criteres et tiers"');

End;

Procedure MajVer913;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //MC DESSEIGNET le 18/06/2008 Version 9.0.913.1 Demande n° 2750
  ExecuteSqlCOntOnExcept ('update tierscompl set ytc_partenaire="-" where ytc_partenaire is null ');
(*
  //MC DESSEIGNET le 18/06/2008 Version 9.0.913.1 Demande n° 2751
    ExecuteSQLNoPCL(' update operations set ROP_ROPCharLIbre1="" ,ROP_ROPCharLIbre2="" ,ROP_ROPCharLIbre3="" ,'
    +'ROP_ROPTableLIbre6="" ,ROP_ROPTableLIbre7="" ,ROP_ROPTableLIbre8="" ,ROP_ROPTableLIbre9="" ,ROP_ROPTableLIbreA="" ,'
    +'ROP_ROPValLIbre1=0,ROP_ROPValLIbre2=0,ROP_ROPValLIbre3=0,'
    +'Rop_ROPBoolLIbre1="-",ROp_RopBoollibre2="-",Rop_ropBoolLIbre3="-",'
    +'ROP_ROPDateLIbre1="' + UsDateTime_(iDate1900) +' ",ROP_ROPDateLIbre2="' + UsDateTime_(iDate1900) + '",ROP_ROPDateLIbre3="' + UsDateTime_(iDate1900) + '",'
    +'ROP_FraisCreation=0,ROP_FraisROutage=0,ROP_FraisSalle=0,ROP_FraisMailing=0,ROP_FraisCocktail=0,ROP_FraisAutres=0 WHERE ROP_ROPCharLIbre1 IS NULL');

    ExecuteSQLNoPCL('update perspectives set rpe_rpeTableLibre6="",rpe_rpeTableLibre7="",rpe_rpeTableLibre8="",'
    +'rpe_rpeTableLibre9="",rpe_rpeTableLibreA="",'
    +'rpe_rpeValLibre1=0,rpe_rpeValLibre2=0,rpe_rpeValLibre3=0,'
    +'Rpe_RpeBoolLIbre1="-",Rpe_RpeBoollibre2="-",Rpe_rpeBoolLIbre3="-",'
    +'rpe_rpeDateLibre1="' + UsDateTime_(iDate1900) +'",rpe_rpeDateLibre2="' + UsDateTime_(iDate1900) +'",rpe_rpeDateLibre3="' + UsDateTime_(iDate1900) +'", '
    +'rpe_rpeCharLibre1="",rpe_rpeCharLibre2="",rpe_rpeCharLibre3=""  WHERE rpe_rpeTableLibre6 IS NULL');

    ExecuteSQLNoPCL(' update actions set rac_racTableLibre1="",rac_racTableLibre2="",'
    +'rac_racTableLibre3="",rac_racTableLibre4="",rac_racTableLibre5="",rac_racTableLibre6=""'
    +',rac_racTableLibre7="",rac_racTableLibre8="",rac_racTableLibre9="",rac_racTableLibreA="",'
    +'Rac_RacBoolLIbre1="-",Rac_RacBoollibre2="-",Rac_racBoolLIbre3="-",'
    +'rac_racCharLibre1="' + UsDateTime_(iDate1900) +'",rac_racCharLibre2="' + UsDateTime_(iDate1900) +'",rac_racCharLibre3="' + UsDateTime_(iDate1900) +'",'
    +'rac_racDateLibre1="",rac_racDateLibre2="",rac_racDateLibre3="",'
    +'rac_racValLibre1=0,rac_racValLibre2=0,rac_racValLibre3=0 WHERE rac_racTableLibre1 IS NULL');

    ExecuteSQLNoPCL(' update projets set rpj_rpjtablelibre6="",rpj_rpjtablelibre7="",rpj_rpjtablelibre8'
    +'="",rpj_rpjtablelibre9="",rpj_rpjtablelibreA="",rpj_rpjCharlibre1="",rpj_rpjCharlibre2="",'
    +'Rpj_RpjBoolLIbre1="-",Rpj_RpjBoollibre2="-",Rpj_rpjBoolLIbre3="-",'
    +'rpj_rpjCharlibre3="",rpj_rpjDatelibre1="' + UsDateTime_(iDate1900) +'",rpj_rpjDatelibre2="' + UsDateTime_(iDate1900) +'",rpj_rpjDatelibre3="' + UsDateTime_(iDate1900) +'",'
    +'rpj_rpjVallibre1=0,rpj_rpjVallibre2=0,rpj_rpjVallibre3=0 WHERE rpj_rpjtablelibre6 IS NULL');
*)
  //M DESGOUTTE le 18/06/2008 Version 9.0.913.1 Demande n° 2753
    if IsMonoOuCommune then
    ExecuteSQLContOnExcept('UPDATE DPSOCIALCAISSE SET DSC_NUMAFFILIATION="", DSC_NUMINTERNE="", DSC_SIRET="", DSC_INSTITUTION="", DSC_REGROUPEMENT="" WHERE DSC_NUMAFFILIATION=""');

  //M FAUDEL le 18/06/2008 Demande n°2755
    PgMajChampSWS;

  //MC DESSEIGNET Le 18/06/2008 Version 9.0.913.1 Demande N° 2757
//  ExecuteSqlNoPcl ('update afsegmcompl set asc_guidper="",asc_modeleseg="",asc_can=0,asc_can1=0');

End;

Procedure MajVer914;
var i:integer;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  // M FAUDEL le 25/06/2008 Demande 2765
    ExecuteSQLContOnExcept('delete from menu where MN_1=47 and mn_2=1 and mn_3=7');

{  // MC DESSEIGNET le 25/06/2008 Demande 2803
  if not IsDossierPCL then
  begin
    for i:=0 to 2 do InsertChoixCode ('RRO', 'BL'+IntToStr(i+1), '.-Opération décision libre ' + intTostr(i + 1), '','');
    for i:=0 to 2 do InsertChoixCode ('RRO', 'ML'+IntToStr(i+1), '.-Opération montant libre ' + intTostr(i + 1), '','');
    for i:=0 to 2 do InsertChoixCode ('RRO', 'DL'+IntToStr(i+1), '.-Opération date libre ' + intTostr(i + 1), '','');
    for i:=0 to 2 do InsertChoixCode ('RRO', 'CL'+IntToStr(i+1), '.-Opération Texte libre ' + intTostr(i + 1), '','');
    for i:=5 to 8 do InsertChoixCode ('RRO', 'TL'+IntToStr(i+1), '.-Opération Table libre ' + intTostr(i + 1), '','');
    InsertChoixCode ('RRO', 'TLA' ,'.-Opération Table libre 10' , '','');

    for i:=0 to 2 do InsertChoixCode ('RRP', 'BL'+IntToStr(i+1), '.- Projet décision libre ' + intTostr(i + 1), '','');
    for i:=0 to 2 do InsertChoixCode ('RRP', 'ML'+IntToStr(i+1), '.- Projet montant libre ' + intTostr(i + 1), '','');
    for i:=0 to 2 do InsertChoixCode ('RRP', 'DL'+IntToStr(i+1), '.- Projet date libre ' + intTostr(i + 1), '','');
    for i:=0 to 2 do InsertChoixCode ('RRP', 'CL'+IntToStr(i+1), '.- Projet Texte libre ' + intTostr(i + 1), '','');
    for i:=5 to 8 do InsertChoixCode ('RRP', 'TL'+IntToStr(i+1), '.-Projet Table libre ' + intTostr(i + 1), '','');
    InsertChoixCode ('RRP', 'TLA' ,'.-Projet Table libre 10' , '','');

    for i:=0 to 2 do InsertChoixCode ('RRS', 'BL'+IntToStr(i+1), '.- Perspective décision libre ' + intTostr(i + 1), '','');
    for i:=0 to 2 do InsertChoixCode ('RRS', 'ML'+IntToStr(i+1), '.- Perspective montant libre ' + intTostr(i + 1), '','');
    for i:=0 to 2 do InsertChoixCode ('RRS', 'DL'+IntToStr(i+1), '.- Perspective date libre ' + intTostr(i + 1), '','');
    for i:=0 to 2 do InsertChoixCode ('RRS', 'CL'+IntToStr(i+1), '.- Perspective Texte libre ' + intTostr(i + 1), '','');
    for i:=5 to 8 do InsertChoixCode ('RRS', 'TL'+IntToStr(i+1), '.- Perspective t Table libre ' + intTostr(i + 1), '','');
    InsertChoixCode ('RRS', 'TLA' ,'.- Perspective Table libre 10' , '','');
  end;}

  //MC DESSEIGNET 3402
  (*
  if V_PGI.ModePCL='1' then
  begin
    // MC DESSEIGNET le 25/06/2008 Demande 2803
    if not IsDossierPCL then
    begin
      for i:=0 to 2 do InsertChoixCode ('RRO', 'BL'+IntToStr(i+1), '.-Opération décision libre ' + intTostr(i + 1), '','');
      for i:=0 to 2 do InsertChoixCode ('RRO', 'ML'+IntToStr(i+1), '.-Opération montant libre ' + intTostr(i + 1), '','');
      for i:=0 to 2 do InsertChoixCode ('RRO', 'DL'+IntToStr(i+1), '.-Opération date libre ' + intTostr(i + 1), '','');
      for i:=0 to 2 do InsertChoixCode ('RRO', 'CL'+IntToStr(i+1), '.-Opération Texte libre ' + intTostr(i + 1), '','');
      for i:=5 to 8 do InsertChoixCode ('RRO', 'TL'+IntToStr(i+1), '.-Opération Table libre ' + intTostr(i + 1), '','');
      InsertChoixCode ('RRO', 'TLA' ,'.-Opération Table libre 10' , '','');

      for i:=0 to 2 do InsertChoixCode ('RRP', 'BL'+IntToStr(i+1), '.- Projet décision libre ' + intTostr(i + 1), '','');
      for i:=0 to 2 do InsertChoixCode ('RRP', 'ML'+IntToStr(i+1), '.- Projet montant libre ' + intTostr(i + 1), '','');
      for i:=0 to 2 do InsertChoixCode ('RRP', 'DL'+IntToStr(i+1), '.- Projet date libre ' + intTostr(i + 1), '','');
      for i:=0 to 2 do InsertChoixCode ('RRP', 'CL'+IntToStr(i+1), '.- Projet Texte libre ' + intTostr(i + 1), '','');
      for i:=5 to 8 do InsertChoixCode ('RRP', 'TL'+IntToStr(i+1), '.-Projet Table libre ' + intTostr(i + 1), '','');
      InsertChoixCode ('RRP', 'TLA' ,'.-Projet Table libre 10' , '','');

      for i:=0 to 2 do InsertChoixCode ('RRS', 'BL'+IntToStr(i+1), '.- Perspective décision libre ' + intTostr(i + 1), '','');
      for i:=0 to 2 do InsertChoixCode ('RRS', 'ML'+IntToStr(i+1), '.- Perspective montant libre ' + intTostr(i + 1), '','');
      for i:=0 to 2 do InsertChoixCode ('RRS', 'DL'+IntToStr(i+1), '.- Perspective date libre ' + intTostr(i + 1), '','');
      for i:=0 to 2 do InsertChoixCode ('RRS', 'CL'+IntToStr(i+1), '.- Perspective Texte libre ' + intTostr(i + 1), '','');
      for i:=5 to 8 do InsertChoixCode ('RRS', 'TL'+IntToStr(i+1), '.- Perspective Table libre ' + intTostr(i + 1), '','');
      InsertChoixCode ('RRS', 'TLA' ,'.- Perspective Table libre 10' , '','');
    end;
  end
  else begin //en PME, pas de libellé derrrière le .-
    if not IsDossierPCL then
    begin
      for i:=0 to 2 do InsertChoixCode ('RRO', 'BL'+IntToStr(i+1), '.- ' + intTostr(i + 1), '','');
      for i:=0 to 2 do InsertChoixCode ('RRO', 'ML'+IntToStr(i+1), '.- ' + intTostr(i + 1), '','');
      for i:=0 to 2 do InsertChoixCode ('RRO', 'DL'+IntToStr(i+1), '.- ' + intTostr(i + 1), '','');
      for i:=0 to 2 do InsertChoixCode ('RRO', 'CL'+IntToStr(i+1), '.- ' + intTostr(i + 1), '','');
      for i:=5 to 8 do InsertChoixCode ('RRO', 'TL'+IntToStr(i+1), '.- ' + intTostr(i + 1), '','');
      InsertChoixCode ('RRO', 'TLA' ,'.-' , '','');

      for i:=0 to 2 do InsertChoixCode ('RRP', 'BL'+IntToStr(i+1), '.-  ' + intTostr(i + 1), '','');
      for i:=0 to 2 do InsertChoixCode ('RRP', 'ML'+IntToStr(i+1), '.-  ' + intTostr(i + 1), '','');
      for i:=0 to 2 do InsertChoixCode ('RRP', 'DL'+IntToStr(i+1), '.-  ' + intTostr(i + 1), '','');
      for i:=0 to 2 do InsertChoixCode ('RRP', 'CL'+IntToStr(i+1), '.- ' + intTostr(i + 1), '','');
      for i:=5 to 8 do InsertChoixCode ('RRP', 'TL'+IntToStr(i+1), '.- ' + intTostr(i + 1), '','');
      InsertChoixCode ('RRP', 'TLA' ,'.-' , '','');

      for i:=0 to 2 do InsertChoixCode ('RRS', 'BL'+IntToStr(i+1), '.-  ' + intTostr(i + 1), '','');
      for i:=0 to 2 do InsertChoixCode ('RRS', 'ML'+IntToStr(i+1), '.-  ' + intTostr(i + 1), '','');
      for i:=0 to 2 do InsertChoixCode ('RRS', 'DL'+IntToStr(i+1), '.- ' + intTostr(i + 1), '','');
      for i:=0 to 2 do InsertChoixCode ('RRS', 'CL'+IntToStr(i+1), '.-  ' + intTostr(i + 1), '','');
      for i:=5 to 8 do InsertChoixCode ('RRS', 'TL'+IntToStr(i+1), '.-  ' + intTostr(i + 1), '','');
      InsertChoixCode ('RRS', 'TLA' ,'.- ' , '','');
    end;
  end;

  // MC DESSEIGNET le 25/06/2008 Demande 2805
  ExecuteSqlNoPcl ('update paractions set rpa_ferme="-"');
  *)
end;

Procedure MajVer915;
var i : integer;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //M DESGOUTTE le 02/07/2008 Demande 2818
  ExecuteSQLNoPcl('update DPPARCELLES set DPC_SUPERFICIENC=0, DPC_ELIGIBLEAIDE=0, DPC_PROPRIETAIRE="" where DPC_SUPERFICIENC is null');

{  //MC DESSEIGNET le 02/07/2008 Demande 2819
  if not IsDossierPCL then
  begin
   for i:=0 to 2 do insertchoixcode ('RRC','BL'+intToStr(i+1), '.-Actions décision libre '+IntTostr(i+1),'','');
   for i:=0 to 2 do insertchoixcode ('RRC','ML'+intToStr(i+1), '.-Actions montant libre '+IntTostr(i+1),'','');
   for i:=0 to 2 do insertchoixcode ('RRC','DL'+intToStr(i+1), '.-Actions Date libre '+IntTostr(i+1),'','');
   for i:=0 to 2 do insertchoixcode ('RRC','CL'+intToStr(i+1), '.-Actions texte libre '+IntTostr(i+1),'','');
   for i:=0 to 8 do insertchoixcode ('RRC','TL'+intToStr(i+1), '.-Actions table libre '+IntTostr(i+1),'','');
   insertchoixcode ('RRC','TLA', '.-Actions table libre 10','','');
  end;}
  (*
  if V_PGI.ModePCL = '1' then
  begin
    if not IsDossierPCL then
    begin
     for i:=0 to 2 do insertchoixcode ('RRC','BL'+intToStr(i+1), '.-Actions décision libre '+IntTostr(i+1),'','');
     for i:=0 to 2 do insertchoixcode ('RRC','ML'+intToStr(i+1), '.-Actions montant libre '+IntTostr(i+1),'','');
     for i:=0 to 2 do insertchoixcode ('RRC','DL'+intToStr(i+1), '.-Actions Date libre '+IntTostr(i+1),'','');
     for i:=0 to 2 do insertchoixcode ('RRC','CL'+intToStr(i+1), '.-Actions texte libre '+IntTostr(i+1),'','');
     for i:=0 to 8 do insertchoixcode ('RRC','TL'+intToStr(i+1), '.-Actions table libre '+IntTostr(i+1),'','');
     insertchoixcode ('RRC','TLA', '.-Actions table libre 10','','');
    end;
  end
  else begin //en PME pas de libellé après le .-
    if not IsDossierPCL then
    begin
     for i:=0 to 2 do insertchoixcode ('RRC','BL'+intToStr(i+1), '.- '+IntTostr(i+1),'','');
     for i:=0 to 2 do insertchoixcode ('RRC','ML'+intToStr(i+1), '.- '+IntTostr(i+1),'','');
     for i:=0 to 2 do insertchoixcode ('RRC','DL'+intToStr(i+1), '.- '+IntTostr(i+1),'','');
     for i:=0 to 2 do insertchoixcode ('RRC','CL'+intToStr(i+1), '.- '+IntTostr(i+1),'','');
     for i:=0 to 8 do insertchoixcode ('RRC','TL'+intToStr(i+1), '.- '+IntTostr(i+1),'','');
     insertchoixcode ('RRC','TLA', '.-','','');
    end;
  end;
  *)
End;

Procedure MajVer917;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

//  ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA = "-" WHERE SOC_NOM = "SO_OKFAMCPTAPIECE"');

{ GC_20080716_JTR_01015818_Début  }
  ExecuteSQLContOnExcept('UPDATE COMMUN SET CO_LIBRE = "{TOUS};#TOBPARAMLIBIMPACTPCE" WHERE CO_TYPE="GEF" AND CO_CODE = "LIB"');
  ExecuteSQLContOnExcept('UPDATE COMMUN SET CO_LIBRE = "#LIBIMPACT" WHERE CO_TYPE = "GLE" AND CO_CODE IN("001", "002")');
{ GC_20080716_JTR_01015818_Fin  }

//JLS le 23/07/2008 Demande 2936
(*
  ExecuteSQLNOPCL('UPDATE EMPLACEMENT SET GEM_ETATEMPLACE="100" WHERE GEM_ETATEMPLACE IS NULL');
  ExecuteSQLNOPCL('UPDATE EMPLACEMENT SET GEM_GUID=PgiGuid WHERE GEM_GUID IS NULL');
  ExecuteSQLNOPCL('UPDATE WJOURNALACTION SET WJA_GUIDORI="" WHERE WJA_GUIDORI IS NULL');
  //MCD le 23/07/2008 Demande 2938
  ExecuteSqlNoPcl ('Update persphisto set RPh_RPETABLELIBRE6="", RPh_RPETABLELIBRE7="", RPh_RPETABLELIBRE8="" ,RPh_RPETABLELIBRE9="", RPh_RPETABLELIBREA="",'
    +'RPH_RPEVALLIBRE1 =0,RPH_RPEVALLIBRE2 =0,RPH_RPEVALLIBRE3 =0,'
    +'RPH_RPEDATELIBRE1 ="' + UsDateTime_(iDate1900) +'",RPH_RPEDATELIBRE2 ="' + UsDateTime_(iDate1900) +'",RPH_RPEDATELIBRE3 ="' + UsDateTime_(iDate1900) +'",'
    +'RPH_RPECHARLIBRE1="", RPH_RPECHARLIBRE2="", RPH_RPECHARLIBRE3="",'
    +'rph_typeprevimiss ="",'
    +'RPH_APPORTEUR1="",RPH_APPEXTERNE1="-",RPH_MONTANTRETRIB1=0,RPH_MOTIFRETRIB1="",'
    +'RPH_APPORTEUR2="",RPH_APPEXTERNE2="-",RPH_MONTANTRETRIB2=0,RPH_MOTIFRETRIB2="",'
    +'RPH_APPORTEUR3="",RPH_APPEXTERNE3="-",RPH_MONTANTRETRIB3=0,RPH_MOTIFRETRIB3=""  WHERE RPh_RPETABLELIBRE6 IS NULL');

*)
  //S BOUSSERT Le 23/07/2008 Version 9.0.917.1 Demande N° 2932
  AglNettoieListes('CPMULEEXBQLIG','CEL_VALIDE',nil);

  //M DESGOUTTE Le 23/07/2008 Version 9.0.917.1 Demande N° 2949
  if IsMonoOuCommune then ExecuteSQL('delete from juformejur where jfj_forme="SCGA"');
(*
  //JLS Le 23/07/2008 Version 9.0.917.1 Demande N° 2951
  // -------- Création du Guid
  ExecuteSqlNoPCL('update wgammetet set wgt_guid=PgiGuid');

  // -------- Mise à jour de WWS_GUID
  ExecuteSqlNoPCL('update wformconvsave set wws_guidori=ISNULL((select Max(wgt_guid) from wgammetet where wgt_identifiant=wws_numeroori), "")'
                + ' where wws_prefixeori = "WGT"')
                ;

  // Mise à jour WJA
  ExecuteSqlNoPCL('update wjournalaction set wja_guidori=ISNULL((select Max(wgt_guid) from wgammetet where wgt_identifiant=wja_identifiantwxx), "")'
                + ' where wja_prefixe="WGT"')
                ;

  //JLS Le 23/07/2008 Version 9.0.917.1 Demande N° 2952
  // -------- Création du Guid
  ExecuteSqlNoPCL('update wNometet set wnt_guid=PgiGuid');

  // -------- Mise à jour de WWS_GUID
  ExecuteSqlNoPCL('update wformconvsave set wws_guidori=ISNULL((select Max(wnt_guid) from wNometet where wnt_identifiant=wws_numeroori), "")'
                + ' where wws_prefixeori = "WNT"')
                ;

  // Mise à jour WJA
  ExecuteSqlNoPCL('update wjournalaction set wja_guidori=ISNULL((select Max(wnt_guid) from wNometet where wnt_identifiant=wja_identifiantwxx), "")'
                + ' where wja_prefixe="WNT"')
                ;

  // Mise à jour GNE
  ExecuteSqlNoPCL('update NomenEnt set Gne_guidWnt=ISNULL((select Max(wnt_guid) from wNometet where wnt_identifiant=Gne_identifiantwnt), "")');

  // Mise à jour GLC
  if not IsDossierPCL then
    UpDateDecoupeLigne ('Glc_guidWnt=ISNULL((select Max(wnt_guid) from wNometet where wnt_identifiant=Glc_identifiantwnt), "")', '', 'GLC');
//  ExecuteSqlNoPCL('update LigneCompl set Glc_guidWnt=ISNULL((select Max(wnt_guid) from wNometet where wnt_identifiant=Glc_identifiantwnt), "")');

  // Mise à jour GNL
  ExecuteSqlNoPCL('update Nomenlig set Gnl_guidWnt=ISNULL((select Max(wnt_guid) from wNometet where wnt_identifiant=Gnl_identifiantwnt), "")');

  ExecuteSqlNoPCL('Update wGammeLIg set Wgl_Guid = PgiGuid');
  //JLS Le 23/07/2008 Version 9.0.917.1 Demande N° 2964
  // -------- Mise à jour de WPL
  ExecuteSqlNoPCL('update wpdrlig set wpl_guidWgl=ISNULL((select Max(wgl_guid) from wgammelig where wgl_identifiant=wpl_IdentifiantWgl), "")');

  // -------- Mise à jour de WOG
  ExecuteSqlNoPCL('update wOrdregamme set wog_guidWgl=ISNULL((select Max(wgl_guid) from wgammelig where wgl_identifiant=wog_IdentifiantWgl), "")');

//GP_20100901_DS_GP17706
  ExecuteSqlNoPCL('UPDATE WORDRELIG SET '
                        + 'WOL_QARECSAIS= IIF((WOL_ETATLIG = "DCL"), WOL_QACCSAIS'
                        +', ISNULL((SELECT IIF((WOP1.WOP_ETATPHASE IN ("TER","SOL")), WOP1.WOP_QRECSAIS, WOP1.WOP_QARECSAIS)'
                              + ' FROM WORDREPHASE WOP1 WHERE WOP1.WOP_NATURETRAVAIL=WOL_NATURETRAVAIL AND WOP1.WOP_LIGNEORDRE=WOL_LIGNEORDRE'
                              + ' AND WOP1.WOP_OPECIRC=(SELECT MAX(WOP_OPECIRC) FROM WORDREPHASE WOP2 '
                                              + ' WHERE WOP2.WOP_NATURETRAVAIL=WOP1.WOP_NATURETRAVAIL AND WOP2.WOP_LIGNEORDRE=WOP1.WOP_LIGNEORDRE) ),WOL_QACCSAIS) )'
                        + ',WOL_QARECSTOC= IIF((WOL_ETATLIG = "DCL"), WOL_QACCSTOC'
                        + ', ISNULL((SELECT IIF((WOP1.WOP_ETATPHASE IN ("TER","SOL")), WOP1.WOP_QRECSTOC, WOP1.WOP_QARECSTOC)'
                              + ' FROM WORDREPHASE WOP1'
                              + ' WHERE WOP1.WOP_NATURETRAVAIL=WOL_NATURETRAVAIL AND WOP1.WOP_LIGNEORDRE=WOL_LIGNEORDRE'
                              + ' AND WOP1.WOP_OPECIRC=(SELECT MAX(WOP_OPECIRC) FROM WORDREPHASE WOP2 '
                                              +' WHERE WOP2.WOP_NATURETRAVAIL=WOP1.WOP_NATURETRAVAIL AND WOP2.WOP_LIGNEORDRE=WOP1.WOP_LIGNEORDRE) ) , WOL_QACCSTOC) )'
                        +  ',WOL_QREBUSAIS=0'
                        +  ',WOL_QREBUSTOC=0'
                        + ' WHERE ((WOL_QARECSAIS IS NULL) OR (WOL_QARECSAIS=0))');


  //JLS Le 23/07/2008 Version 9.0.917.1 Demande N° 2970
  ExecuteSqlNoPCL('Update wGammeRes set Wgr_Guid=PgiGuid');

  //JTR Le 23/07/2008 Version 9.0.917.1 Demande N° 2971
  ExecuteSQLContOnExcept('UPDATE ANNULIEN SET ANL_GUIDCONV = ""');
  ExecuteSQLContOnExcept('UPDATE ANNULIEN SET ANL_GUIDCONV = ANL_CONVTXT ' +
                   'WHERE len(ANL_CONVTXT) = 36 and ANL_CONVTXT <> "" and ANL_GUIDCONV = ""');

  //JLS Le 23/07/2008 Version 9.0.917.1 Demande N° 2974
  ExecutesqlNoPCL('Update wOrdretet set Wot_Guid=PgiGuid');
  ExecuteSqlNoPCL('update piece set gp_guidwot=ISNULL((select Max(Wot_Guid) from Wordretet where wot_identifiant=gp_identifiantWot), "")');
*)

  //G JUGDE Le 23/07/2008 Version 9.0.917.1 Demande N° 2975
  ExecuteSQLNoPCL('UPDATE QGROUPE SET QGR_GRPAUTONOME="-",QGR_ENGAGEBAC="-" WHERE QGR_GRPAUTONOME IS NULL');
  ExecuteSqlNoPCL('UPDATE QOALEA SET QOA_VENTILTTRS="-",QOA_TTRSMAXISAIS=0,QOA_TTRSMAXIHHCC=0,QOA_UNITEPREV="",QOA_COEFTRS=0 ');
  ExecuteSQLNoPCL('UPDATE QORESPRIME SET QOR_VALLIBRE1=0,QOR_VALLIBRE2=0,QOR_VALLIBRE3=0,QOR_VALLIBRE4=0');
  ExecuteSQLNoPCL('UPDATE QSITE SET QSI_SUIVIOPMES="-",QSI_SUIVIGRPMES=IIF(Exists(SELECT 1 FROM QGROUPE  WHERE QGR_CTX=QSI_CTX AND QGR_SITE=QSI_SITE ),"X","-") WHERE QSI_SUIVIGRPMES IS NULL');
  ExecuteSqlNoPCL('UPDATE QWBACTET SET QWB_QREBUSTOC=0,QWB_QREBUSAIS=0,QWB_SITE="",QWB_GRP="",QWB_QENCOURSSTOC=0 ');
  ExecuteSqlNoPCL('UPDATE QWHISTORES SET QWH_CODECOLISMES="",QWH_DEPOT=(SELECT QSI_DEPOT FROM QSITE WHERE QSI_CTX="0" AND QSI_SITE=QWH_SITE) WHERE QWH_CODECOLISMES IS NULL');
  ExecuteSQLNoPCL('UPDATE QWRELEVELIG SET QWL_TPREVMODIFSAIS=0, QWL_TPREVMODIFHHCC=0');
  (*
  ExecuteSQLNoPCL('UPDATE WNATURETRAVAIL SET WNA_CONTROLEMES=WNA_ACTIF WHERE WNA_CONTROLEMES IS NULL');
  ExecuteSqlNoPCL('UPDATE WGAOPER SET WGO_OPJALON="",'
      + ' WGO_LIBREWGO1="",WGO_LIBREWGO2="",WGO_LIBREWGO3="",WGO_LIBREWGO4="",WGO_LIBREWGO5="",'
      + ' WGO_LIBREWGO6="",WGO_LIBREWGO7="",WGO_LIBREWGO8="",WGO_LIBREWGO9="",WGO_LIBREWGOA="",'
      + ' WGO_VALLIBRE1=0,WGO_VALLIBRE2=0,WGO_VALLIBRE3=0,'
      + ' WGO_DATELIBRE1="' + UsDateTime_(iDate1900) + '", WGO_DATELIBRE2="' + UsDateTime_(iDate1900) + '", '
      + ' WGO_DATELIBRE3="' + UsDateTime_(iDate1900) + '", '
      + ' WGO_BOOLLIBRE1="-",WGO_BOOLLIBRE2="-",WGO_BOOLLIBRE3="-",'
      + ' WGO_CHARLIBRE1=""  ,WGO_CHARLIBRE2=""  ,WGO_CHARLIBRE3=""  '
      + ' WHERE WGO_LIBREWGO1 IS NULL ');
  *)
  //G JUGDE Le 23/07/2008 Version 9.0.917.1 Demande N° 2977
  InsertChoixCode('ZLI', 'RA1' , 'Valeur libre 1', 'Valeur libre 1 ', '');
  InsertChoixCode('ZLI', 'RA2' , 'Valeur libre 2', 'Valeur libre 2 ', '');
  InsertChoixCode('ZLI', 'RA3' , 'Valeur libre 3', 'Valeur libre 3 ', '');
  InsertChoixCode('ZLI', 'RA4' , 'Valeur libre 4', 'Valeur libre 4 ', '');

End;

Procedure MajVer918;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //P FAGES Le 28/07/2008 Version 9.0.918.1 Demande N° 2994
 // ExecuteSQLNoPCL('UPDATE PARPIECE SET GPP_CTRLENCOURS = GPP_ENCOURS WHERE GPP_CTRLENCOURS IS NULL');

  //J TRIFILIEFF Le 28/07/2008 Version 9.0.918.1 Demande N° 2994
  { GC_20080724_JTR_01015998_Début }
//  ExecuteSQLContOnExcept('INSERT INTO COMMUN(CO_TYPE, CO_CODE, CO_LIBELLE, CO_ABREGE, CO_LIBRE) VALUES ("GTY", "STK", "Ecritures de stock", "Stock", "")');
  { GC_20080724_JTR_01015998_Fin }

  //G JUGDE Le 29/07/2008 Version 9.0.918.1 Demande N° 3013
    {--- MAJ WOL_DATESOLDE sur WORDRELIG }
  {ExecuteSQLNoPCL('UPDATE WORDRELIG SET WOL_DATESOLDE=IIF(WOL_ETATLIG="SOL",'
    + '(SELECT MAX(WJA_DATECREATION) FROM WJOURNALACTION WHERE WJA_IDENTIFIANTWXX=WOL_IDENTIFIANT AND WJA_ACTION="wtaSoldeWOL" ),'
    + '"' + UsDateTime_(IDate1900) + '")  WHERE WOL_DATESOLDE IS NULL');}

 //Correction G JUGDE PLANTAGE SUR ORACLE
(*
  ExecuteSQLNoPCL('UPDATE WORDRELIG SET WOL_DATESOLDE="' + UsDateTime_(IDate1900) + '"  WHERE WOL_DATESOLDE IS NULL');

  ExecuteSQLNoPCL('UPDATE WORDRELIG SET WOL_DATESOLDE=(SELECT MAX(WJA_DATECREATION) '
  + ' FROM WJOURNALACTION WHERE WJA_IDENTIFIANTWXX=WOL_IDENTIFIANT AND WJA_ACTION="wtaSoldeWOL" )  '
  + ' WHERE EXISTS (SELECT 1 FROM WJOURNALACTION WHERE WJA_IDENTIFIANTWXX=WOL_IDENTIFIANT AND WJA_ACTION="wtaSoldeWOL" )');

  {--- MAJ WOG_TTRSHHCC sur WORDREGAMME }
  ExecuteSQLNoPCL('UPDATE WORDREGAMME SET WOG_TTRSHHCC=0 WHERE WOG_TTRSHHCC IS NULL');
*)

End;

Procedure MajVer919;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //S BOUSSERT le 27/08/2008 Dem n°3031
  // Ajout de champ E_QUALIFORIGINE dans la liste des releves de factures CPMULRELDETAIL
  AglNettoieListes('CPMULRELDETAIL', 'E_QUALIFORIGINE', nil) ;
  //D SCLAVOPOULOS le 27/08/2008 Dem n°3040
//  AGLNettoieListes('WORDRELIG;WORDRELIG2','WOL_AFFAIRE');
  //M MORRETTON le 27/08/2008 Dem n°3075
  ExecuteSQLNoPCL('DELETE FROM PARAMSOC WHERE SOC_NOM = "SCO_PDRMETHVALO" OR SOC_NOM = "SCO_PDRGENEREWPL"');
  //P FAGES le 27/08/2008 Dem n°3079
  ExecuteSQLContOnExcept('Update dechamps set Dh_libelle="Identifiant article" where dh_prefixe="GCC" and DH_NUMCHAMP=2');
  //JL SAUZET le 27/08/2008 Dem n°3088
  // Mise à jour du champ GL_GUIDWOL
(*
  if not IsDossierPCL then
//GP_20100517_DKZ_GP17467 Déb
  begin
    ExecutesqlNoPCL('UPDATE WORDRELIG SET WOL_GUID=PGIGUID');
    UpDateDecoupeLigne ('Gl_guidwol = IsNull((select max(Wol_Guid) From wOrdreLig Where Wol_Identifiant=Gl_IdentifiantWol), "")', ' and gl_identifiantwol<>0');
    UpDateDecoupeLigne ('Gl_guidwol = ""', ' and gl_identifiantwol=0');
  end;
//GP_20100517_DKZ_GP17467 Fin
//  ExecuteSqlNoPCL('update Ligne set Gl_guidwol = IsNull((select max(Wol_Guid) From wOrdreLig Where Wol_Identifiant=Gl_IdentifiantWol), "") where gl_identifiantwol<>0');
  // Mise à jour du WPL_GUIDWNL
  ExecuteSqlNoPCL('update wpdrLig set wpl_guidwnl = IsNull((select max(Wnl_Guid) From wNomeLig Where Wnl_Identifiant=Wpl_IdentifiantWnl), "")');
  //T PETETIN le 27/08/2008 Dem n°3095
  ExecuteSQLNoPCL('UPDATE WFORMCONV SET WWF_ARRONDIRESULT=-1');
  // Init des champs WWS_ARRONDIRESUxx,WWS_COEFCONVxx,WWS_COEFCONVARRxx
  ExecuteSqlNoPCL('UPDATE WFORMCONVSAVE  '  +
  'SET WWS_ARRONDIRESU01=-1,WWS_ARRONDIRESU02=-1,WWS_ARRONDIRESU03=-1,WWS_ARRONDIRESU04=-1,WWS_ARRONDIRESU05=-1'+
  ',WWS_ARRONDIRESU06=-1,WWS_ARRONDIRESU07=-1,WWS_ARRONDIRESU08=-1,WWS_ARRONDIRESU09=-1'+
  ',WWS_COEFCONV01=1,WWS_COEFCONV02=1,WWS_COEFCONV03=1,WWS_COEFCONV04=1,WWS_COEFCONV05=1'+
  ',WWS_COEFCONV06=1,WWS_COEFCONV07=1,WWS_COEFCONV08=1,WWS_COEFCONV09=1'+
  ',WWS_COEFCONVARR01=1,WWS_COEFCONVARR02=1,WWS_COEFCONVARR03=1,WWS_COEFCONVARR04=1,WWS_COEFCONVARR05=1'+
  ',WWS_COEFCONVARR06=1,WWS_COEFCONVARR07=1,WWS_COEFCONVARR08=1,WWS_COEFCONVARR09=1 WHERE WWS_ARRONDIRESU01 IS NULL');

  AGLNettoieListes('GCSTKDISPO;GCSTKPREVISION;GCSTKDISPODETAIL;GCSTKPHYSIQUE','GA_FORMULEVAR',nil);
*)
end;

Procedure MajVer920;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //JLS Dem n°3134
  // Mise à jour du champ WOP_GUID
(*
  ExecuteSqlNoPCL('update Wordrephase set wop_guid = pgiguid');

  // Mise à jour WJA
  ExecuteSqlNoPCL('update wjournalaction set wja_guidori=ISNULL((select Max(wol_guid) from wordrelig '
                  +'where wol_identifiant=wja_identifiantwxx), "")'
                  + ' where wja_prefixe="WOL"')
                  ;

  // Mise à jour WJA
  ExecuteSqlNoPCL('update wjournalaction set wja_guidori=ISNULL((select Max(wop_guid) from wordrephase '
                  +'where wop_identifiant=wja_identifiantwxx), "")'
                  + ' where wja_prefixe="WOP"')
                  ;

  // Mise à jour du champ WOB_GUID
  ExecuteSqlNoPCL('update WordreBes set wob_guid = pgiguid');

  // Mise à jour WJA
  ExecuteSqlNoPCL('update wjournalaction set wja_guidori=ISNULL((select Max(wob_guid) from wordrebes '
                  +'where wob_identifiant=wja_identifiantwxx), "")'
                  + ' where wja_prefixe="WOB"')
                  ;

  // Mise à jour QWY
  ExecuteSqlNoPCL('update QWMESFICHESV set Qwy_GuidWob=ISNULL((select Max(wob_guid) from wordrebes '
                  +'where wob_identifiant=qwy_identifiantWob), "")');

  //JLS Dem n°3139
  // Mise à jour du champ WOG_GUID
  ExecuteSqlNoPCL('update Wordregamme set wog_guid = pgiguid');

  // Mise à jour WJA
  ExecuteSqlNoPCL('update wjournalaction set wja_guidori=ISNULL((select Max(wog_guid) from wordregamme '
                  +'where wog_identifiant=wja_identifiantwxx), "")'
                  + ' where wja_prefixe="WOG"')
                  ;

  // Mise à jour du champ WOR_GUID
  ExecuteSqlNoPCL('update Wordreres set wor_guid = pgiguid');

  // Mise à jour WJA
  ExecuteSqlNoPCL('update wjournalaction set wja_guidori=ISNULL((select Max(wor_guid) from wordreres '
                  +'where wor_identifiant=wja_identifiantwxx), "")'
                  + ' where wja_prefixe="WOR"')
                  ;
*)
end;

Procedure MajVer921;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //T PETETIN 3169
  //GP_20080903_TP_GP15155 >>>
  (*
  ExecuteSqlNoPCL('DELETE FROM PARAMSOC WHERE SOC_NOM="SO_WPLANLIVGRP1"');
  ExecuteSqlNoPCL('DELETE FROM PARAMSOC WHERE SOC_NOM="SCO_WPLANLIVGRPLIB"');
  ExecuteSQLNoPCL('UPDATE PARAMSOC SET SOC_DESIGN="A;0;;" || SUBSTRING(SOC_DESIGN, 6, 194) '
    +'WHERE SOC_NOM="SO_WPLANLIVRGRP2" AND SUBSTRING(SOC_DESIGN,1,5) = "A;1;;"');
  //GP_20080903_TP_GP15155 <<<

  //C PARWEZ 3189
  ExecuteSQLNoPCL('UPDATE WORDRELIG SET WOL_LIBELLE = ISNULL((SELECT GA_LIBELLE FROM ARTICLE '
    +'WHERE GA_ARTICLE = WOL_ARTICLE), "") WHERE WOL_LIBELLE = ""');
  *)
  //P JOUIN 3190
  ExecuteSQLContOnExcept('update TIERSCOMPL set YTC_AVECSOUSTRAIT="-" where ISNULL(YTC_AVECSOUSTRAIT,"")=""');

  //MC DESSEIGNET 3202
  // requête à afiare après mise à jour
  //dans les bases non réduites.
//  ExecuteSQLNoPCl ('update affaire set aff_txrealisation=0,aff_dteeffettxreal="'+UsDateTime_(IDate1900)+'"');     //regroupé en majver940

  //MC DESSEIGNET 3205
  // requête à afiare après mise à jour
  //dans les bases non réduites.
//  ExecuteSQLContOnExcept ('update ressource set ars_servcptaprinc1="",ars_servcptasec1="", '   // GM déplacé en 924
//  +'ars_servcptaprinc2="", ars_servcptasec2="",ars_dateservcptap2="'
//  +UsDateTime_(IDate1900)+'", ars_dateservcptas2="'+UsDateTime_(IDate1900)
//  +'" ,ars_principale1="-", ars_principale2="-", ars_principale3="-", ars_principale4="-", ars_datefinfonc3="'
//  +UsDateTime_(IDate2099)+'",ars_datefinfonc4="'+UsDateTime_(IDate2099)+'"');

  //MVG 3220
  // initialisation des champs créés dans la version 40
  ExecuteSQLContOnExcept ('update Immo SET I_AVOIR="-", I_OPEAVOIR="-", I_TAUXCO2=0')

end;

Procedure MajVer922;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //S MASSON 3218
  ExecuteSqlNoPcl('UPDATE DPORGA SET DOR_NOELEVAGE2=""');
  //MCD 3253
   //existe même en base réduite
//   ExecuteSqlContOnExcept ('update ressource set ars_codeforfait="" where ars_codeforfait is null');  déjà en 924

  //M FAUDEL 3256
  // Pour rétablir la valeur de PTE_STCRITERE2 qui était initialisé à tort à 0
//  ExecuteSqlContOnExcept('UPDATE TABLEDIMENT SET PTE_STCRITERE2 = "" WHERE PTE_STCRITERE2 = "0" AND PTE_NATURECRITERE2=""');

//MCD 3264
//  ExecuteSQLNoPcl ('UPDATE ACTIVITE SET ACT_LIBELLECOMPL = "", ACT_NUMINTER = 0, act_indice=0, ACT_ETATACTIVITE = ""');   déplacé en majver955
//  ExecuteSQLNoPcl('UPDATE EACTIVITE SET EAC_LIBELLECOMPL = "", EAC_NUMINTER = 0 ,eac_indice=0, EAC_ETATACTIVITE = ""');   déplacé en majver955

  //G JUGDE 3271
  ExecuteSQLNoPCL ( 'UPDATE QWHISTORES SET QWH_ORDREEPUR = 0 Where QWH_ORDREEPUR IS NULL');

  //M MORRETTON 3272
  (*
  ExecuteSQLNoPCL('DELETE FROM PARAMSOC WHERE SOC_NOM IN ("SCO_PDRDEFAUTGA","SCO_PDRMETHVALO","SCO_PDRGENEREWPL")');
  ExecuteSQLNoPCL('UPDATE WPDRLIG SET WPL_VALEURPDRENT=0, WPL_VALEURPDRSOR=0,'
    +'WPL_QTEENTREE=0, WPL_COUTENTREE=0, WPL_QTEPHASE=0, WPL_COUTPHASE=0, WPL_QTESORTIE=0, WPL_COUTSORTIE=0,'
    +'WPL_TYPEPDR=(SELECT MAX(WPE_TYPEPDR) FROM WPDRTET WHERE WPE_IDENTIFIANT=WPL_IDENTIFIANTWPE),'
    +'WPL_NATURETRAVAIL=(SELECT MAX(WPE_NATURETRAVAIL) FROM WPDRTET WHERE WPE_IDENTIFIANT=WPL_IDENTIFIANTWPE),'
    +'WPL_LIGNEORDRE=(SELECT MAX(WPE_LIGNEORDRE) FROM WPDRTET WHERE WPE_IDENTIFIANT=WPL_IDENTIFIANTWPE)'
    +'WHERE WPL_TYPEPDR IS NULL');
  *)
End;

Procedure MajVer923;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //MCD 3307   mcd : mise 100 au lieu de 0 au 28/10/2008
 // ExecuteSQLContOnExcept ('update ressource set ars_coefcorrfct1=100,ars_coefcorrfct2=100,ars_coefcorrfct3=100,ars_coefcorrfct4=100 ');  //GM déplacé en 924
  //P DUMET 3320
  ExecuteSQLContOnExcept ('update conventioncoll set PCV_ACTIVITE=""');
  //C DUMAS 3334
//  ExecuteSQLNoPCL ('UPDATE DEPOTS SET GDE_REGION = "" WHERE GDE_REGION IS NULL');
End;


Procedure MajVer924;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  // JLS-DS 24/09/2008
//  if not isDossierPCL then RepriseMotifsRebuts;

  //MCD 3344
  ExecuteSqlContOnExcept('UPDATE RESSOURCE SET ARS_CODEFORFAIT = "", ARS_SECTEURGEO="",'+
  'ars_servcptaprinc1="",ars_servcptasec1="",ars_servcptaprinc2="", ars_servcptasec2="",ars_dateservcptap2="'
  +UsDateTime_(IDate1900)+'", ars_dateservcptas2="'+UsDateTime_(IDate1900)
  +'" ,ars_principale1="-", ars_principale2="-", ars_principale3="-", ars_principale4="-", ars_datefinfonc3="'
  +UsDateTime_(IDate2099)+'",ars_datefinfonc4="'+UsDateTime_(IDate2099)+'",'
  + 'ars_coefcorrfct1=100,ars_coefcorrfct2=100,ars_coefcorrfct3=100,ars_coefcorrfct4=100 '
   );

  ExecutesqlcontOnExcept ('UPDATE RESSOURCE SET ARS_REGION = "" WHERE ARS_REGION IS NULL');
  ExecuteSqlContOnExcept('update fonction set AFO_LIBREFON1 = "",AFO_LIBREFON2 = "",AFO_LIBREFON3 = "",AFO_LIBREFON4 = "",AFO_LIBREFON5 = "", AFO_ORDREETAT =0');

  //T SUBLET 3378
  { Initialisation du champ WPF_METHPROD ajouté à la table WPARAMFONCTION }
//  ExecuteSqlNoPcl('UPDATE WPARAMFONCTION SET WPF_METHPROD="" WHERE WPF_METHPROD IS NULL');


End;

Procedure MajVer925;
var i : integer;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //P DUMET 3386
  ExecuteSqlContOnExcept('UPDATE DETGRPRUBLIEE SET PDG_COMMENTAIRE=""');


  //MC DESSEIGNET 3411
//  ExecuteSQLNoPCL('UPDATE AFMODELETACHE SET AFM_MOISDEBUT=0');

    //N CHAVANNE 3423
(*
    ExecuteSqlNoPcl('UPDATE LISTEINVLIG SET'
                  + ' GIL_PRIXINV=0,'
                  + ' GIL_MNTPHOTOINV=0'
                  + ' WHERE GIL_MNTPHOTOINV IS NULL');
*)
    //MC DESSEIGNET 3427
(*
    if not IsDossierPCL then
    begin
      for i := 1 to 5 do InsertChoixCode('ALF', 'TL' + intTostr(i), '.-Table libre fonction ' + intTostr(i), '', '');
      for i := 1 to 5 do InsertChoixCode('AFX', 'TL' + intTostr(i), '.-Table libre fonction prix ' + intTostr(i), '', '');
    end;
*)
    //D SCLAVOPOULOS 3431
//    ExecuteSqlNoPCL('UPDATE CHOIXEXT SET YX_LIBELLE="Temps d''exécution" WHERE YX_TYPE="WTR" AND YX_CODE="G02"');
//GP_20100517_DKZ_GP17440
//    ExecuteSqlNoPCL('UPDATE CHOIXEXT SET YX_LIBELLE="Perte proportionnelle" WHERE YX_TYPE="WTR" AND YX_LIBELLE="Perte proportionnelle"');

    //D SCLAVOPOULOS 3432
    ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Libellé de la condition 2"  WHERE DH_NOMCHAMP="WIT_C2LIBELLE"');
    ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Libellé d''opération2" WHERE DH_NOMCHAMP="WGO_LIBELLE"');

    //D SCLAVOPOULOS 3438
    ExecuteSQLContOnExcept('DELETE FROM MODELES WHERE MO_TYPE="E" AND MO_NATURE="WO4" AND MO_CODE="OL0"');
    ExecuteSQLContOnExcept('DELETE FROM MODEDATA WHERE MD_CLE LIKE "EWO4OL0%"');

    //S BOUSSERT 3443
    // Initialisation des nouveaux champs de la tables CBALSIT
    ExecuteSQLContOnExcept('UPDATE CBALSIT SET BSI_SUPERVISE="-", BSI_HISTORISE="-"');
End;

Procedure MajVer926;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  // T SUBLET 3365
  // GP_20080929_TS : FQ034;13889
  ExecuteSqlNoPCL('DELETE FROM MENU WHERE MN_1=215 AND MN_2=7 AND MN_3=11 AND MN_4>0');

  // S BOUSSERT 3443
  // Initialisation des nouveaux champs de la tables CBALSIT
  ExecuteSQLContOnExcept('UPDATE CBALSIT SET BSI_SUPERVISE="-", BSI_HISTORISE="-" WHERE BSI_SUPERVISE IS NULL');
End;

Procedure MajVer927;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
//ATTENTION INTEGRATION EN 927 DES ELEMENTS DE STRCUTURE KPMG MULTI ENTITES
  ExecuteSQLContOnExcept('UPDATE ANAGUI SET AG_ENTITY = 0  				  WHERE AG_ENTITY 	  IS NULL');
  ExecuteSQLContOnExcept('UPDATE ANAGUI SET AG_SOUSPLAN1 = ""  			WHERE AG_SOUSPLAN1 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE ANAGUI SET AG_SOUSPLAN2 = ""  			WHERE AG_SOUSPLAN2 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE ANAGUI SET AG_SOUSPLAN3 = ""  			WHERE AG_SOUSPLAN3 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE ANAGUI SET AG_SOUSPLAN4 = ""  			WHERE AG_SOUSPLAN4 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE ANAGUI SET AG_SOUSPLAN5 = ""  			WHERE AG_SOUSPLAN5 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE ANAGUI SET AG_SOUSPLAN6 = ""  			WHERE AG_SOUSPLAN6 	IS NULL');

// GPIOT 10110
//  ExecuteSQLContOnExcept('UPDATE ANALYTIQ SET Y_ENTITY = 0  				WHERE Y_ENTITY 		IS NULL');
  UpDateDecoupeAna('Y_ENTITY = 0','','','Y_ENTITY 		IS NULL') ;

// G PIOT : disparition des champs au profit de Y_REFGUID
//  ExecuteSQLContOnExcept('UPDATE ANALYTIQ SET Y_ID = 0  				WHERE Y_ID 		IS NULL');

  ExecuteSQLContOnExcept('UPDATE BANQUECP SET BQ_AUXILIAIRE = ""  			WHERE BQ_AUXILIAIRE 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE CBALSIT SET BSI_ENTITY = 0 				WHERE BSI_ENTITY 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE CBALSITECR SET BSE_ENTITY = 0 			WHERE BSE_ENTITY 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE CCONDREMISE SET CCB_AUXILIAIRE = ""  		WHERE CCB_AUXILIAIRE 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE CEDTBALANCE SET CED_ENTITY = 0    			WHERE CED_ENTITY 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE CEDTBALBUD SET  CEB_ENTITY  = 0 			WHERE CEB_ENTITY 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE CENGAGEMENT SET CEN_ENTITY = 0  			WHERE CEN_ENTITY 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE CENGAGEMENT SET CEN_ID = 0  				WHERE CEN_ID 		IS NULL');

  ExecuteSQLContOnExcept('UPDATE CETEBAC SET CET_ENTITY = 0  				WHERE CET_ENTITY 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE CFRAISREMISE SET CFR_AUXILIAIRE  =  ""	 	WHERE CFR_AUXILIAIRE 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE CLEREPAR SET RE_INVISIBLE = "-"  			WHERE RE_INVISIBLE 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE CMODELRESTANADET SET CAD_VALEURDEF1 = ""  	WHERE CAD_VALEURDEF1 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE CMODELRESTANADET SET CAD_VALEURDEF2 = ""  	WHERE CAD_VALEURDEF2 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE CMODELRESTANADET SET CAD_VALEURDEF3 = ""  	WHERE CAD_VALEURDEF3 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE CMODELRESTANADET SET CAD_VALEURDEF4 = ""  	WHERE CAD_VALEURDEF4 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE CMODELRESTANADET SET CAD_VALEURDEF5 = ""  	WHERE CAD_VALEURDEF5 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE CMODELRESTANADET SET CAD_VALEURDEF6 = ""  	WHERE CAD_VALEURDEF6 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE CONTABON SET  CB_ENTITY = 0  			WHERE CB_ENTITY 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE CORRESP SET   CR_INVISIBLE = "-"  			WHERE CR_INVISIBLE  	IS NULL');

  ExecuteSQLContOnExcept('UPDATE CPARAMGENER SET CPG_AUXILIAIRE = ""   		WHERE CPG_AUXILIAIRE 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE CPARAMGENER SET CPG_AUXILIAIRESEL = ""  		WHERE CPG_AUXILIAIRESEL IS NULL');

  ExecuteSQLContOnExcept('UPDATE CPBONSAPAYER SET BAP_ENTITY = 0   			WHERE BAP_ENTITY 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE CPMPTEMPOR SET CTT_ENTITY  = 0  			WHERE CTT_ENTITY 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE CRELBQE SET CRL_ENTITY = 0  				WHERE CRL_ENTITY 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE CRELBQE SET CRL_AUXILIAIREBQE = ""   		WHERE CRL_AUXILIAIREBQE IS NULL');

  ExecuteSQLContOnExcept('UPDATE CRELBQEANALYTIQ SET CRY_ENTITY = 0  		WHERE CRY_ENTITY 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE CSECTION SET CSP_TRANCHEGENEDE = "" 		WHERE  CSP_TRANCHEGENEDE IS NULL');
  ExecuteSQLContOnExcept('UPDATE CSECTION SET CSP_TRANCHEGENEA = ""  		WHERE CSP_TRANCHEGENEA 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE CSECTION SET CSP_INDIRECTE = "-"  			WHERE CSP_INDIRECTE 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE CSECTION SET CSP_UO = 0  				WHERE CSP_UO 		IS NULL');
  ExecuteSQLContOnExcept('UPDATE CSECTION SET CSP_UOLIBELLE = ""  			WHERE CSP_UOLIBELLE 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE CSECTION SET CSP_INVISIBLE = "-"   		WHERE CSP_INVISIBLE 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE DOCREGLE SET     DR_BANQUEAUXPREVI = ""  	WHERE DR_BANQUEAUXPREVI IS NULL');
  ExecuteSQLContOnExcept('UPDATE DOCREGLE SET     DR_CPTAUXBANQUE   = ""  	WHERE DR_CPTAUXBANQUE 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE ECRCOMPL SET EC_ENTITY = 0   			WHERE EC_ENTITY 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE ECRCOMPL SET EC_ID  = 0  				WHERE EC_ID 		IS NULL');

  ExecuteSQLContOnExcept('UPDATE ECRGUI SET EG_ENTITY = 0   				WHERE EG_ENTITY 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE ECRGUI SET EG_BANQUEAUXPREVI = ""  		WHERE EG_BANQUEAUXPREVI IS NULL');

// GPIOT 10110
//  ExecuteSQLContOnExcept('UPDATE ECRITURE SET E_ENTITY = 0  				WHERE E_ENTITY 		IS NULL');
  UpDateDecoupeEcr('E_ENTITY = 0','','','E_ENTITY 		IS NULL') ;

// G PIOT : disparition des champs au profit de Y_REFGUID
//  ExecuteSQLContOnExcept('UPDATE ECRITURE SET E_ID = 0  				WHERE E_ID 		IS NULL');
// GPIOT 10110
//  ExecuteSQLContOnExcept('UPDATE ECRITURE SET E_DATEORIGINE = "'+UsDateTime_(idate1900) +'" 	WHERE E_DATEORIGINE 	IS NULL');
  UpDateDecoupeEcr('E_DATEORIGINE = "'+UsDateTime_(idate1900) +'"','','','E_DATEORIGINE 	IS NULL') ;

  ExecuteSQLContOnExcept('UPDATE EEXBQ SET EE_ENTITY = 0  				WHERE EE_ENTITY 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE EEXBQ SET EE_AUXILIAIRE = ""  			WHERE EE_AUXILIAIRE 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE EEXBQLIG SET CEL_ENTITY = 0  			WHERE CEL_ENTITY 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE EEXBQLIG SET CEL_AUXILIAIRE = ""  			WHERE CEL_AUXILIAIRE 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE ETAPEREG SET ER_CPTEAUXDEPART = ""   		WHERE ER_CPTEAUXDEPART IS NULL');
  ExecuteSQLContOnExcept('UPDATE ETAPEREG SET ER_CPTEAUXARRIVE = ""   		WHERE ER_CPTEAUXARRIVE IS NULL');

  ExecuteSQLContOnExcept('UPDATE GENERAUX SET G_VENTILTYPE = ""  			WHERE G_VENTILTYPE 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE GENERAUX SET G_RESTRICTIONA1 = ""  		WHERE G_RESTRICTIONA1 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE GENERAUX SET G_RESTRICTIONA2 = ""  		WHERE G_RESTRICTIONA2 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE GENERAUX SET G_RESTRICTIONA3 = ""    		WHERE G_RESTRICTIONA3 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE GENERAUX SET G_RESTRICTIONA4 = ""   		WHERE G_RESTRICTIONA4 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE GENERAUX SET G_RESTRICTIONA5 = ""   		WHERE G_RESTRICTIONA5 	IS NULL');
  ExecuteSQLContOnExcept('UPDATE GENERAUX SET G_INVISIBLE = "-" 			WHERE G_INVISIBLE 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE GUIDE SET GU_ENTITY = 0  				WHERE GU_ENTITY 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE HISTOBAL SET HB_ENTITY  = 0  			WHERE HB_ENTITY 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE JOURNAL SET J_CONTREPARTIEAUX = "",J_INVISIBLE="-" 		WHERE J_CONTREPARTIEAUX IS NULL');

  ExecuteSQLContOnExcept('UPDATE NATCPTE SET NT_INVISIBLE = "-" 			WHERE NT_INVISIBLE 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE RELANCE SET RR_INVISIBLE = "-"  			WHERE RR_INVISIBLE 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE SECTION SET  S_INVISIBLE = "-"   			WHERE S_INVISIBLE 	IS NULL');



  ExecuteSQLContOnExcept('UPDATE SOUCHE SET SH_ENTITY = 0  				WHERE SH_ENTITY 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE ETABLISS SET ET_INVISIBLE = "-"  			WHERE ET_INVISIBLE 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE MODEREGL SET MR_INVISIBLE = "-"  			WHERE MR_INVISIBLE 	IS NULL');

  ExecuteSQLContOnExcept('UPDATE TIERS SET T_INVISIBLE = "-"  			WHERE T_INVISIBLE 	IS NULL');

  // S BOUSSERT 3483
  // Correction des TAG sur concept de trésorerie utilisant par erreur des TAG du bureau expert sur les dossiers clients (expert ou entreprise)
  ExecuteSQLContOnExcept('UPDATE MENU SET MN_TAG=26059 WHERE MN_1=26 and MN_2=7 and MN_3 = 1 and MN_4 = 1 and MN_tag = 26053');
  ExecuteSQLContOnExcept('UPDATE MENU SET MN_TAG=26060 WHERE MN_1=26 and MN_2=7 and MN_3 = 1 and MN_4 = 1 and MN_tag = 26055');

  // JL SAUZET 3486
  // Mise à jour
//  ExecuteSqlNoPcl('UPDATE STKFORMULE SET GSF_MODIFIABLE=LEFT(GSF_LIBREART21,1) WHERE GSF_MODIFIABLE IS NULL');

  // D SCLAVOPOULOS 3492
  (*
  ExecuteSQLNoPCL('UPDATE WORDRELAS SET WLS_BOOLLIBREWOL1="-",WLS_BOOLLIBREWOL2="-",WLS_BOOLLIBREWOL3="-"'
                    +',WLS_LIBREWOL1="",WLS_LIBREWOL2="",WLS_LIBREWOL3="",WLS_LIBREWOL4="",WLS_LIBREWOL5=""'
                    +',WLS_LIBREWOL6="",WLS_LIBREWOL7="",WLS_LIBREWOL8="",WLS_LIBREWOL9="",WLS_LIBREWOLA=""'
                    +',WLS_BOOLLIBREPF1="-",WLS_BOOLLIBREPF2="-",WLS_BOOLLIBREPF3="-"'
                    +',WLS_LIBREARTPF1="",WLS_LIBREARTPF2="",WLS_LIBREARTPF3="",WLS_LIBREARTPF4="",WLS_LIBREARTPF5=""'
                    +',WLS_LIBREARTPF6="",WLS_LIBREARTPF7="",WLS_LIBREARTPF8="",WLS_LIBREARTPF9="",WLS_LIBREARTPFA=""'
                    +',WLS_COLLECTIONPF="",WLS_COLLECTIONCOMP="",WLS_TIERS="",WLS_AFFAIRE=""'
                    +',WLS_BOOLLIBRECOMP1="-",WLS_BOOLLIBRECOMP2="-",WLS_BOOLLIBRECOMP3="-"'
                    +',WLS_LIBRECOMP1="",WLS_LIBRECOMP2="",WLS_LIBRECOMP3="",WLS_LIBRECOMP4="",WLS_LIBRECOMP5=""'
                    +',WLS_LIBRECOMP6="",WLS_LIBRECOMP7="",WLS_LIBRECOMP8="",WLS_LIBRECOMP9="",WLS_LIBRECOMPA=""'
                    +',WLS_BOOLLIBRETIE1="-",WLS_BOOLLIBRETIE2="-",WLS_BOOLLIBRETIE3="-"'
                    +',WLS_LIBRETIERS1="",WLS_LIBRETIERS2="",WLS_LIBRETIERS3="",WLS_LIBRETIERS4="",WLS_LIBRETIERS5=""'
                    +',WLS_LIBRETIERS6="",WLS_LIBRETIERS7="",WLS_LIBRETIERS8="",WLS_LIBRETIERS9="",WLS_LIBRETIERSA=""'
                    +',WLS_NATUREPIECEG="",WLS_SOUCHE="", WLS_NUMERO=0,WLS_NUMERO1=0,WLS_NUMORDRE=0,WLS_NUMORDRE1=0'
                    +' WHERE WLS_BOOLLIBREWOL1 IS NULL');
   *)
  // MC DESSEIGNET 3499
   //GIGA CB Le typeplanif est maintenant obligatoire si articlecompl exist
//   ExecuteSQLNoPCL ('UPDATE ARTICLECOMPL SET GA2_TYPEPLANIF = "NON" WHERE GA2_TYPEPLANIF = "" OR GA2_TYPEPLANIF IS NULL ');
End;

Procedure MajVer928;
var i: integer;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  // DKZ: Mise à jour de la nouvelle structure WGP
//  wGammePhaseEtFreinte;

  //C DUMAS 3527
  if not IsDossierPCL then
    begin
    for i := 1 to 3 do InsertChoixCode('ZLD', 'DB' + intTostr(i), '.-Décision libre ' + intTostr(i), '', '');
    for i := 1 to 3 do InsertChoixCode('ZLD', 'DC' + intTostr(i), '.-Texte libre ' + intTostr(i), '', '');
    for i := 1 to 3 do InsertChoixCode('ZLD', 'DD' + intTostr(i), '.-Date libre ' + intTostr(i), '', '');
    for i := 1 to 3 do InsertChoixCode('ZLD', 'DM' + intTostr(i), '.-Montant libre ' + intTostr(i), '', '');
    for i := 1 to 9 do InsertChoixCode('ZLD', 'DT' + intTostr(i), '.-Table libre ' + intTostr(i), '', '');
    InsertChoixCode('ZLD', 'DTA', '.-Table libre 10', '', '');
  end;

  //MD3 3528
  // Reporte l'info quand c'est possible
  ExecuteSQLContOnExcept('UPDATE DOSSIER SET DOS_NONTRAITE=(SELECT DOR_NONTRAITE FROM DPORGA WHERE DOR_GUIDPER=DOS_GUIDPER) WHERE DOS_NONTRAITE IS NULL');
  // Complète par une valeur par défaut
  ExecuteSQLContOnExcept('UPDATE DOSSIER SET DOS_NONTRAITE="-" WHERE DOS_NONTRAITE IS NULL');

 //D SCLAVOPOULOS 3531
 (*
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="GQ" AND DH_CONTROLE NOT LIKE "%Q%"'
                       + ' AND (DH_NOMCHAMP="GQ_STOCKMIN" OR DH_NOMCHAMP="GQ_STOCKMAX" OR DH_NOMCHAMP="GQ_PHYSIQUE" OR DH_NOMCHAMP="GQ_RESERVECLI"'
                       + ' OR DH_NOMCHAMP="GQ_RESERVEFOU" OR DH_NOMCHAMP="GQ_PREPACLI" OR DH_NOMCHAMP="GQ_PREPAFOU"'
                       + ' OR DH_NOMCHAMP="GQ_LIVRECLIENT" OR DH_NOMCHAMP="GQ_LIVREFOU" OR DH_NOMCHAMP="GQ_TRANSFERT" OR DH_NOMCHAMP="GQ_QTE1"'
                       + ' OR DH_NOMCHAMP="GQ_QTE2" OR DH_NOMCHAMP="GQ_QTE3" OR DH_NOMCHAMP="GQ_STOCKINITIAL" OR DH_NOMCHAMP="GQ_CUMULSORTIES"'
                       + ' OR DH_NOMCHAMP="GQ_CUMULENTREES" OR DH_NOMCHAMP="GQ_VENTEFFO" OR DH_NOMCHAMP="GQ_ENTREESORTIES" OR DH_NOMCHAMP="GQ_ECARTINV"'
                       + ' OR DH_NOMCHAMP="GQ_QRUPSTOC" OR DH_NOMCHAMP="GQ_TAILLELOT" OR DH_NOMCHAMP="GQ_MINIFAB" OR DH_NOMCHAMP="GQ_TRANSFERTRECU"'
                       + ' OR DH_NOMCHAMP="GQ_CONSOMMATION" OR DH_NOMCHAMP="GQ_PRODUCTION" OR DH_NOMCHAMP="GQ_BLOCAGE" OR DH_NOMCHAMP="GQ_AFFECTE"'
                       + ' OR DH_NOMCHAMP="GQ_PREAFFECTE" OR DH_NOMCHAMP="GQ_STOCKALERTE" OR DH_NOMCHAMP="GQ_STOCKRECOMPL")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="GQC" AND DH_CONTROLE NOT LIKE "%Q%"'
                       + ' AND (DH_NOMCHAMP="GQC_PHYSIQUE" OR DH_NOMCHAMP="GQC_RESERVECLI" OR DH_NOMCHAMP="GQC_RESERVEFOU" OR DH_NOMCHAMP="GQC_PREPACLI")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="GQD" AND DH_CONTROLE NOT LIKE "%Q%" AND DH_NOMCHAMP="GQD_PHYSIQUE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="GQL" AND DH_CONTROLE NOT LIKE "%Q%" AND DH_NOMCHAMP="GQL_PHYSIQUE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="GSM" AND DH_CONTROLE NOT LIKE "%Q%"'
                       + ' AND (DH_NOMCHAMP="GSM_QPREVUE" OR DH_NOMCHAMP="GSM_QPREPA" OR DH_NOMCHAMP="GSM_QRUPTURE" OR DH_NOMCHAMP="GSM_PHYSIQUE" OR DH_NOMCHAMP="GSM_QTEFACT")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="GST" AND DH_CONTROLE NOT LIKE "%Q%"'
                      + ' AND (DH_NOMCHAMP="GST_PHYSIQUE" OR DH_NOMCHAMP="GST_DPA" OR DH_NOMCHAMP="GST_DPR")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WAD" AND DH_CONTROLE NOT LIKE "%Q%" AND DH_NOMCHAMP="WAD_QLIENSAIS"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WBE" AND DH_CONTROLE NOT LIKE "%Q%" AND DH_NOMCHAMP="WBE_QUANTITE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WDK" AND DH_CONTROLE NOT LIKE "%Q%" AND DH_NOMCHAMP="WDK_VALEURCARA"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WEV" AND DH_CONTROLE NOT LIKE "%Q%"'
                      + ' AND (DH_NOMCHAMP="WEV_QMOUSAIS" OR DH_NOMCHAMP="WEV_QMOUSTOC" OR DH_NOMCHAMP="WEV_QEVOLUTIONSTOC" OR DH_NOMCHAMP="WEV_QEVOLDEPOTSTOC"'
                      + ' OR DH_NOMCHAMP="WEV_COUVENCN" OR DH_NOMCHAMP="WEV_QUOTAPCTDEM" OR DH_NOMCHAMP="WEV_QUOTAPCTAPP")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WGL" AND DH_CONTROLE NOT LIKE "%Q%"'
                      + ' AND (DH_NOMCHAMP="WGL_QLOTSAIS" OR DH_NOMCHAMP="WGL_QLOTSTOC")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WGT" AND DH_CONTROLE NOT LIKE "%Q%"'
                      + ' AND (DH_NOMCHAMP="WGT_QLOTSAIS" OR DH_NOMCHAMP="WGT_QLOTSTOC")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WJA" AND DH_CONTROLE NOT LIKE "%Q%"'
                      + ' AND (DH_NOMCHAMP="WJA_QSAIS" OR DH_NOMCHAMP="WJA_QCALC")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WNC" AND DH_CONTROLE NOT LIKE "%Q%" AND DH_NOMCHAMP="WNC_QLIENSAIS"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WND" AND DH_CONTROLE NOT LIKE "%Q%"'
                      + ' AND (DH_NOMCHAMP="WND_QBESSAIS" OR DH_NOMCHAMP="WND_QBESSTOC" OR DH_NOMCHAMP="WND_QBPPROPSAIS" OR DH_NOMCHAMP="WND_QBPPROPSTOC"'
                      + ' OR DH_NOMCHAMP="WND_QBPPERIODESAIS" OR DH_NOMCHAMP="WND_QBPPERIODESTOC" OR DH_NOMCHAMP="WND_QBPFIXESAIS"'
                      + ' OR DH_NOMCHAMP="WND_QBPFIXESTOC" OR DH_NOMCHAMP="WND_QBPFREINTESAIS" OR DH_NOMCHAMP="WND_QBPFREINTESTOC"'
                      + ' OR DH_NOMCHAMP="WND_QLIENSAIS" OR DH_NOMCHAMP="WND_QLIENSTOC" OR DH_NOMCHAMP="WND_QPPERSAIS" OR DH_NOMCHAMP="WND_QPPERSTOC"'
                      + ' OR DH_NOMCHAMP="WND_QPERIODESAIS" OR DH_NOMCHAMP="WND_QPERIODESTOC" OR DH_NOMCHAMP="WND_QPFIXESAIS" OR DH_NOMCHAMP="WND_QPFIXESTOC")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WNL" AND DH_CONTROLE NOT LIKE "%Q%" AND '
                      + ' (DH_NOMCHAMP="WNL_QLOTSAIS" OR DH_NOMCHAMP="WNL_QLIENSAIS" OR DH_NOMCHAMP="WNL_QLIENSTOC" OR DH_NOMCHAMP="WNL_QPFIXESAIS" OR DH_NOMCHAMP="WNL_QPFIXESTOC"'
                      + ' OR DH_NOMCHAMP="WNL_QPPERSAIS" OR DH_NOMCHAMP="WNL_QPPERSTOC" OR DH_NOMCHAMP="WNL_QPERIODESAIS" OR DH_NOMCHAMP="WNL_QPERIODESTOC")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WNS" AND DH_CONTROLE NOT LIKE "%Q%"'
                      + ' AND (DH_NOMCHAMP="WNS_QLIENSAIS" OR DH_NOMCHAMP="WNS_QLIENSTOC")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WNT" AND DH_CONTROLE NOT LIKE "%Q%"'
                      + ' AND (DH_NOMCHAMP="WNT_QLOTSAIS" OR DH_NOMCHAMP="WNT_QLOTSTOC")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WOB" AND DH_CONTROLE NOT LIKE "%Q%"'
                      + ' AND (DH_NOMCHAMP="WOB_QLIENSAIS" OR DH_NOMCHAMP="WOB_QLIENSTOC" OR DH_NOMCHAMP="WOB_QBESSAIS" OR DH_NOMCHAMP="WOB_QBESSTOC"'
                      + ' OR DH_NOMCHAMP="WOB_QRUPSAIS" OR DH_NOMCHAMP="WOB_QRUPSTOC" OR DH_NOMCHAMP="WOB_QAFFSAIS" OR DH_NOMCHAMP="WOB_QAFFSTOC"'
                      + ' OR DH_NOMCHAMP="WOB_QCONSAIS" OR DH_NOMCHAMP="WOB_QCONSTOC" OR DH_NOMCHAMP="WOB_QTSTSAIS" OR DH_NOMCHAMP="WOB_QTSTSTOC"'
                      + ' OR DH_NOMCHAMP="WOB_QLASSAIS" OR DH_NOMCHAMP="WOB_QACCSAISWOP" OR DH_NOMCHAMP="WOB_QARECSAISWOP" OR DH_NOMCHAMP="WOB_QLASSTOC"'
                      + ' OR DH_NOMCHAMP="WOB_QPPERSAIS" OR DH_NOMCHAMP="WOB_QPPERSTOC" OR DH_NOMCHAMP="WOB_QPERIODESAIS" OR DH_NOMCHAMP="WOB_QPERIODESTOC"'
                      + ' OR DH_NOMCHAMP="WOB_QPFIXESAIS" OR DH_NOMCHAMP="WOB_QPFIXESTOC")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WOG" AND DH_CONTROLE NOT LIKE "%Q%"'
                      + ' AND (DH_NOMCHAMP="WOG_QACCSAISWOP" OR DH_NOMCHAMP="WOG_QACCSAIS" OR DH_NOMCHAMP="WOG_QACCSTOC" OR DH_NOMCHAMP="WOG_QPRODSAIS"'
                      + ' OR DH_NOMCHAMP="WOG_QPRODSTOC" OR DH_NOMCHAMP="WOG_QSUSPSAIS" OR DH_NOMCHAMP="WOG_QSUSPSTOC" OR DH_NOMCHAMP="WOG_QREBUSAIS"'
                      + ' OR DH_NOMCHAMP="WOG_QREBUSTOC" OR DH_NOMCHAMP="WOG_QBONNSAIS" OR DH_NOMCHAMP="WOG_QBONNSTOC")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WOL" AND DH_CONTROLE NOT LIKE "%Q%"'
                      + ' AND (DH_NOMCHAMP="WOL_QDEMSAIS" OR DH_NOMCHAMP="WOL_QDEMSTOC" OR DH_NOMCHAMP="WOL_QACCSAIS" OR DH_NOMCHAMP="WOL_QACCSTOC"'
                      + ' OR DH_NOMCHAMP="WOL_QLANSAIS" OR DH_NOMCHAMP="WOL_QLANSTOC" OR DH_NOMCHAMP="WOL_QRECSAIS" OR DH_NOMCHAMP="WOL_QRECSTOC"'
                      + ' OR DH_NOMCHAMP="WOL_QREBUSAIS" OR DH_NOMCHAMP="WOL_QREBUSTOC" OR DH_NOMCHAMP="WOL_QARECSAIS" OR DH_NOMCHAMP="WOL_QARECSTOC")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WOP" AND DH_CONTROLE NOT LIKE "%Q%"'
                      + ' AND (DH_NOMCHAMP="WOP_QACCSAIS" OR DH_NOMCHAMP="WOP_QACCSTOC" OR DH_NOMCHAMP="WOP_QLANSAIS" OR DH_NOMCHAMP="WOP_QLANSTOC"'
                      + ' OR DH_NOMCHAMP="WOP_QARECSAIS" OR DH_NOMCHAMP="WOP_QARECSTOC" OR DH_NOMCHAMP="WOP_QRECSAIS" OR DH_NOMCHAMP="WOP_QRECSTOC"'
                      + ' OR DH_NOMCHAMP="WOP_QSUSPSAIS" OR DH_NOMCHAMP="WOP_QSUSPSTOC" OR DH_NOMCHAMP="WOP_QREBUSAIS" OR DH_NOMCHAMP="WOP_QREBUSTOC")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WOR" AND DH_CONTROLE NOT LIKE "%Q%"'
                      + ' AND (DH_NOMCHAMP="WOR_QPRODSAIS" OR DH_NOMCHAMP="WOR_QPRODSTOC")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WPD" AND DH_CONTROLE NOT LIKE "%Q%" AND DH_NOMCHAMP="WPD_QUANTITE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WPE" AND DH_CONTROLE NOT LIKE "%Q%"'
                      + ' AND (DH_NOMCHAMP="WPE_QDEM" OR DH_NOMCHAMP="WPE_QECOFAB" OR DH_NOMCHAMP="WPE_QVALEUR")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WPL" AND DH_CONTROLE NOT LIKE "%Q%"'
                      + ' AND (DH_NOMCHAMP="WPL_QTEENTREE" OR DH_NOMCHAMP="WPL_QTEPHASE" OR DH_NOMCHAMP="WPL_QTESORTIE"'
                      + ' OR DH_NOMCHAMP="WPL_QSANSPERTE" OR DH_NOMCHAMP="WPL_QAVECPERTE")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WPN" AND DH_CONTROLE NOT LIKE "%Q%"'
                      + ' AND (DH_NOMCHAMP="WPN_QLIENSAIS" OR DH_NOMCHAMP="WPN_QTE")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WSE" AND DH_CONTROLE NOT LIKE "%Q%" AND DH_NOMCHAMP="WSE_QUANTITE"');

  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE="LD" WHERE DH_NOMCHAMP="RQG_REFORIGINE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE="L" WHERE DH_NOMCHAMP="DSO_TXTRANSPORT"');
  *)
  //JTR 3534
  { GC_20081027_JTR_010;13258 }
  ExecuteSQLNoPCL('UPDATE TIERSIMPACTPIECE SET GTI_TYPENATTIERS = "FAC" WHERE GTI_ELEMENTFC = "TXN" AND GTI_TYPENATTIERS = ""');

  //D SCLAVOPOULOS 3537
//  ExecuteSQLContOnExcept('DELETE FROM PARAMSOC WHERE SOC_NOM="SCO_WLOTRECEPWOL"');
//  ExecuteSQLNoPCL('UPDATE WPARAMFONCTION SET WPF_COMBO01=iif(WPF_BOOLEAN01="X", "003", "001") WHERE WPF_CODEFONCTION = "REPRISELOT"');
//  AGLNettoieListesPlus('WPF_REPRISELOT', 'WPF_COMBO01', nil, True, 'WPF_BOOLEAN01');

end;

Procedure MajVer929;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
 //C DUMAS 3550
 //FQ 15632 :
 ExecuteSqlContOnExcept ('UPDATE CODEPOST SET O_PAYS= "FRA" WHERE O_PAYS IS NULL OR O_PAYS=""');
 ExecuteSqlContOnExcept ('UPDATE CODEPOST SET O_CODEINSEE = ""  WHERE O_CODEINSEE IS NULL');

 //D SCLAVOPOULOS 3583
 ExecuteSQLContOnExcept('DELETE FROM MODELES WHERE MO_TYPE="E" AND MO_NATURE="WO4" AND MO_CODE="OL0"');
 ExecuteSQLContOnExcept('DELETE FROM MODEDATA WHERE MD_CLE LIKE "EWO4OL0%"');

 //B MERIAUX 3586
 if IsMonoOuCommune then
 begin
   ExecuteSQL('update juevenement set jev_gedguid = ""');
   ExecuteSQL('update judosopact set joa_gedguid = ""');
 end;
 // DBR 3606
 ExecuteSQLContOnExcept ('update dechamps set dh_controle="LDZ1234" where dh_nomchamp="T_LOCALTAX"');
end;

Procedure MajVer930;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //M DESGOUTTE 3623
  if IsMonoOuCommune then
   ExecuteSQLContOnExcept('UPDATE DOSSIER SET DOS_STATUTEDIFIS="", DOS_STATUTEDISOC="", DOS_STATUTEDIJUR=""');
end;

Procedure MajVer931;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //DBR 3636
  //Mise à jour des contextes des natures de pièces
//  ExecuteSQLNoPCL ('update parpiece set gpp_contextes="GC;" || gpp_contextes where gpp_naturepieceg in ("AVC", "BLF", "CMC", "CMF", "DEF", "FAC")');
//  ExecuteSQLNoPCL ('update parpiece set gpp_contextes="INU;" where gpp_naturepieceg in ("CFR", "CCR", "LCR", "LFR")');

End;

Procedure MajVer932;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
//  if not IsDossierPCL() then
//    EdiMajVer(932).MajApres();

  //N FOURNEL 3679
//  AglNettoieListes('GCGROUPEPIECE', 'GP_DOMAINE',nil);
  //D SCLAVOOPOULOS 3693
//  ExecuteSQLNoPCL( 'UPDATE WCHANGEMENT SET WCH_GUID = PGIGUID, WCH_GUIDORI = "" WHERE WCH_GUID IS NULL' );
  //JTR 3695
  { GC_20081119_JTR_010;13258 }
  ExecuteSQLNoPCL('UPDATE TIERSIMPACTPIECE SET GTI_TYPENATTIERS = "FAC" WHERE GTI_ELEMENTFC = "TXN" AND GTI_TYPENATTIERS = ""');
  //JTR 3701
  { JTR - Suite alignement ERA }
//  ExecuteSQLNoPCL('UPDATE PIEDBASE SET GPB_LIBELLETAXE = "", GPB_COMPTETAXE = "", GPB_NUMTAXE = 0 WHERE GPB_COMPTETAXE IS NULL');
  // F BERGER 3717
  ExecuteSqlContOnExcept('DELETE from MODELES where MO_NATURE = "US0" and MO_CODE IN ("ST0","ST1","ST4")');
  ExecuteSqlContOnExcept('DELETE from MODEDATA where MD_CLE like "EUS0ST0%" or MD_CLE like "EUS0ST1%" or MD_CLE like "EUS0ST4%"');
  //B MERIAUX 3731
  ExecuteSqlContOnExcept('update juridique set jur_socliquid = "-", jur_cotebourse = "-", jur_datecotebourse = "' + UsDateTime_(iDate1900) + '"');
  //D SCLAVOPOULOS 3740
//  ExecuteSQLNoPCL('UPDATE WNOMELIG SET WNL_MODEAPPRO="DEB" WHERE WNL_MODEAPPRO IS NULL') ;
//  ExecuteSQLNoPCL('UPDATE WORDREBES SET WOB_MODEAPPRO="DEB" WHERE WOB_MODEAPPRO IS NULL');
//  ExecuteSQLNoPCL('DELETE FROM WNOMEDEC WHERE WND_CONTEXTE LIKE "CBN;%" OR WND_CONTEXTE LIKE "CBB;%"');
//  ExecuteSQLNoPCL('DELETE FROM QNOMENCLA WHERE QNO_CTX IN (SELECT QSM_CODESIMU FROM QSIMULATION WHERE QSM_CTX="0" AND QSM_MODESIMU IN ("2","3"))');

  //M GUERIN 3743
  {Mis à vide des champs à Null}
  ExecuteSQLNoPCL('update EPERSPECTIVES set EEP_FICHIERORIG="" where ISNULL(EEP_FICHIERORIG,"")=""');
  ExecuteSQLNoPCL('update EOPERATIONS set ERO_FICHIERORIG="" where ISNULL(ERO_FICHIERORIG,"")=""');
  ExecuteSQLNoPCL('update EOPERATIONS set ERO_CODEIMPORT="" where ISNULL(ERO_CODEIMPORT,"")=""');
  ExecuteSQLNoPCL('update OPERATIONS set ROP_CODEIMPORT="" where ISNULL(ROP_CODEIMPORT,"")=""');
  ExecuteSQLNoPCL('update ERTINFOS001 set EER_FICHIERORIG="" where ISNULL(EER_FICHIERORIG,"")=""');
  ExecuteSQLNoPCL('update ERTINFOS002 set ER2_FICHIERORIG="" where ISNULL(ER2_FICHIERORIG,"")=""');
  ExecuteSQLNoPCL('update ERTINFOS00V set ERV_FICHIERORIG="" where ISNULL(ERV_FICHIERORIG,"")=""');


End;

Procedure MajVer933;
var db0 : string;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //T PETETIN 3751
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Qualifiant de mesure"' +
                                    ',DH_EXPLICATION="Qualifiant par défaut des unités de la fiche article" WHERE DH_PREFIXE="GA" AND DH_NOMCHAMP="GA_CODEFORME"');
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Coef. conv. qte conso en qte stock"' +
                                    ',DH_EXPLICATION="Coefficient de conversion entre l''unité de consommation et l''unité de quantité de stock" WHERE DH_PREFIXE="GA" AND DH_NOMCHAMP="GA_COEFCONVCONSO"');
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Coef. conv. qte achat en qte stock"' +
                                    ',DH_EXPLICATION="Coefficient de conversion entre l''unité de quantité d''achat et l''unité de quantité de stock" WHERE DH_PREFIXE="GA" AND DH_NOMCHAMP="GA_COEFCONVQTEACH"');
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Coef. conv. qte vente en qte stock"' +
                                    ',DH_EXPLICATION="Coefficient de conversion entre l''unité de quantité de vente et l''unité de quantité de stock" WHERE DH_PREFIXE="GA" AND DH_NOMCHAMP="GA_COEFCONVQTEVTE"');
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Coef. conv. qte prod. en qte stock"' +
                                    ',DH_EXPLICATION="Coefficient de conversion entre l''unité de production et l''unité de quantité de stock" WHERE DH_PREFIXE="GA" AND DH_NOMCHAMP="GA_COEFPROD"');
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Prix pour quantité (vente)"' +
                                    ',DH_EXPLICATION="Quantité pour laquelle les prix de vente sont exprimés" WHERE DH_PREFIXE="GA" AND DH_NOMCHAMP="GA_PRIXPOURQTE"');
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Prix pour quantité (achat)"' +
                                    ',DH_EXPLICATION="Quantité pour laquelle les prix d''achat sont exprimés" WHERE DH_PREFIXE="GA" AND DH_NOMCHAMP="GA_PRIXPOURQTEAC"');
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Unité de stock" WHERE DH_PREFIXE="GA" AND DH_NOMCHAMP="GA_QUALIFUNITESTO"');
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Unité de prix de vente" WHERE DH_PREFIXE="GA" AND DH_NOMCHAMP="GA_QUALIFUNITEVTE"');
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Unité de consommation" WHERE DH_PREFIXE="GA" AND DH_NOMCHAMP="GA_UNITECONSO"');
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Unité de prix d''achat" WHERE DH_PREFIXE="GA" AND DH_NOMCHAMP="GA_UNITEPRIXACH" ');
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Unité de production" WHERE DH_PREFIXE="GA" AND DH_NOMCHAMP="GA_UNITEPROD"');
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Unité de quantité d''achat" WHERE DH_PREFIXE="GA" AND DH_NOMCHAMP="GA_UNITEQTEACH" ');
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Unité de quantité de vente" WHERE DH_PREFIXE="GA" AND DH_NOMCHAMP="GA_UNITEQTEVTE" ');
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Coef. conv. qte achat en qte stock"' +
                                    ',DH_EXPLICATION="Coefficient de conversion entre l''unité de quantité d''achat de la fiche catalogu et l''unité de quantité de stock de la fiche article" WHERE DH_PREFIXE="GCA" AND DH_NOMCHAMP="GCA_COEFCONVQTEACH"');
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Prix pour quantité (achat)" WHERE DH_PREFIXE="GCA" AND DH_NOMCHAMP="GCA_PRIXPOURQTEAC"');
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Unité de prix d''achat" WHERE DH_PREFIXE="GCA" AND DH_NOMCHAMP="GCA_QUALIFUNITEACH"');
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Unité de quantité d''achat" WHERE DH_PREFIXE="GCA" AND DH_NOMCHAMP="GCA_UNITEQTEACH"');

  //JTR 3758
  { GC_20081126_JTR_010;15435 }
//  ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA = "-" WHERE SOC_NOM = "SO_OKFAMCPTAPIECE"');

  // M DESGOUTTE 3763
  // Purge des orphelins autour des groupes de travail
  if IsMonoOuCommune and (V_PGI.ModePCL='1') then
  begin
    ExecuteSQLContOnExcept('DELETE FROM LIENDOSGRP WHERE LDO_NOM="GROUPECONF" AND NOT EXISTS (select 1 from GRPDONNEES where ldo_grpid=grp_id and grp_nom="GROUPECONF")');
    ExecuteSQLContOnExcept('DELETE FROM LIENDONNEES WHERE LND_NOM="GROUPECONF" AND NOT EXISTS (select 1 from UTILISAT where lnd_userid=us_utilisateur)');
    ExecuteSQLContOnExcept('DELETE FROM LIENDOSGRP WHERE LDO_NOM="GROUPECONF" AND NOT EXISTS (select 1 from DOSSIER where ldo_nodossier=dos_nodossier)');
  end;
  // Vérification de la redirection de la table CODEPOST (même pour la DB0)
  if V_PGI.ModePCL='1' then // => ou sinon comment tester le mode de redirection de deshare ?
  begin
              db0 := v_pgi.DefaultSectionDBName;
              if db0 = '' then
                 db0 := V_PGI.DbName;
              if Not ExisteSQL ('SELECT 1 FROM DB' + V_PGI.NoDossier + '.dbo.DESHARE WHERE DS_NOMTABLE="CODEPOST" AND DS_TYPTABLE="TAB"') then
                   ExecuteSQLContOnExcept('INSERT INTO DB' + V_PGI.NoDossier + '.dbo.DESHARE (DS_NOMTABLE, DS_MODEFONC, DS_NOMBASE, DS_TYPTABLE) VALUES ("CODEPOST", "DB0","' + db0 + '", "TAB")');
  end;
  // D KOZA 3807
//  ExecuteSqlNoPCL('UPDATE WARTNAT SET WAN_CIRCUITSCM = "" WHERE WAN_CIRCUITSCM IS NULL');

  // M MORRETTON 3770
(*
  ExecuteSQLNoPCL('UPDATE YTARIFSPARAMETRES SET YFO_PROFILCOMPO="B01", YFO_COEFTYPEDONNEE=iif(YFO_FONCTIONNALITE="211","COE","POU")'+
    ', YFO_COEFSENSDONNEE=iif((YFO_FONCTIONNALITE IN ("501","601", "211")),"AUG","DIM"), YFO_COEFSENSCALCUL="INI", YFO_OKCOMMERCIAL="----",'+
    ' YFO_LIBELLEDONNEE="" WHERE YFO_COEFTYPEDONNEE IS NULL');

  ExecuteSQLNoPCL('UPDATE YTARIFSPECIAL SET YTP_QUELTIERS="CDE", YTP_PROFILTARIF="-" WHERE YTP_QUELTIERS IS NULL');

  ExecuteSQLNoPCL('UPDATE YTARIFS SET YTS_PROFILCOMPO="B01", YTS_COMMERCIAL="", YTS_TYPECOMM="", YTS_CASCCOMM="-", YTS_CASCTYPECOMM="-",'
  +' YTS_CASCTOUSCOMM="-", YTS_VALORISATIONDE=YTS_FOURCHETTEDE WHERE YTS_PROFILCOMPO IS NULL');
*)
(*  ExecuteSQLNoPCL('UPDATE LIGNEFRAIS SET LF_PROFILCOMPO="B01", LF_LOWLEVELCODE=0, LF_RANGCOMM="", LF_BLOQUETARIF="-", LF_VALORISATIONDE="Q",'
  +' LF_COEFTYPEDONNEE=iif(LF_FONCTIONNALITE="211","COE","POU"), LF_COEFSENSDONNEE=iif((LF_FONCTIONNALITE IN ("211","501","601")),"AUG","DIM"), LF_COEFSENSCALCUL="INI",'
  +' LF_NOMCHAMPBASE="", LF_NOMCHAMPREFE="", LF_NOMCHAMPEDIT="", LF_NOMCHAMPCOMM="", LF_ARRONDIMT="-", LF_TARIFSPECIAL="", LF_VALEUREDIT=0'
  +' WHERE LF_PROFILCOMPO IS NULL');
*)
//  ExecuteSQLNoPCL('UPDATE WPARAM SET WPA_VARCHAR17="", WPA_VARCHAR18="", WPA_VARCHAR19="" WHERE WPA_VARCHAR17 IS NULL');
End;

Procedure MajVer934;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  // C PARWEZ 3811
(*
  ExecuteSQLNoPCL('UPDATE WPARAMFONCTION SET WPF_CODEFONCTION = "CIRCDEP", WPF_VARCHAR01    = WPF_CIRCUIT || ";",'
  +' WPF_CIRCUIT = "" WHERE WPF_CODEFONCTION = "CBNCIRCDEP"');
  // M GUERIN 3817
*)
  {Modif faute d orthographe dans un champ obligatoire pour Business Side}
 // ExecuteSQLContOnExcept('update commun set CO_ABREGE = "EEP_REPRESENTANT" where co_type = "EC3" and co_code = "EP4"'
 // +' and CO_ABREGE = "EEP_REPRESANTANT"');
  {Initialisation des champs à Null}
//  ExecuteSQLNoPCL('update EPERSPECTIVES set EEP_CODEIMPORT="" where ISNULL(EEP_CODEIMPORT,"")=""');
//  ExecuteSQLNoPCL('update PERSPECTIVES set RPE_CODEIMPORT="" where ISNULL(RPE_CODEIMPORT,"")=""');
  //DBR 3823
//  ExecuteSQLContOnExcept ('delete from paramsoc where soc_nom="SCO_GRPECART"');

  //MCD 3848
  // pour initialisation des champs
(*
  ExecuteSQLNoPCL('UPDATE ACTIVITE SET ACT_PHASEAFF1="",ACT_PHASEAFF2="",ACT_PHASEAFF3="",ACT_PHASEAFF4="",'
  +'ACT_PHASEAFF5="",ACT_CODEPHASEAFF1="",ACT_CODEPHASEAFF2="",ACT_CODEPHASEAFF3="",ACT_CODEPHASEAFF4="",'
  +'ACT_CODEPHASEAFF5="" WHERE ACT_PHASEAFF1 IS NULL ');
*)
(*
  ExecuteSQLNoPCL('UPDATE EACTIVITE SET EAC_PHASEAFF1="",EAC_PHASEAFF2="",EAC_PHASEAFF3="",EAC_PHASEAFF4="",'
  +'EAC_PHASEAFF5="",EAC_CODEPHASEAFF1="",EAC_CODEPHASEAFF2="",EAC_CODEPHASEAFF3="",EAC_CODEPHASEAFF4="",'
  +'EAC_CODEPHASEAFF5="" WHERE EAC_PHASEAFF1 IS NULL ');
*)
//  ExecuteSQLNoPCL(' UPDATE FACTAFF SET AFA_NUMORDRE=0,AFA_PHASEAFF1="",AFA_PHASEAFF2="",AFA_PHASEAFF3="",'   // GM regroupé  948
//  +'AFA_PHASEAFF4="",AFA_PHASEAFF5="" WHERE AFA_NUMORDRE IS NULL');

//  ExecuteSQLNoPCL(' UPDATE AFFAIRE SET AFF_MODEFACTPHASE="" WHERE AFF_MODEFACTPHASE IS NULL');    GM regroupé en 949
(*
  if not IsDossierPCL then
    UpDateDecoupeLigne ('GLC_PHASEAFF1="", GLC_PHASEAFF2="", GLC_PHASEAFF3="", GLC_PHASEAFF4="",' +
                        'GLC_PHASEAFF5="", GLC_CODEPHASEAFF1="", GLC_CODEPHASEAFF2="", GLC_CODEPHASEAFF3="", GLC_CODEPHASEAFF4="",' +
                        'GLC_CODEPHASEAFF5=""',
                      'AND GLC_PHASEAFF1 IS NULL', 'GLC');
*)
End;

Procedure MajVer935;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //TS GP_20081216_TS_GP15834
//  if not IsDossierPCL() then
//    EdiMajVer(935).MajApres();

{DBR 4054 on supprime cette demande
  //DBR 3864
  ExecuteSQLNoPCL ('UPDATE LIGNEFRAIS SET LF_CATEGORIETAXE="" WHERE LF_CATEGORIETAXE IS NULL');}

  //MD3 3877
  // Uniquement sur base commune
  If IsMonoOuCommune then
    ExecuteSQLContOnExcept('UPDATE DPOBLIGATIONREALISE SET DPO_DATESYS=DPO_DATEOBLIGATION');

  //F BERGER 3878
   // Je coche "Gestion cout standard par liste" si il existe déjà des listes COUTSTDLIG dans la base (Cas ERA)
  if not IsDossierPCL then
  begin
    if ExisteSQL('SELECT ##TOP 1## CSL_LISTESTD FROM COUTSTDLIG') then
      ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA = "X" WHERE SOC_NOM = "SO_PSTDLISTE"')
    else
      ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA = "-" WHERE SOC_NOM = "SO_PSTDLISTE"');
  end;

  //M MORRETTON 3886
  if (V_PGI.Driver <> dbORACLE7 ) and not IsDossierPCL then
//    MajPoidsTarifs;
(*
  ExecuteSQLNoPCL('UPDATE LIGNEFRAIS SET LF_ORIFORF'    +'=SUBSTRING(LF_ORIFORF'    +',1,31)'
  +'||SUBSTRING(LF_ORIFORF'    +',37,1)||SUBSTRING(LF_ORIFORF'    +',33,47)||"....-|"||'
  +'SUBSTRING(LF_ORIFORF'    +',81,31)||SUBSTRING(LF_ORIFORF'    +',117,1)||'
  +'SUBSTRING(LF_ORIFORF'    +',113,47)||"....-|"||SUBSTRING(LF_ORIFORF'    +',161,40)'
  +' WHERE LF_ORIFORF'    +'<>"m" AND LF_ORIFORF'    +'<>"" AND SUBSTRING(LF_ORIFORF'    +',80,1)="|"');

  ExecuteSQLNoPCL('UPDATE LIGNEFRAIS SET LF_ORIFIXE'    +'=SUBSTRING(LF_ORIFIXE'    +',1,31)'
  +'||SUBSTRING(LF_ORIFIXE'    +',37,1)||SUBSTRING(LF_ORIFIXE'    +',33,47)||"....-|"||'
  +'SUBSTRING(LF_ORIFIXE'    +',81,31)||SUBSTRING(LF_ORIFIXE'    +',117,1)||'
  +'SUBSTRING(LF_ORIFIXE'    +',113,47)||"....-|"||SUBSTRING(LF_ORIFIXE'    +',161,40)'
  +' WHERE LF_ORIFIXE'    +'<>"m" AND LF_ORIFIXE'    +'<>"" AND SUBSTRING(LF_ORIFIXE'    +',80,1)="|"');

  ExecuteSQLNoPCL('UPDATE LIGNEFRAIS SET LF_ORIREMMT'   +'=SUBSTRING(LF_ORIREMMT'   +',1,31)'
  +'||SUBSTRING(LF_ORIREMMT'   +',37,1)||SUBSTRING(LF_ORIREMMT'   +',33,47)||"....-|"||'
  +'SUBSTRING(LF_ORIREMMT'   +',81,31)||SUBSTRING(LF_ORIREMMT'   +',117,1)||'
  +'SUBSTRING(LF_ORIREMMT'   +',113,47)||"....-|"||SUBSTRING(LF_ORIREMMT'   +',161,40)'
  +' WHERE LF_ORIREMMT'   +'<>"m" AND LF_ORIREMMT'   +'<>"" AND SUBSTRING(LF_ORIREMMT'   +',80,1)="|"');

  ExecuteSQLNoPCL('UPDATE LIGNEFRAIS SET LF_ORICOUTBRUT'+'=SUBSTRING(LF_ORICOUTBRUT'+',1,31)'
  +'||SUBSTRING(LF_ORICOUTBRUT'+',37,1)||SUBSTRING(LF_ORICOUTBRUT'+',33,47)||"....-|"||'
  +'SUBSTRING(LF_ORICOUTBRUT'+',81,31)||SUBSTRING(LF_ORICOUTBRUT'+',117,1)||'
  +'SUBSTRING(LF_ORICOUTBRUT'+',113,47)||"....-|"||SUBSTRING(LF_ORICOUTBRUT'+',161,40)'
  +' WHERE LF_ORICOUTBRUT'+'<>"m" AND LF_ORICOUTBRUT'+'<>"" AND SUBSTRING(LF_ORICOUTBRUT'+',80,1)="|"');

  ExecuteSQLNoPCL('UPDATE LIGNEFRAIS SET LF_ORICOUTNET' +'=SUBSTRING(LF_ORICOUTNET' +',1,31)'
  +'||SUBSTRING(LF_ORICOUTNET' +',37,1)||SUBSTRING(LF_ORICOUTNET' +',33,47)||"....-|"||'
  +'SUBSTRING(LF_ORICOUTNET' +',81,31)||SUBSTRING(LF_ORICOUTNET' +',117,1)||'
  +'SUBSTRING(LF_ORICOUTNET' +',113,47)||"....-|"||SUBSTRING(LF_ORICOUTNET' +',161,40)'
  +' WHERE LF_ORICOUTNET' +'<>"m" AND LF_ORICOUTNET' +'<>"" AND SUBSTRING(LF_ORICOUTNET' +',80,1)="|"');

  ExecuteSQLNoPCL('UPDATE LIGNEFRAIS SET LF_ORIPOURCENT'+'=SUBSTRING(LF_ORIPOURCENT'+',1,31)'
  +'||SUBSTRING(LF_ORIPOURCENT'+',37,1)||SUBSTRING(LF_ORIPOURCENT'+',33,47)||"....-|"||'
  +'SUBSTRING(LF_ORIPOURCENT'+',81,31)||SUBSTRING(LF_ORIPOURCENT'+',117,1)||'
  +'SUBSTRING(LF_ORIPOURCENT'+',113,47)||"....-|"||SUBSTRING(LF_ORIPOURCENT'+',161,40)'
  +' WHERE LF_ORIPOURCENT'+'<>"m" AND LF_ORIPOURCENT'+'<>"" AND SUBSTRING(LF_ORIPOURCENT'+',80,1)="|"');
*)
  // D SCLAVOPOULOS 3901
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Budget modifiable" WHERE DH_NOMCHAMP = "AFF_OKMAJBUDGET"');

  //MCD 3902
// ExecuteSQLNoPCL ('UPDATE FACTAFF SET AFA_NUMORDRE=0,AFA_PHASEAFF1="",AFA_PHASEAFF2="",AFA_PHASEAFF3="",'
//  +'AFA_PHASEAFF4="",AFA_PHASEAFF5="" WHERE AFA_PHASEAFF1 IS NULL ');    // GM regroupé 948


//  ExecuteSQLNoPCL (' UPDATE AFFAIRE SET AFF_MODEFACTPHASE="-" WHERE AFF_MODEFACTPHASE IS NULL OR AFF_MODEFACTPHASE =""');   Regroupé en 949

end;

Procedure MajVer936;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  // JP LAURENT 3908
(*
  ExecuteSQLNoPCL('update QBPARBRE SET QBR_NIVALAFF = 1 WHERE QBR_NIVALAFF is null');
  ExecuteSQLNoPCL('update QBPARBREDETAIL SET QBH_NIVALAFF = 1 WHERE QBH_NIVALAFF is null');
  ExecuteSQLNoPCL('update QBPCUBETMP SET QBQ_NIVALAFF = 1,QBQ_COMMENTAIREBP=""  WHERE QBQ_NIVALAFF is null');
*)
  if not isDossierPCL then
  begin
    //N FOURNEL 3920
    (*
    if not ExisteSQL ('SELECT SH_SOUCHE FROM SOUCHE WHERE SH_SOUCHE = "GAA"') then
    begin
      ExecuteSQLNoPCL('INSERT INTO SOUCHE (SH_ENTITY,SH_TYPE, SH_SOUCHE) VALUES (0, "GES", "GAA")');
      ExecuteSQLNoPCL('UPDATE SOUCHE SET SH_ABREGE="Facture d''acompte",'
                                + ' SH_ANALYTIQUE="-",'
                                + ' SH_DATEDEBUT="'+UsDateTime_(iDate1900)+'",'
                                + ' SH_DATEFIN="'+UsDateTime_(iDate1900)+'",'
                                + ' SH_ENTITY=0,'
                                + ' SH_FERME="-",'
                                + ' SH_JOURNAL="",'
                                + ' SH_LIBELLE="Facture d''acompte",'
                                + ' SH_MASQUENUM="",'
                                + ' SH_NATUREPIECE="",'
                                + ' SH_NATUREPIECEG="",'
                                + ' SH_NUMDEPART=1,'
                                + ' SH_NUMDEPARTP=1,'
                                + ' SH_NUMDEPARTS=1,'
                                + ' SH_RESERVEWEB="-",'
                                + ' SH_SIMULATION="-",'
                                + ' SH_SOCIETE="001",'
                                + ' SH_SOUCHE="GAA",'
                                + ' SH_SOUCHEEXO="-",'
                                + ' SH_TYPE="GES"'
                                + ' WHERE SH_SOUCHE = "GAA"');
    end;
    if not ExisteSQL ('SELECT GPP_LIBELLE FROM PARPIECE WHERE GPP_NATUREPIECEG="FAA"') then
    begin
      ExecuteSqlNoPCL ('INSERT INTO PARPIECE (GPP_NATUREPIECEG) VALUES("FAA")');
      ExecuteSqlNoPCL ('UPDATE PARPIECE SET GPP_ACHATACTIVITE="-",'
                                   + ' GPP_ACOMPTE="-",'
                                   + ' GPP_ACTIONFINI="ENR",'
                                   + ' GPP_ACTIVITEPUPR="",'
                                   + ' GPP_AFAFFECTTB="",'
                                   + ' GPP_AFFPIECETABLE="-",'
                                   + ' GPP_APERCUAVETIQ="-",'
                                   + ' GPP_APERCUAVIMP="-",'
                                   + ' GPP_APPELPRIX="",'
                                   + ' GPP_APPLICRG="-",'
                                   + ' GPP_ARTFOURPRIN="-",'
                                   + ' GPP_ARTSTOCK="-",'
                                   + ' GPP_BLOBART="X",'
                                   + ' GPP_BLOBLIENART="",'
                                   + ' GPP_BLOBLIENTIERS="",'
                                   + ' GPP_BLOBTIERS="X",'
                                   + ' GPP_CALCRUPTURE="AUC",'
                                   + ' GPP_CFGART="-",'
                                   + ' GPP_CFGARTASSIST="",'
                                   + ' GPP_CHAINAGE="",'
                                   + ' GPP_CODEPIECEOBL1="-",'
                                   + ' GPP_CODEPIECEOBL2="-",'
                                   + ' GPP_CODEPIECEOBL3="-",'
                                   + ' GPP_CODPIECEDEF1="",'
                                   + ' GPP_CODPIECEDEF2="",'
                                   + ' GPP_CODPIECEDEF3="",'
                                   + ' GPP_COMMENTENT="",'
                                   + ' GPP_COMMENTPIED="",'
                                   + ' GPP_COMPANALLIGNE="",'
                                   + ' GPP_COMPANALPIED="",'
                                   + ' GPP_COMPSTOCKLIGNE="",'
                                   + ' GPP_COMPSTOCKPIED="",'
                                   + ' GPP_CONDITIONTARIF="-",'
                                   + ' GPP_CONTEXTES="GC;AFF;",'
                                   + ' GPP_CONTRECHART1="-",'
                                   + ' GPP_CONTRECHART2="-",'
                                   + ' GPP_CONTRECHART3="-",'
                                   + ' GPP_CONTREMARQUE="-",'
                                   + ' GPP_CONTREMREF="-",'
                                   + ' GPP_CONTROLEMARGE="AUC",'
                                   + ' GPP_CPTCENTRAL="-",'
                                   + ' GPP_CRC=-1269890210,'
                                   + ' GPP_CTRLENCOURS="-",'
                                   + ' GPP_CUMULART1="-",'
                                   + ' GPP_CUMULART2="-",'
                                   + ' GPP_CUMULART3="-",'
                                   + ' GPP_CUMULCOM1="-",'
                                   + ' GPP_CUMULCOM2="-",'
                                   + ' GPP_CUMULCOM3="-",'
                                   + ' GPP_CUMULTIERS1="-",'
                                   + ' GPP_CUMULTIERS2="-",'
                                   + ' GPP_CUMULTIERS3="-",'
                                   + ' GPP_DATELIBART1="AUC",'
                                   + ' GPP_DATELIBART2="AUC",'
                                   + ' GPP_DATELIBART3="AUC",'
                                   + ' GPP_DATELIBCOM1="AUC",'
                                   + ' GPP_DATELIBCOM2="AUC",'
                                   + ' GPP_DATELIBCOM3="AUC",'
                                   + ' GPP_DATELIBTIERS1="AUC",'
                                   + ' GPP_DATELIBTIERS2="AUC",'
                                   + ' GPP_DATELIBTIERS3="AUC",'
                                   + ' GPP_DIMSAISIE="",'
                                   + ' GPP_DUPLICPIECE="FAA;",'
                                   + ' GPP_ECLATEAFFAIRE="-",'
                                   + ' GPP_ECLATEDOMAINE="-",'
                                   + ' GPP_EDITIONNOMEN="",'
                                   + ' GPP_ENCOURS="-",'
                                   + ' GPP_EQUIPIECE="",'
                                   + ' GPP_ESTAVOIR="-",'
                                   + ' GPP_ETATETIQ="",'
                                   + ' GPP_FAR_FAE="",'
                                   + ' GPP_FILTREARTCH="",'
                                   + ' GPP_FILTREARTVAL="",'
                                   + ' GPP_FILTRECOMM="",'
                                   + ' GPP_FORCERUPTURE="-",'
                                   + ' GPP_GEREARTICLELIE="",'
                                   + ' GPP_GEREECHEANCE="",'
                                   + ' GPP_GESTIONGRATUIT="-",'
                                   + ' GPP_HISTORIQUE="X",'
                                   + ' GPP_IFL1="",'
                                   + ' GPP_IFL2="",'
                                   + ' GPP_IFL3="",'
                                   + ' GPP_IFL4="",'
                                   + ' GPP_IFL5="",'
                                   + ' GPP_IFL6="",'
                                   + ' GPP_IFL7="",'
                                   + ' GPP_IFL8="",'
                                   + ' GPP_IMPAUTOBESCBN="-",'
                                   + ' GPP_IMPAUTOETATCBN="-",'
                                   + ' GPP_IMPBESOIN="-",'
                                   + ' GPP_IMPETAT="",'
                                   + ' GPP_IMPETIQ="-",'
                                   + ' GPP_IMPIMMEDIATE="X",'
                                   + ' GPP_IMPMODELE="GFC",' //Remplacer par le nouvel état FAA
                                   + ' GPP_INFOPCEPRECCHX="",'
                                   + ' GPP_INFOPIECEPREC="",'
                                   + ' GPP_INFOSCOMPL="-",'
                                   + ' GPP_INFOSCPLPIECE="-",'
                                   + ' GPP_INITQTE="",'
                                   + ' GPP_INITQTECRE=0,'
                                   + ' GPP_INSERTLIG="X",'
                                   + ' GPP_JOURNALCPTA="",'
                                   + ' GPP_LIBELLE="Facture d''acompte",'
                                   + ' GPP_LIENAFFAIRE="-",'
                                   + ' GPP_LIENTACHE="-",'
                                   + ' GPP_LISTEAFFAIRE="",'
                                   + ' GPP_LISTESAISIE="GCSAISIEFAC",'
                                   + ' GPP_LOT="-",'
                                   + ' GPP_MAJINFOTIERS="X",'
                                   + ' GPP_MAJPRIXVALO="",'
                                   + ' GPP_MASQUERNATURE="-",'
                                   + ' GPP_MENU="",'
                                   + ' GPP_MESSAGEEDIIN="",'
                                   + ' GPP_MESSAGEEDIOUT="",'
                                   + ' GPP_MODBESOIN="",'
                                   + ' GPP_MODEECHEANCES="",'
                                   + ' GPP_MODEGROUPEPORT="",'
                                   + ' GPP_MODELEWORD="",'
                                   + ' GPP_MODIFCOUT="-",'
                                   + ' GPP_MODPLANIFIABLE="",'
                                   + ' GPP_MONTANTMINI=0,'
                                   + ' GPP_MONTANTVISA=0,'
                                   + ' GPP_MULTIGRILLE="-",'
                                   + ' GPP_NATPIECEANNUL="",'
                                   + ' GPP_NATURECPTA="FC",'
                                   + ' GPP_NATUREORIGINE="FAC",'
                                   + ' GPP_NATUREPIECEG="FAA",'
                                   + ' GPP_NATUREREPRISE="",'
                                   + ' GPP_NATURESUIVANTE="",'
                                   + ' GPP_NATURETIERS="CLI;",'
                                   + ' GPP_NBEXEMPLAIRE=0,'
                                   + ' GPP_NIVEAUPARAM="EDI",'
                                   + ' GPP_NUMEROSERIE="-",'
                                   + ' GPP_OBJETDIM="-",'
                                   + ' GPP_OBLIGEREGLE="-",'
                                   + ' GPP_OUVREAUTOPORT="-",'
                                   + ' GPP_PARAMDIM="-",'
                                   + ' GPP_PARAMGRILLEDIM="-",'
                                   + ' GPP_PIECEEDI="-",'
                                   + ' GPP_PIECEPILOTE="-",'
                                   + ' GPP_PIECESAV="-",'
                                   + ' GPP_PIECETABLE1="",'
                                   + ' GPP_PIECETABLE2="",'
                                   + ' GPP_PIECETABLE3="",'
                                   + ' GPP_PILOTEORDRE="-",'
                                   + ' GPP_PREVUAFFAIRE="-",'
                                   + ' GPP_PRIORECHART1="ART",'
                                   + ' GPP_PRIORECHART2="AUC",'
                                   + ' GPP_PRIORECHART3="AUC",'
                                   + ' GPP_PRIXNULOK="X",'
                                   + ' GPP_PROCLI="-",'
                                   + ' GPP_QTEMOINS="",'
                                   + ' GPP_QTEPLUS="",'
                                   + ' GPP_QUALIFMVT="",'
                                   + ' GPP_RACINELIBECR1="",'
                                   + ' GPP_RACINELIBECR2="",'
                                   + ' GPP_RACINEREFINT1="",'
                                   + ' GPP_RACINEREFINT2="",'
                                   + ' GPP_RECALCULPRIX="-",'
                                   + ' GPP_RECHTARIF501="-",'
                                   + ' GPP_RECUPPRE="",'
                                   + ' GPP_REFEXTCTRL="",'
                                   + ' GPP_REFINTCTRL="",'
                                   + ' GPP_REFINTEXT="INT",'
                                   + ' GPP_REGROUPCPTA="",'
                                   + ' GPP_REGROUPE="-",'
                                   + ' GPP_RELIQUAT="-",'
                                   + ' GPP_REPRISEENTAFF="",'
                                   + ' GPP_REPRISELIGAFF="",'
                                   + ' GPP_SENSPIECE="MIX",'
                                   + ' GPP_SOLDETRANSFO="-",'
                                   + ' GPP_SOUCHE="GAA",'
                                   + ' GPP_STKQUALIFMVT="",'
                                   + ' GPP_TARIFGENDATE="",'
                                   + ' GPP_TARIFGENDEPOT="-",'
                                   + ' GPP_TARIFGENPNPB="-",'
                                   + ' GPP_TARIFGENSAISIE="",'
                                   + ' GPP_TARIFGENSPECIA="",'
                                   + ' GPP_TARIFGENTRANSF="",'
                                   + ' GPP_TARIFMODULE="",'
                                   + ' GPP_TAXE="X",'
                                   + ' GPP_TIERS="",'
                                   + ' GPP_TRSFACHAT="-",'
                                   + ' GPP_TRSFVENTE="-",'
                                   + ' GPP_TYPEACTIVITE="",'
                                   + ' GPP_TYPEARTICLE="CNS;FI;FRA;MAR;NOM;PRE;",'
                                   + ' GPP_TYPECOMMERCIAL="",'
                                   + ' GPP_TYPEDIMOBLI1="-",'
                                   + ' GPP_TYPEDIMOBLI2="-",'
                                   + ' GPP_TYPEDIMOBLI3="-",'
                                   + ' GPP_TYPEDIMOBLI4="-",'
                                   + ' GPP_TYPEDIMOBLI5="-",'
                                   + ' GPP_TYPEECRALIM="",'
                                   + ' GPP_TYPEECRCPTA="NOR",'
                                   + ' GPP_TYPEECRSTOCK="RIE",'
                                   + ' GPP_TYPEFACT="",'
                                   + ' GPP_TYPEPASSACC="REE",'
                                   + ' GPP_TYPEPASSACCR="REE",'
                                   + ' GPP_TYPEPASSCPTA="REE",'
                                   + ' GPP_TYPEPASSCPTAR="REE",'
                                   + ' GPP_TYPEPRESDOC="AUC",'
                                   + ' GPP_TYPEPRESENT=0,'
                                   + ' GPP_VALEURLIBECR1="",'
                                   + ' GPP_VALEURLIBECR2="",'
                                   + ' GPP_VALEURREFINT1="",'
                                   + ' GPP_VALEURREFINT2="",'
                                   + ' GPP_VALLIBART1="AUC",'
                                   + ' GPP_VALLIBART2="AUC",'
                                   + ' GPP_VALLIBART3="AUC",'
                                   + ' GPP_VALLIBCOM1="AUC",'
                                   + ' GPP_VALLIBCOM2="AUC",'
                                   + ' GPP_VALLIBCOM3="AUC",'
                                   + ' GPP_VALLIBTIERS1="AUC",'
                                   + ' GPP_VALLIBTIERS2="AUC",'
                                   + ' GPP_VALLIBTIERS3="AUC",'
                                   + ' GPP_VALMODELE="-",'
                                   + ' GPP_VENTEACHAT="VEN",'
                                   + ' GPP_VISA="-"'
                                   + ' WHERE GPP_NATUREPIECEG = "FAA"');
    end;
    *)
  end;
//  ExecuteSqlNoPCL ('UPDATE ACOMPTES SET GAC_TAX = "", GAC_PIECEFAA = "", GAC_DATEACOMPTE = "'+ UsDateTime_(iDate1900) +'", ' +
//                     'GAC_MNTTAXDEV = 0, GAC_MNTTAX = 0, GAC_NUMLIGNEECR = 0 WHERE GAC_TAX IS NULL');
//  ExecuteSqlNoPCL ('UPDATE SOUCHE SET SH_NUMDEPART = 1 WHERE SH_SOUCHE="GAA"');

  //D SCLAVOPOULOS 3921
//  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = "LDC" WHERE DH_NOMCHAMP="GLC_QACCSAIS" OR  DH_NOMCHAMP="GLC_DATEACC" OR DH_NOMCHAMP="GLC_UNITEACC"');

  //D KOZA 3941
  //D KOZA le 01.09.2009 Modification de l'initialisation des champs _DIVPROD de "" => "."
  (*
  ExecuteSqlNoPCL('UPDATE QAPPROS SET QA_DIVPROD=".", QA_GESTIONNAIRE="" WHERE QA_DIVPROD IS NULL');
  ExecuteSqlNoPCL('UPDATE QCAMPAREG SET QCR_DIVPROD="." WHERE QCR_DIVPROD IS NULL');
  ExecuteSqlNoPCL('UPDATE QCHGTGAM SET QCG_DIVPROD="." WHERE QCG_DIVPROD IS NULL');
  ExecuteSqlNoPCL('UPDATE QCIRCUIT SET QCI_NATURECIRC="", QCI_NATURETRAVAIL="", QCI_PAYS="", QCI_POLE="", QCI_TYPECIRC="" WHERE QCI_NATURECIRC IS NULL');
  ExecuteSqlNoPCL('UPDATE QCODELIB SET QCB_CLLIBELLELONG="", QCB_CLLIBRE="" WHERE QCB_CLLIBELLELONG IS NULL');
  ExecuteSqlNoPCL('UPDATE QCOMMENT SET QCO_DIVPROD="." WHERE QCO_DIVPROD IS NULL');
  ExecuteSqlNoPCL('UPDATE QDATEFAB SET QDF_DIVPROD="." WHERE QDF_DIVPROD IS NULL');
  ExecuteSqlNoPCL('UPDATE QDETCIRC SET QDE_PHASE="" WHERE QDE_PHASE IS NULL');
  ExecuteSqlNoPCL('UPDATE QDETPRI SET QDT_CODESOC="", QDT_NUMDETPRI="" WHERE QDT_CODESOC IS NULL');
  ExecuteSqlNoPCL('UPDATE QENVIRONNEMENT SET QEV_DIVPROD="." WHERE QEV_DIVPROD IS NULL');
  ExecuteSqlNoPCL('UPDATE QGROUPE SET QGR_CTXORDO="", QGR_DIVPROD=".", QGR_GRPETALON="", QGR_SITEETALON="" WHERE QGR_CTXORDO IS NULL');
  ExecuteSqlNoPCL('UPDATE QHISTOAFF SET QHA_DIVPROD=".", QHA_LIENPERE="", QHA_NUMDETTR="" WHERE QHA_DIVPROD IS NULL');
  ExecuteSqlNoPCL('UPDATE QLANCEMEN SET QLA_AFFAIRE="", QLA_DATECREATION="'+UsDateTime_(iDate1900)+'", QLA_DIVPROD=".", QLA_GESTIONNAIRE="", QLA_LIENPERE="", QLA_NOUVEAU="-", QLA_REGROUPEMENT="" WHERE QLA_AFFAIRE IS NULL');
  ExecuteSqlNoPCL('UPDATE QLCTQTMIN SET QLC_DIVPROD=".", QLC_LCTTYPEMAX="-", QLC_POIDSMAX=0, QLC_QTMINTAILLE=0 WHERE QLC_DIVPROD IS NULL');
  ExecuteSqlNoPCL('UPDATE QLIGSIMUL SET QLG_DIVPROD="." WHERE QLG_DIVPROD IS NULL');
  ExecuteSqlNoPCL('UPDATE QMACROGAM SET QMG_PHASE="", QMG_PROINDUS="" WHERE QMG_PHASE IS NULL');
  ExecuteSqlNoPCL('UPDATE QMAGASIN SET QMA_CODESOC="" WHERE QMA_CODESOC IS NULL');
  ExecuteSqlNoPCL('UPDATE QMAGENC SET QMN_GENRE="", QMN_MAGASINSPE="", QMN_MATIERE="", QMN_NUMREGLE="", QMN_VARIETE="" WHERE QMN_GENRE IS NULL');
  ExecuteSqlNoPCL('UPDATE QMATIERE SET QMT_GENRE="", QMT_VARIETE="" WHERE QMT_GENRE IS NULL');
  ExecuteSqlNoPCL('UPDATE QNOMENCLA SET QNO_CODELIAISON="", QNO_PHABESOIN="", QNO_PHACONSO="", QNO_PHASUIVICONSO="" WHERE QNO_CODELIAISON IS NULL');
  ExecuteSqlNoPCL('UPDATE QPBASEART SET QPB_DIVPROD=".", QPB_TYPEPBASEART="" WHERE QPB_DIVPROD IS NULL');
  ExecuteSqlNoPCL('UPDATE QPERSONNE SET QPE_DIVPROD="." WHERE QPE_DIVPROD IS NULL');
  ExecuteSqlNoPCL('UPDATE QPHASECAR SET QPX_DIVPROD="." WHERE QPX_DIVPROD IS NULL');
  ExecuteSqlNoPCL('UPDATE QPROFILCA SET QPF_CHAINEORDO="", QPF_COMPQTMINI="-", QPF_COMPTAILLELOT="", QPF_DATECHGCIRC="'+UsDateTime_(iDate1900)+'", QPF_DIVPROD="."'
  + ', QPF_NBDECIMAL=9, QPF_OKCHGCIRCOF="-", QPF_OKREGROUPEMENT="-", QPF_PARAMACTIF="-", QPF_PRIOTYPEDEC="-", QPF_REAPPAPPRO="-", QPF_REDECOUPAGE="-"'
  + ', QPF_REGLEQTMAX="", QPF_RESPECTAFFECT="", QPF_SELARTCHGCIRC="", QPF_SELOFCHGCIRC="", QPF_SELCIRCCHGCIRC="", QPF_TOLCTRTMAT=0 WHERE QPF_CHAINEORDO IS NULL');
  ExecuteSqlNoPCL('UPDATE QPROINDUS SET QPR_CODITI="", QPR_NATURETRAVAIL="", QPR_TYPENOMEN="" WHERE QPR_CODITI IS NULL');
  ExecuteSqlNoPCL('UPDATE QQTMINCAR SET QQM_DIVPROD="." WHERE QQM_DIVPROD IS NULL');
  ExecuteSqlNoPCL('UPDATE QREPARTFAB SET QRF_COLORIS="", QRF_DATEMODIF="'+UsDateTime_(iDate1900)+'", QRF_DIVPROD=".", QRF_GESTIONNAIRE="", QRF_MAGDEPART="", QRF_NATURETRAVAIL="", QRF_NUMPERE="", QRF_RANG=0 WHERE QRF_COLORIS IS NULL');
  ExecuteSqlNoPCL('UPDATE QSATTRIB SET QSB_DIVPROD=".", QSB_LIENPERE="", QSB_NUMDETTR="" WHERE QSB_DIVPROD IS NULL');
  ExecuteSqlNoPCL('UPDATE QSCENARCA SET QSC_DIVPROD=".", QSC_PURGESIMU="-" WHERE QSC_DIVPROD IS NULL');
  ExecuteSqlNoPCL('UPDATE QSIMULATION SET QSM_CTXORDO="", QSM_DESTINATION="", QSM_USERLOCK="" WHERE QSM_CTXORDO IS NULL');
  ExecuteSqlNoPCL('UPDATE QSTATREG SET QSR_CIRCUIT="", QSR_QUANTITE=0, QSR_TRANCHE="" WHERE QSR_CIRCUIT IS NULL');
  ExecuteSqlNoPCL('UPDATE QSTKCOLT SET QSK_PREMIERBESOIN="-" WHERE QSK_PREMIERBESOIN IS NULL');
  ExecuteSqlNoPCL('UPDATE QSTKDISPO SET QSD_DIVPROD=".", QSD_LIENPERE="" WHERE QSD_DIVPROD IS NULL');
  ExecuteSqlNoPCL('UPDATE QTEMPECH SET QTE_DIVPROD="." WHERE QTE_DIVPROD IS NULL');
  ExecuteSqlNoPCL('UPDATE QTEMPF SET QTF_CODESOC="", QTF_DIVPROD="." WHERE QTF_CODESOC IS NULL');
  ExecuteSqlNoPCL('UPDATE QTEMPP SET QTP_DIVPROD=".", QTP_PROINDUS="" WHERE QTP_DIVPROD IS NULL');
  ExecuteSqlNoPCL('UPDATE QTRANDETD SET QTD_CIRCUIT="", QTD_DATEDEB="'+UsDateTime_(iDate1900)+'", QTD_DATEFIN="'+UsDateTime_(iDate1900)+'", QTD_SAICMD="", QTD_TYPEBESOIN="" WHERE QTD_CIRCUIT IS NULL');
  ExecuteSqlNoPCL('UPDATE QVALCAR SET QVA_NUMPRIOCARACT=0 WHERE QVA_NUMPRIOCARACT IS NULL');
  *)
  if not isDossierPCL then
  begin
    //D BROSSET 3953
(*
     if not ExisteSQL ('SELECT GPP_LIBELLE FROM PARPIECE WHERE GPP_NATUREPIECEG="FFF"') then
    begin
      ExecuteSqlNoPCL ('INSERT INTO PARPIECE (GPP_NATUREPIECEG) VALUES("FFF")');
      ExecuteSqlNoPCL ('UPDATE PARPIECE SET GPP_ACHATACTIVITE="-",'
                  + ' GPP_ACOMPTE="-",'
                  + ' GPP_ACTIONFINI="COM",'
                  + ' GPP_ACTIVITEPUPR="",'
                  + ' GPP_AFAFFECTTB="ACH",'
                  + ' GPP_AFFPIECETABLE="-",'
                  + ' GPP_APERCUAVETIQ="-",'
                  + ' GPP_APERCUAVIMP="-",'
                  + ' GPP_APPELPRIX="DPA",'
                  + ' GPP_APPLICRG="-",'
                  + ' GPP_ARTFOURPRIN="-",'
                  + ' GPP_ARTSTOCK="-",'
                  + ' GPP_BLOBART="-",'
                  + ' GPP_BLOBLIENART="",'
                  + ' GPP_BLOBLIENTIERS="",'
                  + ' GPP_BLOBTIERS="-",'
                  + ' GPP_CALCRUPTURE="AUC",'
                  + ' GPP_CFGART="-",'
                  + ' GPP_CFGARTASSIST="",'
                  + ' GPP_CHAINAGE="",'
                  + ' GPP_CODEPIECEOBL1="-",'
                  + ' GPP_CODEPIECEOBL2="-",'
                  + ' GPP_CODEPIECEOBL3="-",'
                  + ' GPP_CODPIECEDEF1="",'
                  + ' GPP_CODPIECEDEF2="",'
                  + ' GPP_CODPIECEDEF3="",'
                  + ' GPP_COMMENTENT="",'
                  + ' GPP_COMMENTPIED="",'
                  + ' GPP_COMPANALLIGNE="DEM",'
                  + ' GPP_COMPANALPIED="DEM",'
                  + ' GPP_COMPSTOCKLIGNE="SAN",'
                  + ' GPP_COMPSTOCKPIED="SAN",'
                  + ' GPP_CONDITIONTARIF="X",'
                  + ' GPP_CONTEXTES="AFF;GC;",'
                  + ' GPP_CONTRECHART1="-",'
                  + ' GPP_CONTRECHART2="-",'
                  + ' GPP_CONTRECHART3="-",'
                  + ' GPP_CONTREMARQUE="-",'
                  + ' GPP_CONTREMREF="-",'
                  + ' GPP_CONTROLEMARGE="AUC",'
                  + ' GPP_CPTCENTRAL="-",'
                  + ' GPP_CRC=-748475751,'
                  + ' GPP_CUMULART1="-",'
                  + ' GPP_CUMULART2="-",'
                  + ' GPP_CUMULART3="-",'
                  + ' GPP_CUMULCOM1="-",'
                  + ' GPP_CUMULCOM2="-",'
                  + ' GPP_CUMULCOM3="-",'
                  + ' GPP_CUMULTIERS1="-",'
                  + ' GPP_CUMULTIERS2="-",'
                  + ' GPP_CUMULTIERS3="-",'
                  + ' GPP_DATELIBART1="AUC",'
                  + ' GPP_DATELIBART2="AUC",'
                  + ' GPP_DATELIBART3="AUC",'
                  + ' GPP_DATELIBCOM1="",'
                  + ' GPP_DATELIBCOM2="",'
                  + ' GPP_DATELIBCOM3="",'
                  + ' GPP_DATELIBTIERS1="AUC",'
                  + ' GPP_DATELIBTIERS2="AUC",'
                  + ' GPP_DATELIBTIERS3="AUC",'
                  + ' GPP_DIMSAISIE="TOU",'
                  + ' GPP_DUPLICPIECE="",'
                  + ' GPP_ECLATEAFFAIRE="-",'
                  + ' GPP_ECLATEDOMAINE="-",'
                  + ' GPP_EDITIONNOMEN="AUC",'
                  + ' GPP_ENCOURS="-",'
                  + ' GPP_EQUIPIECE="",'
                  + ' GPP_ESTAVOIR="-",'
                  + ' GPP_ETATETIQ="",'
                  + ' GPP_FAR_FAE="",'
                  + ' GPP_FILTREARTCH="",'
                  + ' GPP_FILTREARTVAL="",'
                  + ' GPP_FILTRECOMM="",'
                  + ' GPP_FORCERUPTURE="-",'
                  + ' GPP_GEREARTICLELIE="AUT",'
                  + ' GPP_GEREECHEANCE="AUT",'
                  + ' GPP_GESTIONGRATUIT="-",'
                  + ' GPP_HISTORIQUE="X",'
                  + ' GPP_IFL1="004",'
                  + ' GPP_IFL2="002",'
                  + ' GPP_IFL3="",'
                  + ' GPP_IFL4="",'
                  + ' GPP_IFL5="",'
                  + ' GPP_IFL6="",'
                  + ' GPP_IFL7="",'
                  + ' GPP_IFL8="",'
                  + ' GPP_IMPAUTOBESCBN="-",'
                  + ' GPP_IMPAUTOETATCBN="-",'
                  + ' GPP_IMPBESOIN="-",'
                  + ' GPP_IMPETAT="",'
                  + ' GPP_IMPETIQ="-",'
                  + ' GPP_IMPIMMEDIATE="-",'
                  + ' GPP_IMPMODELE="GFF",'
                  + ' GPP_INFOPCEPRECCHX="",'
                  + ' GPP_INFOPIECEPREC="",'
                  + ' GPP_INFOSCOMPL="",'
                  + ' GPP_INFOSCPLPIECE="",'
                  + ' GPP_INITQTE="",'
                  + ' GPP_INITQTECRE=0,'
                  + ' GPP_INSERTLIG="X",'
                  + ' GPP_JOURNALCPTA="ACD",'
                  + ' GPP_LIBELLE="Facture fournisseur financière",'
                  + ' GPP_LIENAFFAIRE="-",'
                  + ' GPP_LIENTACHE="-",'
                  + ' GPP_LISTEAFFAIRE="AFSAISIEFFAC",'
                  + ' GPP_LISTESAISIE="GCSAISIEFAF",'
                  + ' GPP_LOT="-",'
                  + ' GPP_MAJINFOTIERS="X",'
                  + ' GPP_MAJPRIXVALO="",'
                  + ' GPP_MASQUERNATURE="-",'
                  + ' GPP_MENU="",'
                  + ' GPP_MESSAGEEDIIN="",'
                  + ' GPP_MESSAGEEDIOUT="",'
                  + ' GPP_MODEECHEANCES="RS",'
                  + ' GPP_MODEGROUPEPORT="CUM",'
                  + ' GPP_MODELEWORD="",'
                  + ' GPP_MODIFCOUT="-",'
                  + ' GPP_MODPLANIFIABLE="",'
                  + ' GPP_MONTANTMINI=0,'
                  + ' GPP_MONTANTVISA=0,'
                  + ' GPP_MULTIGRILLE="-",'
                  + ' GPP_NATPIECEANNUL="",'
                  + ' GPP_NATURECPTA="FF",'
                  + ' GPP_NATUREORIGINE="",'
                  + ' GPP_NATUREPIECEG="FFF",'
                  + ' GPP_NATUREREPRISE="",'
                  + ' GPP_NATURESUIVANTE="",'
                  + ' GPP_NATURETIERS="FOU;",'
                  + ' GPP_NBEXEMPLAIRE=1,'
                  + ' GPP_NIVEAUPARAM="EDI",'
                  + ' GPP_NUMEROSERIE="-",'
                  + ' GPP_OBJETDIM="-",'
                  + ' GPP_OBLIGEREGLE="-",'
                  + ' GPP_OUVREAUTOPORT="-",'
                  + ' GPP_PARAMDIM="-",'
                  + ' GPP_PARAMGRILLEDIM="-",'
                  + ' GPP_PIECEEDI="-",'
                  + ' GPP_PIECEPILOTE="-",'
                  + ' GPP_PIECESAV="",'
                  + ' GPP_PIECETABLE1="",'
                  + ' GPP_PIECETABLE2="",'
                  + ' GPP_PIECETABLE3="",'
                  + ' GPP_PILOTEORDRE="-",'
                  + ' GPP_PREVUAFFAIRE="-",'
                  + ' GPP_PRIORECHART1="ART",'
                  + ' GPP_PRIORECHART2="REF",'
                  + ' GPP_PRIORECHART3="AUC",'
                  + ' GPP_PRIXNULOK="-",'
                  + ' GPP_PROCLI="-",'
                  + ' GPP_QTEMOINS="",'
                  + ' GPP_QTEPLUS="",'
                  + ' GPP_QUALIFMVT="",'
                  + ' GPP_RACINELIBECR1="",'
                  + ' GPP_RACINELIBECR2="",'
                  + ' GPP_RACINEREFINT1="",'
                  + ' GPP_RACINEREFINT2="",'
                  + ' GPP_RECALCULPRIX="X",'
                  + ' GPP_RECHTARIF501="-",'
                  + ' GPP_RECUPPRE="PRE",'
                  + ' GPP_REFEXTCTRL="000",'
                  + ' GPP_REFINTCTRL="000",'
                  + ' GPP_REFINTEXT="INT",'
                  + ' GPP_REGROUPCPTA="AUC",'
                  + ' GPP_REGROUPE="X",'
                  + ' GPP_RELIQUAT="-",'
                  + ' GPP_REPRISEENTAFF="",'
                  + ' GPP_REPRISELIGAFF="",'
                  + ' GPP_SENSPIECE="ENT",'
                  + ' GPP_SOLDETRANSFO="",'
                  + ' GPP_SOUCHE="GFF",'
                  + ' GPP_STKQUALIFMVT="",'
                  + ' GPP_TARIFGENDATE="010",'
                  + ' GPP_TARIFGENDEPOT="-",'
                  + ' GPP_TARIFGENPNPB="",'
                  + ' GPP_TARIFGENSAISIE="010",'
                  + ' GPP_TARIFGENSPECIA="010",'
                  + ' GPP_TARIFGENTRANSF="010",'
                  + ' GPP_TARIFMODULE="101",'
                  + ' GPP_TAXE="X",'
                  + ' GPP_TIERS="",'
                  + ' GPP_TRSFACHAT="-",'
                  + ' GPP_TRSFVENTE="-",'
                  + ' GPP_TYPEACTIVITE="",'
                  + ' GPP_TYPEARTICLE="CNS;MAR;",'
                  + ' GPP_TYPECOMMERCIAL="",'
                  + ' GPP_TYPEDIMOBLI1="-",'
                  + ' GPP_TYPEDIMOBLI2="-",'
                  + ' GPP_TYPEDIMOBLI3="-",'
                  + ' GPP_TYPEDIMOBLI4="-",'
                  + ' GPP_TYPEDIMOBLI5="-",'
                  + ' GPP_TYPEECRALIM="",'
                  + ' GPP_TYPEECRCPTA="NOR",'
                  + ' GPP_TYPEECRSTOCK="RIE",'
                  + ' GPP_TYPEFACT="",'
                  + ' GPP_TYPEPASSACC="REE",'
                  + ' GPP_TYPEPASSACCR="AUC",'
                  + ' GPP_TYPEPASSCPTA="ZDI",'
                  + ' GPP_TYPEPASSCPTAR="AUC",'
                  + ' GPP_TYPEPRESDOC="DEF",'
                  + ' GPP_TYPEPRESENT=0,'
                  + ' GPP_VALEURLIBECR1="",'
                  + ' GPP_VALEURLIBECR2="",'
                  + ' GPP_VALEURREFINT1="",'
                  + ' GPP_VALEURREFINT2="",'
                  + ' GPP_VALLIBART1="AUC",'
                  + ' GPP_VALLIBART2="AUC",'
                  + ' GPP_VALLIBART3="AUC",'
                  + ' GPP_VALLIBCOM1="",'
                  + ' GPP_VALLIBCOM2="",'
                  + ' GPP_VALLIBCOM3="",'
                  + ' GPP_VALLIBTIERS1="AUC",'
                  + ' GPP_VALLIBTIERS2="AUC",'
                  + ' GPP_VALLIBTIERS3="AUC",'
                  + ' GPP_VALMODELE="-",'
                  + ' GPP_VENTEACHAT="ACH",'
                  + ' GPP_VISA="-" WHERE GPP_NATUREPIECEG="FFF"')

    end;
    *)
  end;
  //S MASSON 3954
  ExecuteSqlNoPcl('UPDATE DPAGRICOLE SET DAG_509AUTRES="-"');

End;

Procedure MajVer937;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //MNG 3977
  // mise à jour dh_controle de la table WPARC pour mise à jour en série

  ExecuteSqlNoPCL('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Z" WHERE DH_PREFIXE ="WPC"'
                         + ' AND (DH_NOMCHAMP LIKE "%LIBRE%" OR DH_NOMCHAMP = "WPC_LIBELLE" OR DH_NOMCHAMP = "WPC_DESCRIPTION")') ;
  ExecuteSqlNoPCL ('UPDATE DECHAMPS SET DH_CONTROLE="LDV" WHERE DH_NOMCHAMP = "WPC_IDENTIFIANT"');

  // initialisation du nouveau champ rac_codeciblage
//  ExecuteSqlNoPCL ('UPDATE ACTIONS SET RAC_CODECIBLAGE ="" ');


  // initialisation du nouveau champ rop_codeciblage
//  ExecuteSqlNoPCL ('UPDATE OPERATIONS SET ROP_CODECIBLAGE ="" WHERE ROP_CODECIBLAGE IS NULL');

  //D SCLAVOPOULOS 3995
//  ExecutesqlNoPCL ('UPDATE PROFILART SET GPF_QPCBCONSO=0 WHERE GPF_QPCBCONSO IS NULL');
(*
  if not ExisteSQL('SELECT 1 FROM WCHAMP WHERE WCA_CONTEXTEPROFIL="PRO" AND WCA_NOMTABLE="ARTICLE" AND WCA_NOMCHAMP="GA_QPCBCONSO"') then
  begin
    ExecuteSQLNoPCL ('INSERT INTO WCHAMP (WCA_CONTEXTEPROFIL, WCA_NOMTABLE, WCA_NOMCHAMP)'
                   + ' VALUES ("PRO", "ARTICLE", "GA_QPCBCONSO")');
  end;
*)
  //Parwez COAN 4002
//  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "1"  WHERE   DH_NOMCHAMP = "GA_DELAIACH"    AND DH_CONTROLE NOT LIKE "%1%" ');

  //D KOZA 4017
//  ExecuteSqlNoPCL('UPDATE WNATURETRAVAIL SET WNA_FONCTION="1" WHERE WNA_FONCTION IS NULL');
//  ExecuteSqlNoPCL('UPDATE QSAISON SET QSA_ORIGINQT="" WHERE QSA_ORIGINQT IS NULL');

  // M MORRETTON 4039
//  ExecuteSQLNoPCL('UPDATE YTARIFSPROFILSC SET YPO_QUELTIERS="" WHERE YPO_QUELTIERS IS NULL');
//  ExecuteSQLNoPCL('UPDATE YTARIFSPROFILSD SET YPT_TARIFSPECIAL="" WHERE YPT_TARIFSPECIAL IS NULL');
//  ExecuteSQLNoPCL('UPDATE YTARIFS SET YTS_CATALOGUE="", YTS_CASCCATA="-", YTS_CASCTOUSCATA="-", YTS_PRESCRIPTEUR="", YTS_CASCPRES="-", YTS_CASCTOUSPRES="-", YTS_APPORTEUR="", YTS_CASCAPPO="-", YTS_CASCTOUSAPPO="-" WHERE YTS_CATALOGUE IS NULL');
//  ExecuteSQLNoPCL('UPDATE YTARIFSPARAMETRES SET YFO_OKCATALOGUE="----", YFO_OKPRESCRIPTEUR="----", YFO_OKAPPORTEUR="----", YFO_PRIORITERECH="" WHERE YFO_OKCATALOGUE IS NULL');
//  ExecuteSQLNoPCL('UPDATE WPARAM SET WPA_VARCHAR20="", WPA_VARCHAR21="", WPA_VARCHAR22="", WPA_VARCHAR23="", WPA_VARCHAR24="", WPA_VARCHAR25="" WHERE WPA_VARCHAR20 IS NULL');


End;

Procedure MajVer938;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //M GUERIN 4051
  {Initialisation du nouveau champ GP_HISTORISE}
//  ExecuteSqlNoPCL('update PIECE set GP_HISTORISE = "-" where isnull(GP_HISTORISE,"")=""');
  // DBR 4054
//  ExecuteSQLNoPCL ('UPDATE LIGNEFRAIS SET LF_FAMILLETAXE="", LF_FRAISINDPR="'
//  + GetParamSocSecur ('SO_FRAISINDPR', 'NON') + '" WHERE LF_FAMILLETAXE IS NULL');
  //D SCLAVOPOULOS 4071
//  ExecuteSQLContOnExcept('DELETE FROM PARAMSOC WHERE SOC_NOM = "SO_WINITLOTINTERNE"');
  //D KOZA 4077
//  ExecuteSqlNoPCL('UPDATE QCIRCUIT SET QCI_TYPENOMEN="" WHERE QCI_TYPENOMEN IS NULL');
//  ExecuteSqlNoPCL('UPDATE QFINSPE SET QFS_CARACTFS="" WHERE QFS_CARACTFS IS NULL');
//  ExecuteSqlNoPCL('UPDATE QPLANAFF SET QPL_COLORIS="" WHERE QPL_COLORIS IS NULL');
//  ExecuteSqlNoPCL('UPDATE QQTMINCAR SET QQM_NBARTMAX=0, QQM_PREFIXEREGT="", QQM_QTMINMIN=0, QQM_SEUILNBJ=0 WHERE QQM_NBARTMAX IS NULL');
//  ExecuteSqlNoPCL('UPDATE QTEMPP SET QTP_CONTENU="", QTP_REGROUPEMENT="" WHERE QTP_CONTENU IS NULL');
//  ExecuteSqlNoPCL('UPDATE QTRANDETD SET QTD_INFOTRANCHE="" WHERE QTD_INFOTRANCHE IS NULL');
  //G JUGDE 4081
  ExecuteSQLNoPCL('UPDATE JOURNEETYPE SET PJO_TYPEHORAIRE= "", PJO_PAUSEMIN="' + UsDateTime_(IDate1900) + '",PJO_PAUSEMAX="' + UsDateTime_(IDate1900)
  + '",PJO_DUREEREPOS="' + UsDateTime_(IDate1900) + '",PJO_DEBINTERDIT1="' + UsDateTime_(IDate1900) + '",PJO_FININTERDIT1="' + UsDateTime_(IDate1900)
  + '",PJO_DEBINTERDIT2="' + UsDateTime_(IDate1900) + '",PJO_FININTERDIT2="' + UsDateTime_(IDate1900) + '",PJO_DEBINTERDIT3="' + UsDateTime_(IDate1900)
  + '",PJO_FININTERDIT3="' + UsDateTime_(IDate1900) + '",PJO_ARRONDIHOR1=0,PJO_ARRONDIPAUSE1=0,PJO_ARRONDIPAUSE2=0,PJO_ARRONDIHOR2=0,PJO_TOLEREHOR1=0'
  + ',PJO_TOLEREPAUSE1=0,PJO_TOLEREPAUSE2=0,PJO_TOLEREHOR2=0,PJO_ZNEUTREHOR1=0,PJO_ZNEUTREHOR2=0,PJO_DUREELEGALE="' + UsDateTime_(IDate1900)
  + '",PJO_HEUREMINHOR1="' + UsDateTime_(IDate1900) + '",PJO_HEUREMAXHOR1="' + UsDateTime_(IDate1900) + '",PJO_HEUREMINHOR2="' + UsDateTime_(IDate1900)
  + '",PJO_HEUREMAXHOR2="' + UsDateTime_(IDate1900) + '",PJO_TYPEARRONDI="",PJO_AFFECTJOUR="-",PJO_DEBTOLEREEX=0,PJO_FINTOLEREEX=0,PJO_DEBARRONDIEX=0'
  + ',PJO_FINARRONDIEX=0  WHERE PJO_ARRONDIHOR1 IS NULL');
(*
  ExecuteSQLNoPCL('UPDATE QOPOINTAGE SET QOP_SALARIE="", QOP_JOURNEETYPE="",QOP_DATEAFFECTEE="' + UsDateTime_(IDate1900) + '",QOP_POSTE=""'
  + ',QOP_DEBCAPTPT="' + UsDateTime_(IDate1900)+ '",QOP_DEBJOURREL="' + UsDateTime_(IDate1900) + '",QOP_FINJOURREL="' + UsDateTime_(IDate1900) + '",QOP_DEBP1REL="' + UsDateTime_(IDate1900) + '",QOP_FINP1REL="' + UsDateTime_(IDate1900)
  + '",QOP_DEBP2REL="' + UsDateTime_(IDate1900) + '",QOP_FINP2REL="' + UsDateTime_(IDate1900) + '",QOP_DEBABS1REL="' + UsDateTime_(IDate1900)
  + '",QOP_FINABS1REL="' + UsDateTime_(IDate1900) + '",QOP_DEBABS2REL="' + UsDateTime_(IDate1900)+ '",QOP_FINABS2REL="' + UsDateTime_(IDate1900)
  + '",QOP_DEBABS3REL="' + UsDateTime_(IDate1900) + '",QOP_FINABS3REL="' + UsDateTime_(IDate1900) + '",QOP_DEBJOURRET="' + UsDateTime_(IDate1900)
  + '",QOP_FINJOURRET="' + UsDateTime_(IDate1900) + '",QOP_DEBP1RET="' + UsDateTime_(IDate1900) + '",QOP_FINP1RET="' + UsDateTime_(IDate1900)
  + '",QOP_DEBP2RET="' + UsDateTime_(IDate1900) + '",QOP_FINP2RET="' + UsDateTime_(IDate1900) + '",QOP_DEBABS1RET="' + UsDateTime_(IDate1900)
  + '",QOP_FINABS1RET="' + UsDateTime_(IDate1900) + '",QOP_DEBABS2RET="' + UsDateTime_(IDate1900)+ '",QOP_FINABS2RET="' + UsDateTime_(IDate1900)
  + '",QOP_DEBABS3RET="' + UsDateTime_(IDate1900) + '",QOP_FINABS3RET="' + UsDateTime_(IDate1900)+ '",QOP_CUMABSREL=0,QOP_CUMABSRET=0,QOP_CUMPAUSEREL=0,QOP_CUMPAUSERET=0'
  + ',QOP_TPRESTHEO=0,QOP_TPRESREEL=0,QOP_COMMENTAIRE="", QOP_FINJOUR="-"  WHERE QOP_CUMABSREL IS NULL');
  //G JUGDE 4082
  ExecuteSQLNoPCL ( 'UPDATE QWHISTORES SET QWH_POSTE="",QWH_SITESVT="",QWH_GRPSVT="" WHERE QWH_POSTE IS NULL');
  ExecuteSQLNoPCL ( 'UPDATE QWBACTET SET QWB_QENTGRPSTOC=0,QWB_QSORGRPSTOC=0 WHERE QWB_QENTGRPSTOC IS NULL');
  ExecuteSQLNoPCL ( 'UPDATE QWBACGAMME SET QWG_POSTE="" WHERE QWG_POSTE IS NULL');
*)
  //D KOZA 4085
//  ExecuteSqlNoPCL('UPDATE WCBNEVOLUTION SET WEV_CHOIXQUALITE="" WHERE WEV_CHOIXQUALITE IS NULL');

  //G JUGDE 4091
  ExecuteSQLNoPCL ( 'UPDATE RESSOURCE SET ARS_TYPETRAVAIL="" WHERE ARS_TYPETRAVAIL IS NULL');
//  ExecuteSQLNoPCL ( 'UPDATE QGROUPE SET QGR_GRPMES="X" WHERE QGR_GRPMES IS NULL');

End;

Procedure MajVer939;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //M MORETTON 4001
//  ExecuteSQLNoPCL('UPDATE CATALOGU SET GCA_PRIXSIMULACH=0 WHERE GCA_PRIXSIMULACH IS NULL');

  //M DESGOUTTE 4125
  // Rectification forme juridique sté civile exploitation agricole
  ExecuteSQLContOnExcept('update choixcod set cc_libelle="SCEA" where cc_type="JUR" and  cc_code="SCG" and cc_libelle="SCGA"');

  // Suppression de paramsoc obsolètes
  If IsMonoOuCommune then
    ExecuteSQLContOnExcept('delete FROM paramsoc where soc_nom in ("SO_OBGDATEGEN", "SO_OBGVERSION", '
    +'"SO_OBGPERIODICITE", "SO_OBGPASALERTE", "SO_OBGTELECHARGEDIRECT")');

  //MCD 4127
  //modif pour Lenotre ... qui aura eu ces modif, donc obligation de faire l'initialisation champ par champ
  //1 - TABLE AFPLANNING
 // ExecuteSQLNoPCL(' UPDATE AFPLANNING SET APL_QTEPLAINTERV = 1 WHERE APL_QTEPLAINTERV IS NULL ');
 // ExecuteSQLNoPCL(' UPDATE AFPLANNING SET APL_QTEINITINTERV = APL_QTEPLAINTERV WHERE APL_QTEINITINTERV IS NULL ');
 // ExecuteSQLNoPCL(' UPDATE AFPLANNING SET APL_DUREE = APL_QTEPLANIFIEE WHERE APl_DUREE IS NULL ');
 // ExecuteSQLNoPCL(' UPDATE AFPLANNING SET APL_DUREEUREF = APL_QTEPLANIFUREF WHERE APL_DUREEUREF IS NULL ');
 // ExecuteSQLNoPCL(' UPDATE AFPLANNING SET APL_DATESUPP1 = "' + UsDateTime_(iDate1900) + '" WHERE APL_DATESUPP1 IS NULL ');
 // ExecuteSQLNoPCL(' UPDATE AFPLANNING SET APL_DATESUPP2 = "' + UsDateTime_(iDate1900) + '" WHERE APL_DATESUPP2 IS NULL ');
 // ExecuteSQLNoPCL(' UPDATE AFPLANNING SET APL_DATESUPP3 = "' + UsDateTime_(iDate1900) + '" WHERE APL_DATESUPP3 IS NULL ');
 // ExecuteSQLNoPCL(' UPDATE AFPLANNING SET APL_DATESUPP4 = "' + UsDateTime_(iDate1900) + '" WHERE APL_DATESUPP4 IS NULL ');
 (*
    ExecuteSQLNoPCL(' UPDATE AFPLANNING SET '
    +' APL_QTEPLAINTERV = 1'
    +',APL_QTEINITINTERV = 1'
    +',APL_DUREE = APL_QTEPLANIFIEE'
    +',APL_DUREEUREF = APL_QTEPLANIFUREF'
    +',APL_DATESUPP1 = "' + UsDateTime_(iDate1900) + '"'
    +',APL_DATESUPP2 = "' + UsDateTime_(iDate1900) + '"'
    +',APL_DATESUPP3 = "' + UsDateTime_(iDate1900) + '"'
    +',APL_DATESUPP4 = "' + UsDateTime_(iDate1900) + '"'
    +'  WHERE APl_DUREE IS NULL ');
	*)
  //2 - TABLE RESSOURCE
//  ExecuteSqlContOnExcept(' UPDATE RESSOURCE SET ARS_CAPACITE = 0 WHERE ARS_CAPACITE IS NULL ');
//  ExecuteSqlContOnExcept(' UPDATE RESSOURCE SET ARS_CAPACITEMAX = 0 WHERE ARS_CAPACITEMAX IS NULL ');
//  ExecuteSqlContOnExcept(' UPDATE RESSOURCE SET ARS_PRETE = "" WHERE ARS_PRETE IS NULL ');
//  ExecuteSqlContOnExcept(' UPDATE RESSOURCE SET ARS_SITEGEO = "" WHERE ARS_SITEGEO IS NULL ');
//  ExecuteSqlContOnExcept(' UPDATE RESSOURCE SET ARS_PLANIFMASSE = "-" WHERE ARS_PLANIFMASSE IS NULL ');
//  ExecuteSqlContOnExcept(' UPDATE RESSOURCE SET ARS_PLAAFFECT = "-" WHERE ARS_PLAAFFECT IS NULL ');

  ExecuteSqlContOnExcept(' UPDATE RESSOURCE SET  '
  + 'ARS_CAPACITE = 0,ARS_CAPACITEMAX = 0,  ARS_PRETE = "" ,ARS_SITEGEO = "", ARS_PLANIFMASSE = "-" ,ARS_PLAAFFECT = "-" '
  + ' WHERE ARS_PLANIFMASSE IS NULL ');


  //3 - TABLE AFPLANNINGPARAM
  (*
  ExecuteSQLNoPCL(' UPDATE AFPLANNINGPARAM SET APP_CODEINTERNE = "" WHERE APP_CODEINTERNE IS NULL ');
  ExecuteSQLNoPCL(' UPDATE AFPLANNINGPARAM SET APP_LIBELLEREDUIT = "" WHERE APP_LIBELLEREDUIT IS NULL ');
  ExecuteSQLNoPCL(' UPDATE AFPLANNINGPARAM SET APP_USERSGROUP = "" WHERE APP_USERSGROUP IS NULL ');
  ExecuteSQLNoPCL(' UPDATE AFPLANNINGPARAM SET APP_GROUPRIGHTS = "" WHERE APP_GROUPRIGHTS IS NULL ');
  *)
  //4 - TABLE AFFAIRE
//  ExecuteSQLNoPCL(' UPDATE AFFAIRE SET AFF_EVENEMENT = "" WHERE AFF_EVENEMENT IS NULL '); // GM Attention Lenotre
  //5 - TABLE AFCODEAFFAIRE
//  ExecuteSQLNoPCL(' UPDATE AFCODEAFFAIRE SET ACM_MODEAFFAIRE="ECO" WHERE ACM_ISECOLE="X" AND ACM_MODEAFFAIRE IS NULL ');
//  ExecuteSQLNoPCL(' UPDATE AFCODEAFFAIRE SET ACM_MODEAFFAIRE="PLA" WHERE (ACM_ISECOLE="" OR ACM_ISECOLE="-") AND ACM_MODEAFFAIRE IS NULL ');
  //6 - TABLE PARPIECE
  (*
  ExecuteSQLNoPCL(' UPDATE PARPIECE SET GPP_PLANNINGFLUX="TRF" WHERE GPP_NATUREPIECEG="CC" AND GPP_PLANNINGFLUX IS NULL ');
  ExecuteSQLNoPCL(' UPDATE PARPIECE SET GPP_PLANNINGFLUX="FAC" WHERE GPP_NATUREPIECEG="FAC" AND GPP_PLANNINGFLUX IS NULL ');
  ExecuteSQLNoPCL(' UPDATE PARPIECE SET GPP_PLANNINGFLUX="" WHERE GPP_PLANNINGFLUX IS NULL AND GPP_NATUREPIECEG<>"CC" AND GPP_NATUREPIECEG<>"FAC" ');
  ExecuteSQLNoPCL(' UPDATE PARPIECE SET GPP_ETATAFFAIRE="" WHERE GPP_ETATAFFAIRE IS NULL ');
  ExecuteSQLNoPCL(' UPDATE PARPIECE SET GPP_PLAETATCREAT="" WHERE GPP_PLAETATCREAT IS NULL ');
  ExecuteSQLNoPCL(' UPDATE PARPIECE SET GPP_PLAETATSOLDE="" WHERE GPP_PLAETATSOLDE IS NULL ');
  *)
  //S BOUSSERT 4107
  // Initialisation des nouveaux champs de la table FEMPRUNT
  ExecuteSQLContOnExcept('UPDATE FEMPRUNT SET EMP_MODIFCPTIC="-", EMP_MODIFCPTCA="-", EMP_MODIFCPTFF="-", EMP_MODIFCPTAS="-"')

End;

Procedure MajVer940;
Begin
  //T SUBLET 4053
  { moulinette permettant le passage des "jetons" aux "séquences CBP" }
  if not isDossierPCL then
    SwitchToSequences();

  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //JA VORACHACK
//  ExecuteSQLNoPCL(' UPDATE AFPLANNINGPARAM SET APP_LIBELLEREDUIT = "" WHERE APP_LIBELLEREDUIT IS NULL ');
  //T SUBLET 4053
  { Table ARTICLETIERS (oublie 938) }
  (*
  ExecuteSqlNoPCL('UPDATE ARTICLETIERS'
               + ' SET GAT_REFGROUP="-"'
               + ', GAT_CODEEDI=""'
               + ', GAT_STATUTPREV=""'
               + ', GAT_ARTSECU="-"'
               + ', GAT_ARTREGL="-"'
               + ', GAT_ARTAQP="-"'
               + ', GAT_REFUM=""'
               + ', GAT_REFUC=""'
               + ', GAT_QTEARTUC=0'
               + ', GAT_QTEUCUM=0'
               + ', GAT_MODELEUM=""'
               + ', GAT_MODELEUC=""'
               + ', GAT_NBETIQUM=0'
               + ', GAT_NBETIQUC=0'
               + ', GAT_RANGEMENTCOLIS=""'
               + ', GAT_DELAIMOYEN=0'
               + ', GAT_CODEBARRE=""'
               + ', GAT_QUALIFCODEBARRE=""'
               + ', GAT_WBMEMO="-"'
               + ', GAT_BLOCNOTE=""'
               );
  *)
  //M MORRETTON 4138
  ExecuteSQLNoPCL('UPDATE DEVISE SET '
      +'D_ARRONDIPRIXACHAT=(SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM = "SO_DECPRIX"), '
      +'D_ARRONDIPRIXVENTE=(SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM = "SO_DECPRIX") '
      +'WHERE D_DEVISE=(SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM = "SO_DEVISEPRINC")');
  //O BERTHIER 4139
//  ExecuteSQLNoPCL('update afsegmissionpotent set asm_reglecritere=""');

  //J TRIFILIEFF 4143
//  if not isDossierPCL then InitNaturePceQteMini;

  //S MASSON 4149
  ExecuteSqlNoPcl('DELETE FROM CHOIXDPSTD WHERE YDS_NODOSSIER="000000" AND YDS_TYPE = "DAG" '
  +'AND (YDS_CODE LIKE "605%" OR YDS_CODE LIKE "607%" OR YDS_CODE LIKE "608%")');

  //M GUERIN 4151
  {Initialisation du nouveau champ ajouté à la table}
//  ExecuteSQLNoPCL('update ETABLES set EDT_CONTEXTE="-" where ISNULL(EDT_CONTEXTE,"")=""');

  //S MASSON 4159
  ExecuteSqlNoPcl('UPDATE CHOIXDPSTD SET YDS_CODE = YDS_CODE||"00" WHERE LEN(YDS_CODE) = 6 '
  +'AND YDS_TYPE = "DAG" AND (YDS_CODE NOT LIKE "605%" AND YDS_CODE NOT LIKE "607%" AND YDS_CODE NOT LIKE "608%" AND YDS_CODE NOT LIKE "N%")');

  ExecuteSqlNoPcl('UPDATE CSUIVICULTURE SET SVC_CODECULTURE = SVC_CODECULTURE||"00" WHERE LEN(SVC_CODECULTURE) = 6');

  //D SCLAVOPOULOS 4168
//  ExecuteSQLNoPCL('UPDATE WPARAMFONCTION SET WPF_BOOLLIBREART1="-", WPF_BOOLLIBREART2="-", WPF_BOOLLIBREART3="-" WHERE WPF_BOOLLIBREART1 IS NULL');

  //MNG 4171
  // MNG ajout des codes pour les nouveaux champs infos compl actions
//  RT_InsertLibelleInfoComplAction2();

  //B MERIAUX 4172
  if not isDossierPCL then ExecuteSQL('update annuaire set ann_sigle = ""');

  //MCD 4183
  //GIGA CB + JVO
// Table Tache

  // J VORACHACK, l'ATA_IDENTLIGNE de la table tache peut contenir l'ancien format du refpiece
  {*ExecuteSQLNoPCL('UPDATE TACHE SET ATA_NATUREPIECEG = SUBSTRING(ATA_IDENTLIGNE,1,3) WHERE ATA_IDENTLIGNE <> "" AND ATA_NATUREPIECEG IS NULL');
  ExecuteSQLNoPCL('UPDATE TACHE SET ATA_SOUCHE=SUBSTRING(ATA_IDENTLIGNE,5,3) WHERE ATA_IDENTLIGNE <> "" AND ATA_SOUCHE IS NULL');
  ExecuteSQLNoPCL('UPDATE TACHE SET ATA_NUMERO=CAST(SUBSTRING(ATA_IDENTLIGNE,9,9) AS INTEGER) WHERE ATA_IDENTLIGNE <> "" AND ATA_NUMERO IS NULL');
  ExecuteSQLNoPCL('UPDATE TACHE SET ATA_INDICEG=CAST(SUBSTRING(ATA_IDENTLIGNE, 19, 3) AS INTEGER) WHERE ATA_IDENTLIGNE <> "" AND ATA_INDICEG IS NULL');
  ExecuteSQLNoPCL('UPDATE TACHE SET ATA_NUMORDRE=CAST(SUBSTRING(ATA_IDENTLIGNE, 23, 6) AS INTEGER) WHERE ATA_IDENTLIGNE <> "" AND ATA_NUMORDRE IS NULL');*}
  (*
  ExecuteSQLNoPCL(' UPDATE TACHE SET ATA_NATUREPIECEG=SUBSTRING(ATA_IDENTLIGNE,1,3), '
                                  + 'ATA_SOUCHE=SUBSTRING(ATA_IDENTLIGNE,5,3), '
                                  + 'ATA_NUMERO=CAST(SUBSTRING(ATA_IDENTLIGNE,9,9) AS INTEGER), '
                                  + 'ATA_INDICEG=CAST(SUBSTRING(ATA_IDENTLIGNE, 19, 3) AS INTEGER), '
                                  + 'ATA_NUMORDRE=CAST(SUBSTRING(ATA_IDENTLIGNE, 23, 6) AS INTEGER) '
                                  + 'WHERE ATA_IDENTLIGNE <> "" AND ATA_NATUREPIECEG IS NULL '
                                  + 'AND SUBSTRING(ATA_IDENTLIGNE,1,1) NOT IN ("0", "1", "2", "3", "4", "5", "6", "7", "8", "9") ');

  ExecuteSQLNoPCL('UPDATE TACHE SET ATA_NATUREPIECEG="", ATA_SOUCHE="", ATA_NUMERO=0, ATA_INDICEG=0,ATA_NUMORDRE=0 WHERE ATA_NATUREPIECEG IS NULL');

  // GA_200912_JVO_16743_DEBUT
  ExecuteSQLNoPCL('UPDATE TACHE SET ATA_QTEINTERVENT = '
  	                              + '(SELECT IIF(GLC_QTEAPLANIF = 0, 1, GLC_QTEAPLANIF) '
	                                + 'FROM LIGNECOMPL WHERE GLC_NATUREPIECEG=ATA_NATUREPIECEG AND GLC_SOUCHE=ATA_SOUCHE '
	                                + 'AND GLC_NUMERO=ATA_NUMERO AND GLC_INDICEG=ATA_INDICEG AND GLC_NUMORDRE=ATA_NUMORDRE) '
                  + 'WHERE ATA_TYPEPLANIF = "FOR" AND ATA_IDENTLIGNE <> ""');

  ExecuteSQLNoPCL('UPDATE TACHE SET ATA_QTEINITINTERV=IIF((ATA_TYPEPLANIF="FOR" AND ATA_QTEINTERVENT > 0), (ATA_QTEINITPLA/ATA_QTEINTERVENT), 1) '
                + 'WHERE ATA_QTEINITINTERV IS NULL');
  // GA_200912_JVO_16743_FIN
 // ExecuteSQLNoPCL('UPDATE TACHE SET ATA_QTEPLAINTERV = 0 WHERE ATA_QTEPLAINTERV IS NULL');
 // ExecuteSQLNoPCL('UPDATE TACHE SET ATA_PIECEPRECEDENTE="" WHERE ATA_PIECEPRECEDENTE IS NULL');
 // ExecuteSQLNoPCL('UPDATE TACHE SET ATA_HEUREFIN = "' + UsDateTime_(iDate1900) + '" WHERE ATA_HEUREFIN IS NULL');
  ExecuteSQLNoPCL('UPDATE TACHE SET ATA_PIECEPRECEDENTE="",ATA_QTEPLAINTERV = 0,ATA_HEUREFIN = "' + UsDateTime_(iDate1900) + '"' +' WHERE ATA_PIECEPRECEDENTE IS NULL');

  ExecuteSQLNoPCL('UPDATE TACHE SET ATA_UNITEFAC=(SELECT GA_QUALIFUNITEVTE FROM ARTICLE WHERE GA_ARTICLE=ATA_ARTICLE) WHERE ATA_UNITEFAC IS NULL');
  ExecuteSQLNoPCL('UPDATE TACHE SET ATA_NUMEROTACHEM = 0 WHERE ATA_NUMEROTACHEM IS NULL');
  // Table AFModeleTache
  ExecuteSQLNoPCL('UPDATE AFMODELETACHE SET AFM_HEUREDEBUT = "' + UsDateTime_(GetParamSocSecur('SO_AFAMDEBUT', '08:00:00')) + '"'
  +  ',AFM_HEUREFIN = "' + UsDateTime_(iDate1900) + '"'
  + ' WHERE AFM_HEUREDEBUT IS NULL');
  // ExecuteSQLNoPCL('UPDATE AFMODELETACHE SET AFM_HEUREFIN = "' + UsDateTime_(iDate1900) + '" WHERE AFM_HEUREFIN IS NULL');
  // Table AFFTiers

  ExecuteSQLNoPCL('UPDATE AFFTIERS SET AFT_ESTFACTURE = "-" WHERE AFT_ESTFACTURE IS NULL');
  ExecuteSQLNoPCL('UPDATE AFFTIERS SET AFT_NATUREPIECEG = "", AFT_SOUCHE = "", AFT_NUMERO = 0, AFT_INDICEG = 0, AFT_NUMORDRE = 0 WHERE AFT_NATUREPIECEG IS NULL');

  ExecuteSQLNoPCL('UPDATE AFHISTOPLANCHARGE SET AHP_INITIAL = "X", AHP_IDENTLIGNE = "", AHP_NUMEROTACHEM = 0 WHERE AHP_INITIAL IS NULL');
 // ExecuteSQLNoPCL('UPDATE LIGNECOMPLAFF SET GLA_FONCTION = "", GLA_UNITETEMPS = "J" WHERE GLA_FONCTION IS NULL'); // déplacé  en 949

  ExecuteSQLNoPCL('UPDATE AFFAIRE SET AFF_PDCPOURCENT = "-", aff_txrealisation=0,aff_dteeffettxreal="'+UsDateTime_(IDate1900)+'"'+' WHERE AFF_PDCPOURCENT IS NULL');
  *)
   //G JUGDE 4186
//  ExecuteSQLNoPCL ('UPDATE WGAMMELIG SET WGL_SITE="",WGL_DEPOT="",WGL_TIERS="",WGL_SOCIETEGROUPE="" WHERE WGL_SITE IS NULL');
(*
  ExecuteSQLNoPCL ('INSERT INTO WGAMMELIG '
  + '(WGL_NATURETRAVAIL, WGL_ARTICLE, WGL_MAJEUR, WGL_OPEITI, WGL_NUMOPERGAMME, WGL_IDENTIFIANT, '
  + 'WGL_GUID, WGL_CIRCUIT, WGL_SITE, WGL_DEPOT, WGL_TIERS, WGL_SOCIETEGROUPE, WGL_CODEOPERATION, '
  + 'WGL_LIBELLE, WGL_TYPEOPERATION, WGL_PHASE, WGL_PHASELIB, WGL_CODEARTICLE, WGL_SOCIETE, '
  + 'WGL_QLOTSAIS, WGL_UNITELOT, WGL_COEFLOT, WGL_QLOTSTOC, WGL_QUALIFUNITESTO, WGL_TYPEEMPLOI, '
  + 'WGL_TPREVSAIS, WGL_UNITEPREV, WGL_COEFPREV, WGL_TPREVHHCC, WGL_SECTIONPDR, WGL_RUBRIQUEPDR, '
  + 'WGL_FREINTE, WGL_LIBREWGL1, WGL_LIBREWGL2, WGL_LIBREWGL3, WGL_LIBREWGL4, WGL_LIBREWGL5, '
  + 'WGL_LIBREWGL6, WGL_LIBREWGL7, WGL_LIBREWGL8, WGL_LIBREWGL9, WGL_LIBREWGLA, WGL_VALLIBRE1, '
  + 'WGL_VALLIBRE2, WGL_VALLIBRE3, WGL_DATELIBRE1, WGL_DATELIBRE2, WGL_DATELIBRE3, WGL_BOOLLIBRE1, '
  + 'WGL_BOOLLIBRE2, WGL_BOOLLIBRE3, WGL_CHARLIBRE1, WGL_CHARLIBRE2, WGL_CHARLIBRE3, WGL_DQUNIC, '
  + 'WGL_RESSOURCESCM, WGL_TYPEOPE, WGL_DATEAPP, WGL_DATEPER, WGL_SUIVIOPMES, WGL_ANTERIORITEMES, '
  + 'WGL_GROUPAGEMES, WGL_DATECREATION, WGL_DATEMODIF, WGL_CREATEUR, WGL_UTILISATEUR, WGL_WBMEMO ) '
  + '(SELECT G1.WGL_NATURETRAVAIL, G1.WGL_ARTICLE, G1.WGL_MAJEUR, G1.WGL_OPEITI, G1.WGL_NUMOPERGAMME, G1.WGL_IDENTIFIANT, '
  + 'PgiGuid, "", G1.WGL_SITE, G1.WGL_DEPOT, G1.WGL_TIERS, G1.WGL_SOCIETEGROUPE, G1.WGL_CODEOPERATION, '
  + 'G1.WGL_LIBELLE, G1.WGL_TYPEOPERATION, G1.WGL_PHASE, G1.WGL_PHASELIB, G1.WGL_CODEARTICLE, G1.WGL_SOCIETE, '
  + 'G1.WGL_QLOTSAIS, G1.WGL_UNITELOT, G1.WGL_COEFLOT, G1.WGL_QLOTSTOC, G1.WGL_QUALIFUNITESTO, G1.WGL_TYPEEMPLOI, '
  + 'G1.WGL_TPREVSAIS, G1.WGL_UNITEPREV, G1.WGL_COEFPREV, G1.WGL_TPREVHHCC, G1.WGL_SECTIONPDR, G1.WGL_RUBRIQUEPDR, '
  + 'G1.WGL_FREINTE, G1.WGL_LIBREWGL1, G1.WGL_LIBREWGL2, G1.WGL_LIBREWGL3, G1.WGL_LIBREWGL4, G1.WGL_LIBREWGL5, '
  + 'G1.WGL_LIBREWGL6, G1.WGL_LIBREWGL7, G1.WGL_LIBREWGL8, G1.WGL_LIBREWGL9, G1.WGL_LIBREWGLA, G1.WGL_VALLIBRE1, '
  + 'G1.WGL_VALLIBRE2, G1.WGL_VALLIBRE3, G1.WGL_DATELIBRE1, G1.WGL_DATELIBRE2, G1.WGL_DATELIBRE3, G1.WGL_BOOLLIBRE1, '
  + 'G1.WGL_BOOLLIBRE2, G1.WGL_BOOLLIBRE3, G1.WGL_CHARLIBRE1, G1.WGL_CHARLIBRE2, G1.WGL_CHARLIBRE3, G1.WGL_DQUNIC, '
  + 'G1.WGL_RESSOURCESCM, G1.WGL_TYPEOPE, "' + UsDateTime_( IDate1900 ) + '" , "' + UsDateTime_( IDate1900 ) + '" , G1.WGL_SUIVIOPMES, G1.WGL_ANTERIORITEMES, '
  + 'G1.WGL_GROUPAGEMES, G1.WGL_DATECREATION, G1.WGL_DATEMODIF, G1.WGL_CREATEUR, G1.WGL_UTILISATEUR, G1.WGL_WBMEMO '
  + ' FROM WGAMMELIG G1 '
  + ' WHERE G1.WGL_CIRCUIT<>""  AND NOT EXISTS(SELECT 1 FROM WGAMMELIG G2 '
  + ' WHERE G1.WGL_NATURETRAVAIL=G2.WGL_NATURETRAVAIL  AND G1.WGL_ARTICLE=G2.WGL_ARTICLE '
  + ' AND G1.WGL_MAJEUR=G2.WGL_MAJEUR  AND G1.WGL_OPEITI=G2.WGL_OPEITI '
  + ' AND G1.WGL_NUMOPERGAMME=G2.WGL_NUMOPERGAMME  AND G2.WGL_CIRCUIT="") ) ');

  ExecuteSQLNoPCL ('UPDATE WGAMMERES SET WGR_SITE="", WGR_DEPOT="", WGR_TIERS="", WGR_SOCIETEGROUPE="" WHERE WGR_SITE IS NULL');
  ExecuteSQLNoPCL ('UPDATE WGAMMECIR SET WGC_SOCIETEGROUPE="" WHERE WGC_SOCIETEGROUPE IS NULL');

  //M GUERIN 4169
  {Initialisation du nouveau champ ajouté à la table}
  ExecuteSQLNoPCL('update ECHAMPS set EDH_LONGUEUR=0 where ISNULL(EDH_LONGUEUR,0)=0');
*)

End;

Procedure MajVer941;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //D KOZA 4188

//  ExecuteSqlNoPCL('UPDATE QTRANDETD SET QTD_QTP1=0,QTD_QTP2=0,QTD_QTP3=0,QTD_QTP4=0,QTD_QTP5=0,'
//  +'QTD_QTP6=0,QTD_QTP7=0,QTD_QTP8=0,QTD_QTP9=0,QTD_QTP10=0,QTD_QTP11=0,QTD_QTP12=0,QTD_QTP13=0,'
//  +'QTD_QTP14=0,QTD_QTP15=0,QTD_QTP16=0,QTD_QTP17=0,QTD_QTP18=0,QTD_QTP19=0,QTD_QTP20=0 WHERE QTD_QTP1 IS NULL');

  //M MORRETTON 4191
//  ExecuteSQLNoPCL('UPDATE LISTEINVLIG SET GIL_DPAGST=0, GIL_DPRGST=0 WHERE GIL_DPAGST IS NULL');

  //M MORRETTON 4193
//  ExecuteSQLNoPCL('UPDATE PORT SET GPO_FRAISINDPR="" WHERE GPO_FRAISINDPR IS NULL');

  //M DESGOUTTE 4206
  // Requêtes volontairement non protégées (par is null) pour améliorer les performances
  ExecuteSQLNoPcl ('UPDATE DOSSIER SET DOS_SWSACTIVE="-", DOS_SWSDATEACTIV="'+UsDateTime_(iDate1900)+'"');
  ExecuteSQLNoPcl ('UPDATE YALERTESWS SET YAS_GAN="-"');
  ExecuteSQLNoPcl ('UPDATE DPTABGENPAIE SET DT1_SWSENVOYE="-"');

  //MNG 4219
  // initialisation des nouveaux champs infos compl actions
  (*
  ExecuteSqlNoPcl('UPDATE RTINFOS001 SET RD1_RD1LIBTEXTE5="", RD1_RD1LIBTEXTE6="",RD1_RD1LIBTEXTE7="",RD1_RD1LIBTEXTE8="",'+
  'RD1_RD1LIBTEXTE9="", RD1_RD1LIBVAL5=0,RD1_RD1LIBVAL6=0,RD1_RD1LIBVAL7=0,RD1_RD1LIBVAL8=0,RD1_RD1LIBVAL9=0,'+
  'RD1_RD1LIBTABLE5="", RD1_RD1LIBTABLE6="", RD1_RD1LIBTABLE7="", RD1_RD1LIBTABLE8="", RD1_RD1LIBTABLE9="", '+
  'RD1_RD1LIBMUL5="", RD1_RD1LIBMUL6="", RD1_RD1LIBMUL7="", RD1_RD1LIBMUL8="", RD1_RD1LIBMUL9="", '+
  'RD1_RD1LIBBOOL5="-", RD1_RD1LIBBOOL6="-", RD1_RD1LIBBOOL7="-", RD1_RD1LIBBOOL8="-", RD1_RD1LIBBOOL9="-", '+
  'RD1_RD1LIBDATE5="'+UsDateTime_(iDate1900)+'", RD1_RD1LIBDATE6="'+UsDateTime_(iDate1900)+'", RD1_RD1LIBDATE7="'+UsDateTime_(iDate1900)+'", '+
  'RD1_RD1LIBDATE8="'+UsDateTime_(iDate1900)+'", RD1_RD1LIBDATE9="'+UsDateTime_(iDate1900)+'" '+
  'WHERE RD1_RD1LIBTEXTE5 IS NULL');
   *)
  //MCD 4220
  //ajout des enrgt dans une tablette CC qui doivent exister en clientèle
  InsertChoixCode('ACH', '001', 'Liste 1 infos bulle', 'AFHINTPLANN01', '') ;
  InsertChoixCode('ACH', '002', 'Liste 2 infos bulle', 'AFHINTPLANN02', '') ;
  InsertChoixCode('ACH', '003', 'Liste 3 infos bulle', 'AFHINTPLANN03', '') ;
  InsertChoixCode('ACI', '001', 'Liste 1 de la forme', 'AFITEMSPLANN01', '') ;
  InsertChoixCode('ACI', '002', 'Liste 2 de la forme', 'AFITEMSPLANN02', '') ;
  InsertChoixCode('ACI', '003', 'Liste 3 de la forme', 'AFITEMSPLANN03', '') ;
  InsertChoixCode('ACI', '004', 'Liste 4 de la forme', 'AFITEMSPLANN04', '') ;
  InsertChoixCode('ACI', '005', 'Liste 5 de la forme', 'AFITEMSPLANN05', '') ;
  InsertChoixCode('ACI', '006', 'Liste 6 de la forme', 'AFITEMSPLANN06', '') ;
  InsertChoixCode('ACI', '007', 'Liste 7 de la forme', 'AFITEMSPLANN07', '') ;
  InsertChoixCode('ACI', '008', 'Liste 8 de la forme', 'AFITEMSPLANN08', '') ;
  InsertChoixCode('ACI', '009', 'Liste 9 de la forme', 'AFITEMSPLANN09', '') ;
  InsertChoixCode('ACI', '010', 'Liste 10 de la forme', 'AFITEMSPLANN10', '') ;
  InsertChoixCode('ACI', '011', 'Liste 11 de la forme', 'AFITEMSPLANN11', '') ;
  InsertChoixCode('ACI', '012', 'Liste 12 de la forme', 'AFITEMSPLANN12', '') ;
  InsertChoixCode('ACI', '013', 'Liste 13 de la forme', 'AFITEMSPLANN13', '') ;
  InsertChoixCode('ACI', '014', 'Liste 14 de la forme', 'AFITEMSPLANN14', '') ;
  InsertChoixCode('ACI', '015', 'Liste 15 de la forme', 'AFITEMSPLANN15', '') ;

  //D SCLAVOPOULOS 4222
  ExecuteSqlNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Budget modifiable" WHERE DH_NOMCHAMP="AFF_OKMAJBUDGET"');

  //JP LAURENT 4201
//  ExecuteSQLNoPCL('update ligneda set dal_datelivraison="'+UsDateTime_(IDate1900)+'"  where dal_datelivraison is null')

End;

Procedure MajVer942;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //M FAUDEL 4185
  ExecuteSQLContOnExcept( 'UPDATE PROFILPAIE SET PPI_TYPPERSOURSSAF = "",PPI_TYPPERSOURSSAM = "", PPI_CONVENTION="000"');
  ExecuteSQLContOnExcept ('UPDATE CUMULPAIE SET PCL_TYPCOEFFCUM = "ELN" WHERE PCL_COEFFAFFECT <> ""');
  ExecuteSQLContOnExcept ('UPDATE CUMULPAIE SET PCL_TYPCOEFFCUM = "" WHERE PCL_COEFFAFFECT = ""');
//  ExecuteSQLContOnExcept ('UPDATE PUBLICOTIS SET PUO_METHODALIMDADS = "VAL"');
  ExecuteSQLContOnExcept ('UPDATE DADSLEXIQUE SET PDL_DADSRUPT="-"');
// on revient en arière 255 champs max en access  ExecuteSQLContOnExcept ('UPDATE SALARIES SET PSA_DOSTYPE="", PSA_PREDTYPE="", PSA_DATEVALTYPE=UsDateTime_(Idate1900), PSA_DATMAJTYPE=UsDateTime_(Idate1900), PSA_SALARIETYPE="", PSA_POSITIONCONV="", PSA_CATEGCONV="", PSA_DEGRECONV="", PSA_ECHELONCONV="", PSA_STATUTCONV="", PSA_DATEANCCONV=UsDateTime_(Idate1900)') ;

  //P DUMET 4225
  ExecuteSQLContOnExcept ('UPDATE CONTRATTRAVAIL SET PCI_DNASITUATADM="10",PCI_DNADATESIGNCEA="'+UsDateTime_(Idate1900)
  +'" ,PCI_DNADATERUPT="'+UsDateTime_(Idate1900)+'",PCI_DNAREFCEA="", PCI_DNAHORJOURN=0,PCI_DNADATEENGLIC="'+UsDateTime_(Idate1900)
  +'" ,PCI_DNATYPEPREAVIS=""');
  ExecuteSQLContOnExcept ('UPDATE CONTRATTRAVAIL SET PCI_DNAREFCONTRAT="",PCI_DNACLAUSECONC="-"');
  ExecuteSQLContOnExcept( 'UPDATE PAIEENCOURS SET PPU_PROFILRBS="",PPU_TYPPERSOURSSAF="",PPU_TYPPERSOURSSAM="",PPU_TYPEPERSOIRC="", PPU_COMM1="",PPU_COMM2="",PPU_COMM3="",PPU_COMMMOD="-"');
  ExecuteSQLContOnExcept( 'UPDATE DEPORTSAL SET PSE_TYPEPEXOMSA="ETB",PSE_PEXOMSA="", PSE_MSATYPLIEU="PER"');
  ExecuteSQLContOnExcept ('UPDATE ETABCOMPL SET ETB_PEXOMSA="",ETB_MSALIEUTRAV="",ETB_CONVENTION3=""');
// on revient en arière 255 champs max en access  ExecuteSQLContOnExcept ('UPDATE SALARIES SET PSA_DNALIENPARENTE="",PSA_DNACPTRANS="-",PSA_DNASALENTETR="",PSA_DNACATSPEC="100",PSA_TYPPRORATA="ETB"');
// on revient en arière 255 champs max en access  ExecuteSQLContOnExcept ('UPDATE SALARIES SET PSA_PRORATATVA = (SELECT ETB_PRORATATVA FROM ETABCOMPL WHERE ETB_ETABLISSEMENT=PSA_ETABLISSEMENT)');
//  ExecuteSQLContOnExcept ('UPDATE PROFILPAIE SET PPI_CONVENTION="000" WHERE PPI_CONVENTION="" OR PPI_CONVENTION IS NULL');

  //F BERGER 4232
//GP_20100517_DKZ_GP17440 Déb
//  ExecuteSQLNoPCL('UPDATE WNOMELIG SET WNL_DEPOT="", WNL_SITE="", WNL_TIERS="", WNL_SOCIETEGROUPE="" WHERE WNL_DEPOT IS NULL');
  {DKZ en commentaire pour regrouper les UPDATEs ... }
  //ExecuteSQLNoPCL( 'UPDATE WNOMELIG SET WNL_SITE="" WHERE WNL_SITE IS NULL' );
  //ExecuteSQLNoPCL( 'UPDATE WNOMELIG SET WNL_TIERS="" WHERE WNL_TIERS IS NULL' );
  //ExecuteSQLNoPCL( 'UPDATE WNOMELIG SET WNL_SOCIETEGROUPE="" WHERE WNL_SOCIETEGROUPE IS NULL' );
//GP_20100517_DKZ_GP17440 Fin

 //MVG 4233
  ExecuteSQLNoPCL('UPDATE IMMO SET I_SBVDATE1=I_SBVDATE, I_SBVMT1=I_SBVMT, I_SBVMTC1=I_SBVMTC, I_SBVREPRISE1=I_CORRECTIONVR, I_SBVCPT1=I_CPTSBVB, I_SBVREPDOT1=I_DPIEC, '+
  'I_SBVDATE2="' + UsDateTime_(iDate1900) + '", I_SBVDATE3="' + UsDateTime_(iDate1900) + '", I_SBVMT2=0, I_SBVMT3=0, I_SBVMTC2=0, I_SBVMTC3=0, I_SBVREPRISE2=0, I_SBVREPRISE3=0, I_SBVCPT2="", I_SBVCPT3="", '+
  'I_SBVREPDOT2="-", I_SBVREPDOT3="-", I_SBVREPCED2=0, I_SBVREPCED3=0');
  ExecuteSQLNoPCL ('UPDATE IMMO SET I_SBVREPCED1 = (SELECT IL_MONTANTAVMB FROM IMMOLOG WHERE (IL_IMMO=I_IMMO AND IL_TYPEOP="CES"))');
  ExecuteSQLNoPCL ('UPDATE IMMOLOG SET IL_TYPEOP="SB1" WHERE IL_TYPEOP="SBV"');
  ExecuteSQLNoPCL ('UPDATE IMMOLOG SET IL_TYPEOP="RS1" WHERE IL_TYPEOP="RSB"');

  //D MASSON 4240
  ExecuteSqlNoPcl('UPDATE DPAGRIDIV SET DAD_DOMINANTES = ""');

  //M DESGOUTTE 4249
  // Suppression de paramsoc obsolètes
  ExecuteSQLContOnExcept('DELETE FROM PARAMSOC WHERE SOC_NOM="SCO_SWSVERSIONCCN"');

  //J TRIFILIEFF 4254
//  if not isDossierPCL then InitNaturePceQteMini;

  // M GUERIN 4270
  {Initialisation du nouveau champ ajouté à la table}
  ExecuteSQLNoPCL('update PERSPHISTO set RPH_CODEIMPORT="" where ISNULL(RPH_CODEIMPORT,"")=""');

  //MC DESSEIGNET 4284
  ExecuteSqlContOnExcept
 ('update yliencomsx set ylo_majcpte="-", ylo_ETBINTEGRE="-", ylo_CTRLDOUBLONS="-", ylo_INTEGREEC="-", '
 +'ylo_CREATTIERSINT="-", ylo_BLANC="-", ylo_BOOLLIBRE1="-", ylo_BOOLLIBRE2="-", ylo_BOOLLIBRE3="-", '
 +'ylo_BOOLLIBRE4="-", ylo_BOOLLIBRE5="-",ylo_email="", ylo_JOURNAUX="", ylo_POUR="", ylo_TEXTELIBRE1="", '
 +'ylo_TEXTELIBRE2="", ylo_TEXTELIBRE3="", ylo_TEXTELIBRE4="", ylo_TEXTELIBRE5="" where ylo_email is null');

 //V GALLIOT 4295
  ExecuteSQLContOnExcept ('UPDATE PUBLICOTIS SET PUO_METHODALIMDADS = "VAL", PUO_DNATYPEREGUL = "99"');
  ExecuteSQLContOnExcept ('UPDATE ENVOISOCIAL SET PES_REFDECLREMP = 0');

 //MC DESSEIGNET 4296
  InsertChoixCode('ZLI', 'TI1', '.- PLA date libre 1', '.- ', '') ;
  InsertChoixCode('ZLI', 'TI2', '.- PLA date libre 2', '.- ', '') ;
  InsertChoixCode('ZLI', 'TI3', '.- PLA date libre 3', '.- ', '') ;
  InsertChoixCode('ZLI', 'TI4', '.- PLA date libre 4', '.- ', '') ;
//  InsertChoixCode('AET', 'PRE', 'Prévisionnel', 'PRE', 'AFF') ;
//  InsertChoixCode('TRE', 'FON', 'Fonctions', 'Fonctions', '') ;

  // Initialisation du paramsoc SO_AFFAIREMONOPIECE si il y a des affaire de type AFG
  (*
  if not isDossierPCL then
    if ExisteSQL('SELECT 1 FROM AFFAIRE WHERE AFF_STATUTAFFAIRE = "AFG"') then
      ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA = "X" WHERE SOC_NOM = "SO_AFFAIREMONOPIECE"');
  *)

 //R HARANG 4300
  { Modif GFO : Initialisation des champs systemes ajoutés dans les tables du domaine H.  }

 // 1
 (*
  ExecuteSqlNoPCL('UPDATE HRALLOTEMENT SET HAL_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HAL_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HAL_CREATEUR="" ,HAL_UTILISATEUR="",HAL_SOCIETE="",HAL_PREDEFINI=""'
      +' , HAL_NODOSSIER="" WHERE HAL_DATECREATION IS NULL');
 // 2
    ExecuteSqlNoPCL('UPDATE HRBADGES SET HBA_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HBA_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HBA_CREATEUR="" ,HBA_UTILISATEUR="",HBA_SOCIETE="",HBA_PREDEFINI=""'
      +' , HBA_NODOSSIER="" WHERE HBA_DATECREATION IS NULL');
 //3
    ExecuteSqlNoPCL('UPDATE HRCONTINGENT SET HCG_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HCG_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HCG_CREATEUR="" ,HCG_UTILISATEUR="",HCG_SOCIETE="",HCG_PREDEFINI=""'
      +' , HCG_NODOSSIER="" WHERE HCG_DATECREATION IS NULL');
  // 4
    ExecuteSqlNoPCL('UPDATE HRDOSSIER SET HDC_SOCIETE="",HDC_PREDEFINI=""'
      +' , HDC_NODOSSIER="" WHERE HDC_SOCIETE IS NULL');
  // 5
    ExecuteSqlNoPCL('UPDATE HRDOSRES SET HDR_SOCIETE="",HDR_PREDEFINI=""'
      +' , HDR_NODOSSIER="" WHERE HDR_SOCIETE IS NULL');
  // 6
    ExecuteSqlNoPCL('UPDATE HRFAMRES SET HFR_SOCIETE="",HFR_PREDEFINI=""'
      +' , HFR_NODOSSIER="" WHERE HFR_SOCIETE IS NULL');
  // 7
    ExecuteSqlNoPCL('UPDATE HRTYPRES SET HTR_SOCIETE="",HTR_PREDEFINI=""'
      +' , HTR_NODOSSIER="" WHERE HTR_SOCIETE IS NULL');
  // 8
    ExecuteSqlNoPCL('UPDATE HRETAT SET HES_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HES_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HES_CREATEUR="" ,HES_UTILISATEUR="",HES_SOCIETE="",HES_PREDEFINI=""'
      +' , HES_NODOSSIER="" WHERE HES_DATECREATION IS NULL');
  // 9
    ExecuteSqlNoPCL('UPDATE HREVENEMENT SET HEV_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HEV_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HEV_CREATEUR="" ,HEV_UTILISATEUR="",HEV_SOCIETE="",HEV_PREDEFINI=""'
      +' , HEV_NODOSSIER="" WHERE HEV_DATECREATION IS NULL');
  // 10
    ExecuteSqlNoPCL('UPDATE HRHISTDOSSIER SET HHD_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HHD_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HHD_CREATEUR="" ,HHD_UTILISATEUR="",HHD_SOCIETE="",HHD_PREDEFINI=""'
      +' , HHD_NODOSSIER="" WHERE HHD_DATECREATION IS NULL');
  // 11
    ExecuteSqlNoPCL('UPDATE HRMODERESA SET HMR_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HMR_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HMR_CREATEUR="" ,HMR_UTILISATEUR="",HMR_SOCIETE="",HMR_PREDEFINI=""'
      +' , HMR_NODOSSIER="" WHERE HMR_DATECREATION IS NULL');
  // 12
    ExecuteSqlNoPCL('UPDATE HRNBPERSONNE SET HNP_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HNP_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HNP_CREATEUR="" ,HNP_UTILISATEUR="",HNP_SOCIETE="",HNP_PREDEFINI=""'
      +' , HNP_NODOSSIER="" WHERE HNP_DATECREATION IS NULL');
  // 13
    ExecuteSqlNoPCL('UPDATE HRPARAMEMAIL SET HPE_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HPE_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HPE_CREATEUR="" ,HPE_UTILISATEUR="",HPE_SOCIETE="",HPE_PREDEFINI=""'
      +' , HPE_NODOSSIER="" WHERE HPE_DATECREATION IS NULL');
  // 14
    ExecuteSqlNoPCL('UPDATE HRPARAMPLANNING SET HPP_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HPP_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HPP_CREATEUR="" ,HPP_UTILISATEUR="",HPP_SOCIETE="",HPP_PREDEFINI=""'
      +' , HPP_NODOSSIER="" WHERE HPP_DATECREATION IS NULL');
  // 15
    ExecuteSqlNoPCL('UPDATE HRPARAMTYPEPREP SET HPT_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HPT_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HPT_CREATEUR="" ,HPT_UTILISATEUR="",HPT_SOCIETE="",HPT_PREDEFINI=""'
      +' , HPT_NODOSSIER="" WHERE HPT_DATECREATION IS NULL');
  // 16
    ExecuteSqlNoPCL('UPDATE HRPENSION SET HFP_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HFP_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HFP_CREATEUR="" ,HFP_UTILISATEUR="",HFP_SOCIETE="",HFP_PREDEFINI=""'
      +' , HFP_NODOSSIER="" WHERE HFP_DATECREATION IS NULL');
  // 17
    ExecuteSqlNoPCL('UPDATE HRPOINTPROD SET HPD_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HPD_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HPD_CREATEUR="" ,HPD_UTILISATEUR="",HPD_SOCIETE="",HPD_PREDEFINI=""'
      +' , HPD_NODOSSIER="" WHERE HPD_DATECREATION IS NULL');
  // 18
    ExecuteSqlNoPCL('UPDATE HRPOSTE SET HRP_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HRP_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HRP_CREATEUR="" ,HRP_UTILISATEUR="",HRP_SOCIETE="",HRP_PREDEFINI=""'
      +' , HRP_NODOSSIER="" WHERE HRP_DATECREATION IS NULL');
  // 19
    ExecuteSqlNoPCL('UPDATE HRPRIORITE SET HPR_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HPR_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HPR_CREATEUR="" ,HPR_UTILISATEUR="",HPR_SOCIETE="",HPR_PREDEFINI=""'
      +' , HPR_NODOSSIER="" WHERE HPR_DATECREATION IS NULL');
  // 20
    ExecuteSqlNoPCL('UPDATE HRREGROUPELIGNE SET HGP_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HGP_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HGP_CREATEUR="" ,HGP_UTILISATEUR="",HGP_SOCIETE="",HGP_PREDEFINI=""'
      +' , HGP_NODOSSIER="" WHERE HGP_DATECREATION IS NULL');
  // 21
    ExecuteSqlNoPCL('UPDATE HRRYTHMEGESTION SET HRG_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HRG_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HRG_CREATEUR="" ,HRG_UTILISATEUR="",HRG_SOCIETE="",HRG_PREDEFINI=""'
      +' , HRG_NODOSSIER="" WHERE HRG_DATECREATION IS NULL');
  // 22
    ExecuteSqlNoPCL('UPDATE HRTYPDOS SET HTD_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HTD_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HTD_CREATEUR="" ,HTD_UTILISATEUR="",HTD_SOCIETE="",HTD_PREDEFINI=""'
      +' , HTD_NODOSSIER="" WHERE HTD_DATECREATION IS NULL');
  // 23
    ExecuteSqlNoPCL('UPDATE HRTYPEPREPA SET HTP_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HTP_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HTP_CREATEUR="" ,HTP_UTILISATEUR="",HTP_SOCIETE="",HTP_PREDEFINI=""'
      +' , HTP_NODOSSIER="" WHERE HTP_DATECREATION IS NULL');
  // 24
    ExecuteSqlNoPCL('UPDATE HRPARAMPLANNING SET HPP_DATECREATION="'+UsDateTime_(iDate1900)+'"'
      +' , HPP_DATEMODIF="'+UsDateTime_(iDate1900)+'"'
      +' , HPP_CREATEUR="" ,HPP_UTILISATEUR="",HPP_SOCIETE="",HPP_PREDEFINI=""'
      +' , HPP_NODOSSIER="" WHERE HPP_DATECREATION IS NULL');
   { Modif GFO : Initialisation des nouveaux champs ajoutés dans les tables du domaine H.  }
   { Table HRPOSTE }
   ExecuteSqlNoPCL('UPDATE HRPOSTE SET HRP_TPEIDCAISSE="",HRP_TPEIDBTQ="" '
      +' WHERE HRP_TPEIDCAISSE IS NULL');
   { Table HRLIGNE }
   ExecuteSqlNoPCL('UPDATE HRLIGNE SET HRL_NUMTRANS=0,HRL_SOUSMENU=""'
      +' , HRL_TYPETRANS="",HRL_NUMMOD=0 WHERE HRL_NUMTRANS IS NULL');
   { Table HRCAISSE}
   ExecuteSqlNoPCL('UPDATE HRCAISSE SET HRC_PROFILMODEPAIE="",HRC_PARTAGE=""'
      +' , HRC_CATALOGUE="-",HRC_ARTICLECATA=""  WHERE HRC_PROFILMODEPAIE IS NULL');
   *)
    //MVG 4302
   ExecuteSQLNoPCL ('UPDATE IMMOAMOR SET IA_MONTANTSB2=0, IA_MONTANTSB3=0, IA_CESSIONSB2=0, IA_CESSIONSB3=0');

End;

Procedure MajVer943;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //N FOURNEL 4305
  if not isDossierPCL then
  begin
  (*
      if not ExisteSQL ('SELECT GPP_LIBELLE FROM PARPIECE WHERE GPP_NATUREPIECEG="AVA"') then
      begin
        ExecuteSqlContOnExcept ('INSERT INTO PARPIECE (GPP_NATUREPIECEG) VALUES("AVA")');
        ExecuteSqlContOnExcept ('UPDATE PARPIECE SET GPP_ACHATACTIVITE="-",'
                                     + ' GPP_ACOMPTE="-",'
                                     + ' GPP_ACTIONFINI="ENR",'
                                     + ' GPP_ACTIVITEPUPR="",'
                                     + ' GPP_AFAFFECTTB="",'
                                     + ' GPP_AFFPIECETABLE="-",'
                                     + ' GPP_APERCUAVETIQ="-",'
                                     + ' GPP_APERCUAVIMP="X",'
                                     + ' GPP_APPELPRIX="PUH",'
                                     + ' GPP_APPLICRG="-",'
                                     + ' GPP_ARTFOURPRIN="-",'
                                     + ' GPP_ARTSTOCK="-",'
                                     + ' GPP_BLOBART="X",'
                                     + ' GPP_BLOBLIENART="",'
                                     + ' GPP_BLOBLIENTIERS="",'
                                     + ' GPP_BLOBTIERS="X",'
                                     + ' GPP_CALCRUPTURE="AUC",'
                                     + ' GPP_CFGART="-",'
                                     + ' GPP_CFGARTASSIST="",'
                                     + ' GPP_CHAINAGE="",'
                                     + ' GPP_CODEPIECEOBL1="-",'
                                     + ' GPP_CODEPIECEOBL2="-",'
                                     + ' GPP_CODEPIECEOBL3="-",'
                                     + ' GPP_CODPIECEDEF1="",'
                                     + ' GPP_CODPIECEDEF2="",'
                                     + ' GPP_CODPIECEDEF3="",'
                                     + ' GPP_COMMENTENT="",'
                                     + ' GPP_COMMENTPIED="",'
                                     + ' GPP_COMPANALLIGNE="DEM",'
                                     + ' GPP_COMPANALPIED="DEM",'
                                     + ' GPP_COMPSTOCKLIGNE="SAN",'
                                     + ' GPP_COMPSTOCKPIED="SAN",'
                                     + ' GPP_CONDITIONTARIF="-",'
                                     + ' GPP_CONTEXTES="GC;AFF;",'
                                     + ' GPP_CONTRECHART1="-",'
                                     + ' GPP_CONTRECHART2="-",'
                                     + ' GPP_CONTRECHART3="-",'
                                     + ' GPP_CONTREMARQUE="-",'
                                     + ' GPP_CONTREMREF="-",'
                                     + ' GPP_CONTROLEMARGE="AUC",'
                                     + ' GPP_CPTCENTRAL="-",'
                                     + ' GPP_CRC=1155701647,'
                                     + ' GPP_CTRLENCOURS="-",'
                                     + ' GPP_CUMULART1="-",'
                                     + ' GPP_CUMULART2="-",'
                                     + ' GPP_CUMULART3="-",'
                                     + ' GPP_CUMULCOM1="-",'
                                     + ' GPP_CUMULCOM2="-",'
                                     + ' GPP_CUMULCOM3="-",'
                                     + ' GPP_CUMULTIERS1="-",'
                                     + ' GPP_CUMULTIERS2="-",'
                                     + ' GPP_CUMULTIERS3="-",'
                                     + ' GPP_DATELIBART1="AUC",'
                                     + ' GPP_DATELIBART2="AUC",'
                                     + ' GPP_DATELIBART3="AUC",'
                                     + ' GPP_DATELIBCOM1="",'
                                     + ' GPP_DATELIBCOM2="",'
                                     + ' GPP_DATELIBCOM3="",'
                                     + ' GPP_DATELIBTIERS1="AUC",'
                                     + ' GPP_DATELIBTIERS2="AUC",'
                                     + ' GPP_DATELIBTIERS3="AUC",'
                                     + ' GPP_DIMSAISIE="TOU",'
                                     + ' GPP_DUPLICPIECE="",'
                                     + ' GPP_ECLATEAFFAIRE="-",'
                                     + ' GPP_ECLATEDOMAINE="-",'
                                     + ' GPP_EDITIONNOMEN="AUC",'
                                     + ' GPP_ENCOURS="-",'
                                     + ' GPP_EQUIPIECE="",'
                                     + ' GPP_ESTAVOIR="X",'
                                     + ' GPP_ETATAFFAIRE="",'
                                     + ' GPP_ETATETIQ="",'
                                     + ' GPP_FAR_FAE="",'
                                     + ' GPP_FILTREARTCH="",'
                                     + ' GPP_FILTREARTVAL="",'
                                     + ' GPP_FILTRECOMM="",'
                                     + ' GPP_FORCERUPTURE="-",'
                                     + ' GPP_GEREARTICLELIE="AUT",'
                                     + ' GPP_GEREECHEANCE="DEM",'
                                     + ' GPP_GESTIONGRATUIT="-",'
                                     + ' GPP_HISTORIQUE="X",'
                                     + ' GPP_IFL1="011",'
                                     + ' GPP_IFL2="015",'
                                     + ' GPP_IFL3="021",'
                                     + ' GPP_IFL4="",'
                                     + ' GPP_IFL5="",'
                                     + ' GPP_IFL6="",'
                                     + ' GPP_IFL7="",'
                                     + ' GPP_IFL8="",'
                                     + ' GPP_IMPAUTOBESCBN="-",'
                                     + ' GPP_IMPAUTOETATCBN="-",'
                                     + ' GPP_IMPBESOIN="-",'
                                     + ' GPP_IMPETAT="Y98",'
                                     + ' GPP_IMPETIQ="-",'
                                     + ' GPP_IMPIMMEDIATE="X",'
                                     + ' GPP_IMPMODELE="",'
                                     + ' GPP_INFOPCEPRECCHX="",'
                                     + ' GPP_INFOPIECEPREC="",'
                                     + ' GPP_INFOSCOMPL="X",'
                                     + ' GPP_INFOSCPLPIECE="-",'
                                     + ' GPP_INITQTE="001",'
                                     + ' GPP_INITQTECRE=1,'
                                     + ' GPP_INSERTLIG="X",'
                                     + ' GPP_JOURNALCPTA="VTE",'
                                     + ' GPP_LIBELLE="Avoir de factures d''acomptes",'
                                     + ' GPP_LIENAFFAIRE="-",'
                                     + ' GPP_LIENTACHE="-",'
                                     + ' GPP_LISTEAFFAIRE="AFSAISIEFAC",'
                                     + ' GPP_LISTESAISIE="HLSAISIEH",'
                                     + ' GPP_LOT="-",'
                                     + ' GPP_MAJINFOTIERS="X",'
                                     + ' GPP_MAJPRIXVALO="",'
                                     + ' GPP_MASQUERNATURE="-",'
                                     + ' GPP_MENU="",'
                                     + ' GPP_MESSAGEEDIIN="",'
                                     + ' GPP_MESSAGEEDIOUT="",'
                                     + ' GPP_MODBESOIN="",'
                                     + ' GPP_MODEECHEANCES="SR",'
                                     + ' GPP_MODEGROUPEPORT="CHA",'
                                     + ' GPP_MODELEWORD="",'
                                     + ' GPP_MODIFCOUT="-",'
                                     + ' GPP_MODPLANIFIABLE="NON",'
                                     + ' GPP_MONTANTMINI=0,'
                                     + ' GPP_MONTANTVISA=0,'
                                     + ' GPP_MULTIGRILLE="-",'
                                     + ' GPP_NATPIECEANNUL="",'
                                     + ' GPP_NATURECPTA="AC",'
                                     + ' GPP_NATUREORIGINE="",'
                                     + ' GPP_NATUREPIECEG="AVA",'
                                     + ' GPP_NATUREREPRISE="",'
                                     + ' GPP_NATURESUIVANTE="",'
                                     + ' GPP_NATURETIERS="CLI;",'
                                     + ' GPP_NBEXEMPLAIRE=1,'
                                     + ' GPP_NIVEAUPARAM="EDI",'
                                     + ' GPP_NUMEROSERIE="-",'
                                     + ' GPP_OBJETDIM="-",'
                                     + ' GPP_OBLIGEREGLE="-",'
                                     + ' GPP_OUVREAUTOPORT="-",'
                                     + ' GPP_PARAMDIM="-",'
                                     + ' GPP_PARAMGRILLEDIM="-",'
                                     + ' GPP_PIECEEDI="-",'
                                     + ' GPP_PIECEPILOTE="-",'
                                     + ' GPP_PIECESAV="-",'
                                     + ' GPP_PIECETABLE1="",'
                                     + ' GPP_PIECETABLE2="",'
                                     + ' GPP_PIECETABLE3="",'
                                     + ' GPP_PILOTEORDRE="-",'
                                     + ' GPP_PLAETATCREAT="",'
                                     + ' GPP_PLAETATSOLDE="",'
                                     + ' GPP_PLANNINGFLUX="",'
                                     + ' GPP_PREVUAFFAIRE="-",'
                                     + ' GPP_PRIORECHART1="ART",'
                                     + ' GPP_PRIORECHART2="REF",'
                                     + ' GPP_PRIORECHART3="BAR",'
                                     + ' GPP_PRIXNULOK="X",'
                                     + ' GPP_PROCLI="-",'
                                     + ' GPP_QTEMOINS="",'
                                     + ' GPP_QTEPLUS="",'
                                     + ' GPP_QUALIFMVT="",'
                                     + ' GPP_RACINELIBECR1="",'
                                     + ' GPP_RACINELIBECR2="",'
                                     + ' GPP_RACINEREFINT1="",'
                                     + ' GPP_RACINEREFINT2="",'
                                     + ' GPP_RECALCULPRIX="-",'
                                     + ' GPP_RECHTARIF501="-",'
                                     + ' GPP_RECUPPRE="PRE",'
                                     + ' GPP_REFEXTCTRL="000",'
                                     + ' GPP_REFINTCTRL="000",'
                                     + ' GPP_REFINTEXT="INT",'
                                     + ' GPP_REGROUPCPTA="AUC",'
                                     + ' GPP_REGROUPE="X",'
                                     + ' GPP_RELIQUAT="-",'
                                     + ' GPP_REPRISEENTAFF="",'
                                     + ' GPP_REPRISELIGAFF="",'
                                     + ' GPP_SENSPIECE="",'
                                     + ' GPP_SOLDETRANSFO="-",'
                                     + ' GPP_SOUCHE="GAA",'
                                     + ' GPP_STKQUALIFMVT="",'
                                     + ' GPP_TARIFGENDATE="010",'
                                     + ' GPP_TARIFGENDEPOT="-",'
                                     + ' GPP_TARIFGENPNPB="-",'
                                     + ' GPP_TARIFGENSAISIE="010",'
                                     + ' GPP_TARIFGENSPECIA="010",'
                                     + ' GPP_TARIFGENTRANSF="010",'
                                     + ' GPP_TARIFMODULE="201",'
                                     + ' GPP_TAXE="X",'
                                     + ' GPP_TIERS="",'
                                     + ' GPP_TRSFACHAT="-",'
                                     + ' GPP_TRSFVENTE="-",'
                                     + ' GPP_TYPEACTIVITE="",'
                                     + ' GPP_TYPEARTICLE="CNS;FI;FRA;MAR;NOM;PRE;",'
                                     + ' GPP_TYPECOMMERCIAL="REP",'
                                     + ' GPP_TYPEDIMOBLI1="-",'
                                     + ' GPP_TYPEDIMOBLI2="-",'
                                     + ' GPP_TYPEDIMOBLI3="-",'
                                     + ' GPP_TYPEDIMOBLI4="-",'
                                     + ' GPP_TYPEDIMOBLI5="-",'
                                     + ' GPP_TYPEECRALIM="",'
                                     + ' GPP_TYPEECRCPTA="NOR",'
                                     + ' GPP_TYPEECRSTOCK="RIE",'
                                     + ' GPP_TYPEFACT="",'
                                     + ' GPP_TYPEPASSACC="REE",'
                                     + ' GPP_TYPEPASSACCR="REE",'
                                     + ' GPP_TYPEPASSCPTA="REE",'
                                     + ' GPP_TYPEPASSCPTAR="REE",'
                                     + ' GPP_TYPEPRESDOC="DEF",'
                                     + ' GPP_TYPEPRESENT=0,'
                                     + ' GPP_VALEURLIBECR1="",'
                                     + ' GPP_VALEURLIBECR2="",'
                                     + ' GPP_VALEURREFINT1="",'
                                     + ' GPP_VALEURREFINT2="",'
                                     + ' GPP_VALLIBART1="AUC",'
                                     + ' GPP_VALLIBART2="AUC",'
                                     + ' GPP_VALLIBART3="AUC",'
                                     + ' GPP_VALLIBCOM1="",'
                                     + ' GPP_VALLIBCOM2="",'
                                     + ' GPP_VALLIBCOM3="",'
                                     + ' GPP_VALLIBTIERS1="AUC",'
                                     + ' GPP_VALLIBTIERS2="AUC",'
                                     + ' GPP_VALLIBTIERS3="AUC",'
                                     + ' GPP_VALMODELE="X",'
                                     + ' GPP_VENTEACHAT="VEN",'
                                     + ' GPP_VISA="-"'
                                     + ' WHERE GPP_NATUREPIECEG = "AVA"');
      end;
      *)
  End;
    //C DUMAS 4330
    ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Montant frais création" WHERE dh_prefixe = "ROP" and DH_NOMCHAMP = "ROP_FRAISCREATION"');

    ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Nombre d''évaluations enregistrées" WHERE dh_prefixe = "RQT" and DH_NOMCHAMP = "RQT_NBREVAL"');

    ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Libellé du critère d''évaluation" WHERE DH_PREFIXE = "RQC" AND DH_NOMCHAMP = "RQC_LIBELLE"');

    ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Numéro contact partenaire" WHERE dh_prefixe = "RPN" and DH_NOMCHAMP = "RPN_NUMEROCONTACT"');

    ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Type d''écart  / défaut principal" WHERE dh_prefixe = "RQN" and DH_NOMCHAMP = "RQN_TYPEECARTNIV1"');

    ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Type d''écart niveau 2" WHERE dh_prefixe = "RQN" and DH_NOMCHAMP = "RQN_TYPEECARTNIV2"');

    // M FAUDEL 4339
    // Initialisations des nouveaux champs de la table SALARIESTYPE
    ExecuteSQLContOnExcept ('UPDATE SALARIESTYPE SET PSK_LIBRE2="",PSK_LIBRE3="",PSK_LIBRE4="",PSK_LIBRE5=""'
    +', PSK_LIBRE6="", PSK_DNACATSPEC="", PSK_TYPEPERSOIRC="",PSK_TYPPROFILANC="",PSK_PROFILANCIEN=""'
    +',PSK_ANCIENNETE="",PSK_ETATBULLETIN=""');

    //M FAUDEL 4340
    // Initialisation des nouveaux champs de la table ORGANISMEPAIE
    ExecuteSQLContOnExcept ('UPDATE ORGANISMEPAIE SET POG_MUTUELLE="-", POG_NOMUTUELLE=""');

    //S BOUSSERT 4346
    // Initialisation du champ EMP_EMPCATEGORIE
    ExecuteSQLContOnExcept('UPDATE FEMPRUNT SET EMP_EMPCATEGORIE="EXP", EMP_EMPFONCIER="-"');

    //MC DESSEIGNET 4348
//    ExecuteSQLNoPCL( 'UPDATE TACHE SET ATA_TYPELIGNEPLA = "" WHERE ATA_TYPELIGNEPLA IS NULL');

    if not isDossierPCL then
    begin
      (*
      if not ExisteSQL ('SELECT GPP_LIBELLE FROM PARPIECE WHERE GPP_NATUREPIECEG="FFC"') then
        begin
           ExecuteSQLContOnExcept ('INSERT INTO PARPIECE (GPP_NATUREPIECEG) VALUES("FFC")');

           ExecuteSQLContOnExcept  ('UPDATE PARPIECE SET GPP_ACHATACTIVITE="-",'
                        + ' GPP_ACOMPTE="X",'
                        + ' GPP_ACTIONFINI="COM",'
                        + ' GPP_ACTIVITEPUPR="",'
                        + ' GPP_AFAFFECTTB="",'
                        + ' GPP_AFFPIECETABLE="-",'
                        + ' GPP_APERCUAVETIQ="-",'
                        + ' GPP_APERCUAVIMP="-",'
                        + ' GPP_APPELPRIX="PUH",'
                        + ' GPP_APPLICRG="-",'
                        + ' GPP_ARTFOURPRIN="-",'
                        + ' GPP_ARTSTOCK="-",'
                        + ' GPP_BLOBART="X",'
                        + ' GPP_BLOBLIENART="",'
                        + ' GPP_BLOBLIENTIERS="",'
                        + ' GPP_BLOBTIERS="X",'
                        + ' GPP_CALCRUPTURE="AUC",'
                        + ' GPP_CFGART="-",'
                        + ' GPP_CFGARTASSIST="",'
                        + ' GPP_CHAINAGE="",'
                        + ' GPP_CODEPIECEOBL1="-",'
                        + ' GPP_CODEPIECEOBL2="-",'
                        + ' GPP_CODEPIECEOBL3="-",'
                        + ' GPP_CODPIECEDEF1="",'
                        + ' GPP_CODPIECEDEF2="",'
                        + ' GPP_CODPIECEDEF3="",'
                        + ' GPP_COMMENTENT="",'
                        + ' GPP_COMMENTPIED="",'
                        + ' GPP_COMPANALLIGNE="DEM",'
                        + ' GPP_COMPANALPIED="DEM",'
                        + ' GPP_COMPSTOCKLIGNE="",'
                        + ' GPP_COMPSTOCKPIED="",'
                        + ' GPP_CONDITIONTARIF="X",'
                        + ' GPP_CONTEXTES="GC;",'
                        + ' GPP_CONTRECHART1="-",'
                        + ' GPP_CONTRECHART2="-",'
                        + ' GPP_CONTRECHART3="-",'
                        + ' GPP_CONTREMARQUE="-",'
                        + ' GPP_CONTREMREF="-",'
                        + ' GPP_CONTROLEMARGE="AUC",'
                        + ' GPP_CPTCENTRAL="-",'
                        + ' GPP_CRC=0,'
                        + ' GPP_CTRLENCOURS="X",'
                        + ' GPP_CUMULART1="-",'
                        + ' GPP_CUMULART2="-",'
                        + ' GPP_CUMULART3="-",'
                        + ' GPP_CUMULCOM1="-",'
                        + ' GPP_CUMULCOM2="-",'
                        + ' GPP_CUMULCOM3="-",'
                        + ' GPP_CUMULTIERS1="-",'
                        + ' GPP_CUMULTIERS2="-",'
                        + ' GPP_CUMULTIERS3="-",'
                        + ' GPP_DATELIBART1="",'
                        + ' GPP_DATELIBART2="",'
                        + ' GPP_DATELIBART3="",'
                        + ' GPP_DATELIBCOM1="",'
                        + ' GPP_DATELIBCOM2="",'
                        + ' GPP_DATELIBCOM3="",'
                        + ' GPP_DATELIBTIERS1="",'
                        + ' GPP_DATELIBTIERS2="",'
                        + ' GPP_DATELIBTIERS3="",'
                        + ' GPP_DIMSAISIE="",'
                        + ' GPP_DUPLICPIECE="",'
                        + ' GPP_ECLATEAFFAIRE="-",'
                        + ' GPP_ECLATEDOMAINE="-",'
                        + ' GPP_EDITIONNOMEN="",'
                        + ' GPP_ENCOURS="X",'
                        + ' GPP_EQUIPIECE="",'
                        + ' GPP_ESTAVOIR="-",'
                        + ' GPP_ETATAFFAIRE="",'
                        + ' GPP_ETATETIQ="",'
                        + ' GPP_FAR_FAE="",'
                        + ' GPP_FILTREARTCH="",'
                        + ' GPP_FILTREARTVAL="",'
                        + ' GPP_FILTRECOMM="",'
                        + ' GPP_FORCERUPTURE="-",'
                        + ' GPP_GEREARTICLELIE="AUT",'
                        + ' GPP_GEREECHEANCE="DEM",'
                        + ' GPP_GESTIONGRATUIT="-",'
                        + ' GPP_HISTORIQUE="X",'
                        + ' GPP_IFL1="",'
                        + ' GPP_IFL2="",'
                        + ' GPP_IFL3="",'
                        + ' GPP_IFL4="",'
                        + ' GPP_IFL5="",'
                        + ' GPP_IFL6="",'
                        + ' GPP_IFL7="",'
                        + ' GPP_IFL8="",'
                        + ' GPP_IMPAUTOBESCBN="-",'
                        + ' GPP_IMPAUTOETATCBN="-",'
                        + ' GPP_IMPBESOIN="-",'
                        + ' GPP_IMPETAT="",'
                        + ' GPP_IMPETIQ="-",'
                        + ' GPP_IMPIMMEDIATE="-",'
                        + ' GPP_IMPMODELE="",'
                        + ' GPP_INFOPCEPRECCHX="",'
                        + ' GPP_INFOPIECEPREC="",'
                        + ' GPP_INFOSCOMPL="-",'
                        + ' GPP_INFOSCPLPIECE="-",'
                        + ' GPP_INITQTE="001",'
                        + ' GPP_INITQTECRE=0,'
                        + ' GPP_INSERTLIG="X",'
                        + ' GPP_JOURNALCPTA="VED",'
                        + ' GPP_LIBELLE="Facture client fincancière",'
                        + ' GPP_LIENAFFAIRE="-",'
                        + ' GPP_LIENTACHE="-",'
                        + ' GPP_LISTEAFFAIRE="",'
                        + ' GPP_LISTESAISIE="GCSAISIEFAC",'
                        + ' GPP_LOT="X",'
                        + ' GPP_MAJINFOTIERS="X",'
                        + ' GPP_MAJPRIXVALO="",'
                        + ' GPP_MASQUERNATURE="-",'
                        + ' GPP_MENU="",'
                        + ' GPP_MESSAGEEDIIN="",'
                        + ' GPP_MESSAGEEDIOUT="",'
                        + ' GPP_MODBESOIN="",'
                        + ' GPP_MODEECHEANCES="RS",'
                        + ' GPP_MODEGROUPEPORT="CHA",'
                        + ' GPP_MODELEWORD="",'
                        + ' GPP_MODIFCOUT="-",'
                        + ' GPP_MODPLANIFIABLE="",'
                        + ' GPP_MONTANTMINI=0,'
                        + ' GPP_MONTANTVISA=0,'
                        + ' GPP_MULTIGRILLE="-",'
                        + ' GPP_NATPIECEANNUL="",'
                        + ' GPP_NATURECPTA="FC",'
                        + ' GPP_NATUREORIGINE="",'
                        + ' GPP_NATUREPIECEG="FFC",'
                        + ' GPP_NATUREREPRISE="FCF;",'
                        + ' GPP_NATURESUIVANTE="",'
                        + ' GPP_NATURETIERS="CLI;",'
                        + ' GPP_NBEXEMPLAIRE=1,'
                        + ' GPP_NIVEAUPARAM="EDI",'
                        + ' GPP_NUMEROSERIE="X",'
                        + ' GPP_OBJETDIM="-",'
                        + ' GPP_OBLIGEREGLE="-",'
                        + ' GPP_OUVREAUTOPORT="-",'
                        + ' GPP_PARAMDIM="-",'
                        + ' GPP_PARAMGRILLEDIM="-",'
                        + ' GPP_PIECEEDI="-",'
                        + ' GPP_PIECEPILOTE="-",'
                        + ' GPP_PIECESAV="-",'
                        + ' GPP_PIECETABLE1="",'
                        + ' GPP_PIECETABLE2="",'
                        + ' GPP_PIECETABLE3="",'
                        + ' GPP_PILOTEORDRE="-",'
                        + ' GPP_PLAETATCREAT="",'
                        + ' GPP_PLAETATSOLDE="",'
                        + ' GPP_PLANNINGFLUX="",'
                        + ' GPP_PREVUAFFAIRE="-",'
                        + ' GPP_PRIORECHART1="ART",'
                        + ' GPP_PRIORECHART2="REF",'
                        + ' GPP_PRIORECHART3="BAR",'
                        + ' GPP_PRIXNULOK="X",'
                        + ' GPP_PROCLI="-",'
                        + ' GPP_QTEMOINS="",'
                        + ' GPP_QTEPLUS="",'
                        + ' GPP_QUALIFMVT="",'
                        + ' GPP_RACINELIBECR1="",'
                        + ' GPP_RACINELIBECR2="",'
                        + ' GPP_RACINEREFINT1="",'
                        + ' GPP_RACINEREFINT2="",'
                        + ' GPP_RECALCULPRIX="X",'
                        + ' GPP_RECHTARIF501="-",'
                        + ' GPP_RECUPPRE="PRE",'
                        + ' GPP_REFEXTCTRL="000",'
                        + ' GPP_REFINTCTRL="000",'
                        + ' GPP_REFINTEXT="INT",'
                        + ' GPP_REGROUPCPTA="",'
                        + ' GPP_REGROUPE="X",'
                        + ' GPP_RELIQUAT="X",'
                        + ' GPP_REPRISEENTAFF="",'
                        + ' GPP_REPRISELIGAFF="",'
                        + ' GPP_SENSPIECE="MIX",'
                        + ' GPP_SOLDETRANSFO="-",'
                        + ' GPP_SOUCHE="GFC",'
                        + ' GPP_STKQUALIFMVT="",'
                        + ' GPP_TARIFGENDATE="010",'
                        + ' GPP_TARIFGENDEPOT="-",'
                        + ' GPP_TARIFGENPNPB="-",'
                        + ' GPP_TARIFGENSAISIE="010",'
                        + ' GPP_TARIFGENSPECIA="010",'
                        + ' GPP_TARIFGENTRANSF="010",'
                        + ' GPP_TARIFMODULE="201",'
                        + ' GPP_TAXE="X",'
                        + ' GPP_TIERS="",'
                        + ' GPP_TRSFACHAT="-",'
                        + ' GPP_TRSFVENTE="-",'
                        + ' GPP_TYPEACTIVITE="",'
                        + ' GPP_TYPEARTICLE="CNS;FRA;MAR;NOM;PRE;",'
                        + ' GPP_TYPECOMMERCIAL="REP",'
                        + ' GPP_TYPEDIMOBLI1="-",'
                        + ' GPP_TYPEDIMOBLI2="-",'
                        + ' GPP_TYPEDIMOBLI3="-",'
                        + ' GPP_TYPEDIMOBLI4="-",'
                        + ' GPP_TYPEDIMOBLI5="-",'
                        + ' GPP_TYPEECRALIM="",'
                        + ' GPP_TYPEECRCPTA="NOR",'
                        + ' GPP_TYPEECRSTOCK="",'
                        + ' GPP_TYPEFACT="",'
                        + ' GPP_TYPEPASSACC="REE",'
                        + ' GPP_TYPEPASSACCR="REE",'
                        + ' GPP_TYPEPASSCPTA="ZDI",'
                        + ' GPP_TYPEPASSCPTAR="ZDI",'
                        + ' GPP_TYPEPRESDOC="",'
                        + ' GPP_TYPEPRESENT=0,'
                        + ' GPP_VALEURLIBECR1="",'
                        + ' GPP_VALEURLIBECR2="",'
                        + ' GPP_VALEURREFINT1="",'
                        + ' GPP_VALEURREFINT2="",'
                        + ' GPP_VALLIBART1="",'
                        + ' GPP_VALLIBART2="",'
                        + ' GPP_VALLIBART3="",'
                        + ' GPP_VALLIBCOM1="",'
                        + ' GPP_VALLIBCOM2="",'
                        + ' GPP_VALLIBCOM3="",'
                        + ' GPP_VALLIBTIERS1="",'
                        + ' GPP_VALLIBTIERS2="",'
                        + ' GPP_VALLIBTIERS3="",'
                        + ' GPP_VALMODELE="X",'
                        + ' GPP_VENTEACHAT="VEN",'
                        + ' GPP_VISA="-"'
        + ' WHERE GPP_NATUREPIECEG="FFC"')
    end;
    *)
  end;

End;

procedure MajVer944;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //TS GP_20090302_TS_TD4759
//  if not IsDossierPCL() then
//    EdiMajVer(944).MajApres();

  //F BERGER 4365
//  ExecuteSQLNoPCL( 'UPDATE WPARAMFONCTION SET WPF_LISTESTD="" WHERE WPF_LISTESTD IS NULL' );

  //M FAUDEL 4368
  // Table  PROFILSPECIAUX : Initialisation du champ PPS_SALARIETYPE
  ExecuteSQLContOnExcept('UPDATE PROFILSPECIAUX SET PPS_SALARIETYPE=""');

  //MNG 4378
  // MAJ dh_controle SAV pour maj en série
//  ExecuteSQLContOnExcept( 'UPDATE DECHAMPS SET DH_CONTROLE = "LDC" WHERE DH_PREFIXE = "WAP" AND DH_NUMCHAMP IN (10,190,200,210,220)' );
//  ExecuteSQLContOnExcept( 'UPDATE DECHAMPS SET DH_CONTROLE = "LD" WHERE DH_PREFIXE = "WAP" AND DH_NUMCHAMP IN (40,990,1000)' );

  //M MORRETTON 4392
  ExecuteSQLNoPCL('DELETE FROM PARAMSOC WHERE SOC_NOM IN ("SCO_PDRMETHVALO","SCO_PDRGENEREWPL","SCO_PDRDEFAUTGA","SCO_PDRCONSOAVANCE")');
//GP_20100517_DKZ_GP17440
  //ExecuteSQLNoPCL('DELETE FROM PARAMSOC WHERE SOC_NOM IN ("SCO_PDRMETHVALO","SCO_PDRGENEREWPL","SCO_PDRDEFAUTGA","SCO_PDRCONSOAVANCE")');

{  //MCD 4410
  // pour Pme uniquement (car fait en pcl par script en exe820)

  if V_PGI.ModePCL <> '1' then
  begin
  ExecuteSQLNoPCL
  ('update menu set mn_tag =0 where mn_1=26 and mn_2=4 and mn_3=8 and mn_4=0');
  ExecuteSQLNoPCL
  ('update menu set mn_accesgrp = (select mn_accesgrp from menu where mn_1=26 and mn_2=4 and mn_3=8 and mn_4=3) where mn_1=26 and mn_2=4 and mn_3=8');
  end;}

  //CORRECTION VIOLATION D'ACCES (le SELECT renvoyait NULL et un MN_ACCESGRP a NULL provoque une violation d'accès)
  //suite modif menu 26;4, ce qui est fait en MjaVER944
  //MCD 4410
  // pour Pme uniquement (car fait en pcl par script en exe820)
  // fait sur le tag au lieu place dans le menu pour eviter les pb si modif ordre menu

  if V_PGI.ModePCL <> '1' then
  begin
  ExecuteSQLNoPCL('update menu set mn_accesgrp = (select mn_accesgrp from menu where mn_tag=26032 ) where mn_tag=26168 or mn_tag=26169');
  end;

  //M BOUDIN 4388
  ExecuteSQLNoPCL ('UPDATE IMMO SET I_OPEBAT="-"');

  //T PETETIN 5042
  // TP 944-4389 WFORMCONVSAVE V5
//  ExecuteSQLNOPCL('UPDATE WFORMCONVSAVE SET WWS_ECARTUNI01=0,WWS_ECARTUNI02=0,WWS_ECARTUNI03=0,WWS_ECARTUNI04=0,WWS_ECARTUNI05=0,WWS_ECARTUNI06=0,WWS_ECARTUNI07=0,WWS_ECARTUNI08=0,WWS_ECARTUNI09=0');
  // TP 944-4389 WORDRELIG V23
//  ExecuteSQLNOPCL('UPDATE WORDRELIG SET WOL_ECARTUNIDEM=0,WOL_ECARTUNIACC=0,WOL_ECARTUNILAN=0,WOL_ECARTUNIREC=0,WOL_ECARTUNIREB=0');
  // TP 944-4389 WORDREPHASE V11
//  ExecuteSQLNOPCL('UPDATE WORDREPHASE SET WOP_ECARTUNIACC=0,WOP_ECARTUNILAN=0,WOP_ECARTUNIARE=0,WOP_ECARTUNIREC=0,WOP_ECARTUNISUS=0,WOP_ECARTUNIREB=0');
  // TP 944-4389 WORDREBES V15
//  ExecuteSQLNOPCL('UPDATE WORDREBES SET WOB_ECARTUNILIEN=0,WOB_ECARTUNIPERI=0,WOB_ECARTUNIPPER=0,WOB_ECARTUNIPFIXE=0,WOB_ECARTUNIBES=0,WOB_ECARTUNIAFF=0,WOB_ECARTUNIRUP=0,WOB_ECARTUNICON=0,WOB_ECARTUNITST=0,WOB_ECARTUNILAS=0');


end;

procedure MajVer945;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //MM 4422
  ExecuteSQLNoPCL('DELETE FROM PARAMSOC WHERE SOC_NOM IN ("SCO_GBORDRE","SCO_GBORDRECDE")');

  if not IsDossierPCL then
  begin
    //Ajout de champs obligatoires dans BusinessSide FQ10297

    If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "EC3" AND CO_CODE = "EA1"') then
      ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) values("EC3", "EA1", "Raison sociale", "EAG_LIBELLE", "")');

    If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "EC3" AND CO_CODE = "EA2"') then
      ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) values("EC3", "EA2", "Code Postal", "EAG_CODEPOSTAL", "")');

    If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "EC3" AND CO_CODE = "EA3"') then
      ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) values("EC3", "EA3", "Ville", "EAG_VILLE", "")');

    If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "EC3" AND CO_CODE = "EA5"') then
      ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) values("EC3", "EA5", "N° interne de l''adresse", "EAG_NUMEROADRESSE", "")');

    If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "EC3" AND CO_CODE = "PR1"') then
      ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) values("EC3", "PR1", "Tiers", "EPR_AUXILIAIRE", "")');

    If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "EC3" AND CO_CODE = "LT1"') then
      ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) values("EC3", "LT1", "Nature", "ELT_NATUREPIECEG", "")');

    If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "EC3" AND CO_CODE = "LT2"') then
      ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) values("EC3", "LT2", "Souche", "ELT_SOUCHE", "")');

    If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "EC3" AND CO_CODE = "LT3"') then
      ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) values("EC3", "LT3", "Numéro", "ELT_NUMERO", "")');

    If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "EC3" AND CO_CODE = "LT4"') then
      ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) values("EC3", "LT4", "Numéro de ligne", "ELT_NUMLIGNE", "")');

    If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "EC3" AND CO_CODE = "LT5"') then
      ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) values("EC3", "LT5", "Indice", "ELT_CLEDATA", "")');

    If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "EC3" AND CO_CODE = "R41"') then
      ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) values("EC3", "R41", "Clé", "ER4_CLEDATA", "")');

    If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "EC3" AND CO_CODE = "R21"') then
      ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) values("EC3", "R21", "Clé", "ER2_CLEDATA", "")');

    If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "EC3" AND CO_CODE = "EE1"') then
      ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) values("EC3", "EE1", "Clé", "EER_CLEDATA", "")');
  End;
(*
  //C DUMAS 4451
    ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Libellé partenaire" WHERE dh_prefixe = "RPN" and DH_NOMCHAMP = "RPN_LIBELLE"');
    ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Pièce origine (Critères délai)" WHERE dh_prefixe = "RQE" and DH_NOMCHAMP = "RQE_REFORIGINE"');
    ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Ordre production" WHERE dh_prefixe = "RAC" and DH_NOMCHAMP = "RAC_LIGNEORDRE"');
    ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Ordre production" WHERE dh_prefixe = "WPE" and DH_NOMCHAMP = "WPE_LIGNEORDRE"');
    ExecuteSQLContOnExcept('UPDATE DETABLES SET DT_LIBELLE = "Partenaires opérations marketing" WHERE dt_prefixe = "RPN" and DT_NOMTABLE = "PARTENAIRES"');
*)
  // D SCLAVOPOULOS 4461
//  AGLNettoieListes('AFMULFACTIERSAFF' ,'AFF_TYPEAFFAIRE', Nil);

  //G JUGDE 4483
//  ExecuteSQLNoPCL('UPDATE QPHASE SET QPH_PRENUMOPER=0,QPH_DERNUMOPER=0 WHERE QPH_PRENUMOPER IS NULL');

  //M GUERIN 4484
  ExecuteSqlContOnExcept('update dechamps set dh_libelle = "Date des cumuls non renseignée" where dh_prefixe = "EMG" '
  +'and DH_NUMCHAMP = 660 and dh_libelle = "Date des cumul non renseignée"')

End;

procedure MajVer946;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  { JTR - Calcul du champ GPB_NUMTAXE de la table PIEDBASE }
//  if not isDossierPCL then  MajGPB_NUMTAXE;

  // S MASSON 4485
  ExecuteSqlNoPcl('UPDATE DPAGRIDIV SET DAD_DOMSECOND = ""');

  //M GUERIN 4488
  {correction fautes d'orthographes sur libellés}
//GP_20100517_DKZ_GP17440
(*
  ExecuteSqlContOnExcept('update decombos set do_libelle = "Etat coût standard" where do_combo = "GCCOUTSTDWPF"'
  +' and do_libelle = "Etat coût standard"');
//GP_20100517_DKZ_GP17440
  ExecuteSqlContOnExcept('update decombos set do_libelle = "Champs à récupérer de l''entête de" where'
  +' do_combo = "GCRECUPENTETEFAA" and do_libelle = "Champs à récupérer dans l''entête de"');
  ExecuteSqlContOnExcept('update DECHAMPS set DH_LIBELLE = "Prix revient : Valorisation/Qté éco"'
  +' where DH_PREFIXE = "GA" and DH_NUMCHAMP = 1660 and DH_LIBELLE = "Prix evient : Valorisation/Qté éco"');
//GP_20100517_DKZ_GP17440
  ExecuteSqlContOnExcept('update DECHAMPS set DH_LIBELLE = "Contrôle de capacité" where'
  +' DH_PREFIXE = "GEM" and DH_NUMCHAMP = 220 and DH_LIBELLE = "Contrôle de capacité"');
*)
  //MCD 4494
//  AglNettoieListesPlus('AFMULRECHAFFAIRE', 'AFF_TYPEAFFAIRE', nil, true);

  //JTR 4497
  { JTR - Initialisation à vide du champs GPT_VENTCOMPTA }
//  ExecuteSQLNoPCL('UPDATE PIEDPORT SET GPT_VENTCOMPTA = ""');

  //TP 4512
  ExecuteSQLNoPCL('UPDATE CHOIXCOD SET CC_LIBRE="CLI" WHERE CC_TYPE="RQO" AND CC_CODE="CL"');
  ExecuteSQLNoPCL('UPDATE CHOIXCOD SET CC_LIBRE="FOU" WHERE CC_TYPE="RQO" AND CC_CODE="FO"');
  ExecuteSQLNoPCL('UPDATE CHOIXCOD SET CC_LIBRE="SAL" WHERE CC_TYPE="RQO" AND CC_CODE="IN"');

  //TS 4517
    { TD;4759 }
//  if not IsDossierPCL() then
//    EdiMajVer(946).MajApres();

  //D KOZA 4526
//  ExecuteSqlNoPCL('UPDATE QQTMINCAR SET QQM_SEUILCTRTMAT=0,QQM_SEUILNBJART=0 WHERE QQM_SEUILCTRTMAT IS NULL');

  //T PETETIN 4527
  { 4527 - WPARAMFONCTION de V23 en V24 }
//  ExecuteSQLNOPCL('UPDATE WPARAMFONCTION SET WPF_UNITEDEPART="",WPF_UNITEARRIVEE="",WPF_FORMULEVAR=""');

End;

procedure MajVer947;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //C DUMAS 4500
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Code opération" WHERE dh_prefixe = "RPN" and DH_NOMCHAMP = "RPN_OPERATION"');
  //MNG 4529
  // ajout d'un champ origine ciblage
//  ExecuteSqlNoPCL  ('UPDATE CIBLAGE SET RCB_ORIGINECIBLAGE = ""');
  ExecuteSqlNoPCL  ('UPDATE CIBLAGE SET RCB_ORIGINECIBLAGE = "" WHERE RCB_ORIGINECIBLAGE IS NULL');

  //DBR 4548
  ExecuteSQLContOnExcept ('update dechamps set dh_libelle="Type d''unité de manutention" where dh_nomchamp="ELI_REFUM"');

  // initialisation des séquences à partir des souches de type GES
  GCInitSequenceSouche;
End;

procedure MajVer948;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //M MORRETTON 4447
  ExecuteSQLNoPCL('UPDATE PARAMSOC'
              +' SET SOC_DATA= IIF((SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_PDRMETHVALOSTD")="-", "001"'
                           +', IIF((SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_PSTDLISTE")="-", "005", "004"))'
              +' WHERE SOC_NOM="SO_METHVALOCOUTSTD"');
  //T PETETIN 4587
//  ExecuteSQLNoPCL('UPDATE WORDRELIG SET WOL_ECARTUNIAREC=0');

  // T PETETIN 4588
  // Demande 4588
//  EXECUTESQLNOPCL('UPDATE WORDREGAMME SET WOG_ECARTUNIACC=0,WOG_ECARTUNIPROD=0,WOG_ECARTUNISUSP=0,WOG_ECARTUNIREBU=0'
//  +',WOG_ECARTUNIBONN=0');

  //D KOZA 4592
//  ExecuteSqlNoPCL('UPDATE QTEMPP SET QTP_POIDS=0 WHERE QTP_POIDS IS NULL');
//  ExecuteSqlNoPCL('UPDATE QSIMULATION SET QSM_FILTREBES="" WHERE QSM_FILTREBES IS NULL');
//  ExecuteSqlNoPCL('UPDATE QARTTECH SET QAR_VALCARLIB016="",QAR_VALCARLIB017="",QAR_VALCARLIB018="",QAR_VALCARLIB019="",QAR_VALCARLIB020="",QAR_VALCARLIB021="" WHERE QAR_VALCARLIB016 IS NULL');
//  ExecuteSqlNoPCL('UPDATE QCOULTECH SET QC7_VALCARCOLIB001="",QC7_VALCARCOLIB002="",QC7_VALCARCOLIB004="",QC7_VALCARCOLIB005=""'
//  +',QC7_VALCARCOLIB006="",QC7_VALCARCOLIB007="",QC7_VALCARCOLIB008="",QC7_VALCARCOLIB009="",QC7_VALCARCOLIB010=""'
//  +',QC7_VALCARCOLIB011="",QC7_VALCARCOLIB012="",QC7_VALCARCOLIB013="",QC7_VALCARCOLIB014="",QC7_VALCARCOLIB015=""'
//  +',QC7_VALCARCOLIB016="",QC7_VALCARCOLIB017="",QC7_VALCARCOLIB018="",QC7_VALCARCOLIB019="",QC7_VALCARCOLIB020="",QC7_VALCARCOLIB021="" WHERE QC7_VALCARCOLIB001 IS NULL');
//  ExecuteSqlNoPCL('UPDATE QGROUPE SET QGR_UNITECAD="" WHERE QGR_UNITECAD IS NULL');
//  ExecuteSqlNoPCL('UPDATE QCIRCUIT SET QCI_NATURETRAVAIL="1" WHERE (QCI_NATURETRAVAIL IS NULL) OR (QCI_NATURETRAVAIL="")');

  //MN GARNIER 4596
  //Suppresion de l'ancien paramétre origine ciblage
//  ExecuteSQLContOnExcept( 'DELETE FROM COMMUN WHERE CO_TYPE="RCO" AND CO_CODE="ASM"' );

  //M MORRETTON 4601
//  ExecuteSQLNoPCL('UPDATE STKVALOPARAM SET GVP_SENSMVTS="" WHERE GVP_SENSMVTS IS NULL');
//  ExecuteSQLNoPCL('UPDATE WPARAMFONCTION SET WPF_SENSMVTS="" WHERE WPF_SENSMVTS IS NULL');

//  AglNettoieListesPlus('GCSTKVALOPARAM', 'GVP_SENSMVTS', nil, true);
//  AglNettoieListesPlus('GCSTKVALOARTICLE', 'WPF_SENSMVTS', nil, true);

  //MNG 4612
    // ajout d'un champ type de modèle de chaînage
  //ExecuteSqlNoPCL  ('UPDATE PARCHAINAGES SET RPG_TYPECHAINAGE="" ');
//  ExecuteSqlNoPCL  ('UPDATE PARCHAINAGES SET RPG_TYPECHAINAGE="" WHERE RPG_TYPECHAINAGE IS NULL');

  //C DUMAS 4623
  //  Suite FQ 010/17150 : pour mettre à jour le champ GPA_AUXICONTACT qui parfois contient un code tiers au lieu d'un code auxiliaire
(*
  ExecuteSQLNoPCL ('update pieceadresse set gpa_auxicontact = (select t_auxiliaire from tiers where t_tiers=gpa_auxicontact) '
  +' where gpa_auxicontact<>"" and (select t_auxiliaire from tiers where t_tiers=gpa_auxicontact) is not null and gpa_typecontact = "T"');
*)
  //O BERTHIER 4611
  ExecuteSQLContOnExcept ('update contact set C_TYPECONTACT="T" where C_TYPECONTACT="TIE"');

  //M DESGOUTTE 4644
  If IsMonoOuCommune then
    ExecuteSQLContOnExcept('UPDATE DOSSIER SET DOS_TYPEDOSSIER="PRO"');

  // C PARWEZ 4646
//  ExecuteSQLNoPcl('UPDATE STKFICHETRACE SET GST_GUID = PGIGUID WHERE GST_GUID IS NULL');
//GP_20091104_PCO_GP16737 deb
//  AGLNettoieListesPlus('GCSTKFICHELOT','GST_GUID',nil,True);
//  AGLNettoieListesPlus('GCSTKFICHESERIE','GST_GUID',nil,True);
//GP_20091104_PCO_GP16737 fin

  //MCD 4647
//  AglNettoieListesPlus('AFMULAFFAIRE', 'AFF_TYPEAFFAIRE', nil, true);

  //MCD 4648
(*
  ExecuteSQLNoPCL('UPDATE TACHE SET ATA_CODEPHASEAFF1="",ATA_CODEPHASEAFF2="",'
      +'ATA_CODEPHASEAFF3="",ATA_CODEPHASEAFF4="",ATA_CODEPHASEAFF5="",'
      +'ATA_PHASEAFF1="",ATA_PHASEAFF2="",ATA_PHASEAFF3="",ATA_PHASEAFF4="",ATA_PHASEAFF5=""'
      +'  where  ATA_CODEPHASEAFF1 is null');
*)
  ExecuteSqlContOnExcept('UPDATE RESSOURCE  SET ARS_SURBOOK="-" WHERE ARS_SURBOOK is Null');
(*
  ExecuteSQLNoPCL('UPDATE FACTAFF  SET AFA_BLOQUE="-",AFA_CODEPHASEAFF1="",AFA_CODEPHASEAFF2="",'
      +'AFA_CODEPHASEAFF3="",AFA_CODEPHASEAFF4="",AFA_CODEPHASEAFF5="",'
      +'AFA_NUMORDRE=0,AFA_PHASEAFF1="",AFA_PHASEAFF2="",AFA_PHASEAFF3="",AFA_PHASEAFF4="",AFA_PHASEAFF5="" '
      +' where  afa_codephaseaff1 is null');
*)
  //S MASSON 4650
  ExecuteSqlNoPcl('UPDATE DPAGRDOMC SET DOC_OBLIGATOIRE = "-"');

  //G JUGDE 4653
  AglNettoieListesPlus( 'PSE_MULQOPOINTAGE', 'EOP_GUID', nil, False, '' );
  AglNettoieListesPlus( 'PSE_MULQWHISTORES', 'EWH_GUID', nil, False, '' );

  // C DUMAS 4656
  ExecuteSQLContOnExcept ('UPDATE CONTACT SET C_TIERS=(SELECT T_TIERS FROM TIERS WHERE '
  +'(T_NATUREAUXI=CONTACT.C_NATUREAUXI) and (T_AUXILIAIRE=CONTACT.C_AUXILIAIRE)) WHERE C_TYPECONTACT="T" AND C_TIERS = ""');

  //D SCLAVOPOULOS 4658
//  AGLNettoieListesPlus('WCOMPAREWNL','WNL_PHASE',nil,True);

End;

procedure MajVer949;
var StEtabDefaut : string;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //S MASSON 4714
  ExecuteSqlNoPcl('delete from CHOIXDPSTD where yds_type = "DAG" and YDS_NODOSSIER = "000000" and yds_predefini = "STD"');
  ExecuteSqlNoPcl('delete from DPAGRDOML where DOL_CODEDOMINANTE < "1000"');
  ExecuteSqlNoPcl('delete from DPAGRDOMD where DOD_CODEDOMINANTE < "1000"');
  ExecuteSqlNoPcl('delete from DPAGRDOMN where DON_CODEDOMINANTE < "1000"');
  ExecuteSqlNoPcl('delete from DPAGRCARTECH where DCT_CODECARTECH < "1000"');
  ExecuteSqlNoPcl('delete from DPAGRDOMC where DOC_CODECARTECH < "1000"');
  ExecuteSqlNoPcl('delete from DPAGRDOME where DOE_CODEDOMINANTE1 < "1000" AND DOE_CODEDOMINANTE2 < "1000"');

  //MCD 4675
//  ExecuteSqlNoPCL('update afsegcriteremission set asi_opercomp=""');
  //D KOZA 4677
//  ExecuteSqlNoPCL('UPDATE QSIMULATION SET QSM_RESULTATCALCUL="1" WHERE QSM_RESULTATCALCUL IS NULL');
  //MNG 4693
    // initialisation des nouveaux champs infos compl actions
(*
  ExecuteSqlNoPcl('UPDATE RTINFOS00V SET RDV_RDVLIBTEXTE5="", RDV_RDVLIBTEXTE6="",RDV_RDVLIBTEXTE7="",RDV_RDVLIBTEXTE8="",'+
  'RDV_RDVLIBTEXTE9="", RDV_RDVLIBVAL5=0,RDV_RDVLIBVAL6=0,RDV_RDVLIBVAL7=0,RDV_RDVLIBVAL8=0,RDV_RDVLIBVAL9=0,'+
  'RDV_RDVLIBTABLE5="", RDV_RDVLIBTABLE6="", RDV_RDVLIBTABLE7="", RDV_RDVLIBTABLE8="", RDV_RDVLIBTABLE9="", '+
  'RDV_RDVLIBMUL5="", RDV_RDVLIBMUL6="", RDV_RDVLIBMUL7="", RDV_RDVLIBMUL8="", RDV_RDVLIBMUL9="", '+
  'RDV_RDVLIBBOOL5="-", RDV_RDVLIBBOOL6="-", RDV_RDVLIBBOOL7="-", RDV_RDVLIBBOOL8="-", RDV_RDVLIBBOOL9="-", '+
  'RDV_RDVLIBDATE5="'+UsDateTime_(iDate1900)+'", RDV_RDVLIBDATE6="'+UsDateTime_(iDate1900)+'", RDV_RDVLIBDATE7="'+UsDateTime_(iDate1900)+'", '+
  'RDV_RDVLIBDATE8="'+UsDateTime_(iDate1900)+'", RDV_RDVLIBDATE9="'+UsDateTime_(iDate1900)+'" '+
  'WHERE RDV_RDVLIBTEXTE5 IS NULL');
  // MNG 4693
  // MNG FQ;012;11069 ajout des codes pour les nouveaux champs infos compl propositions
  RT_InsertLibelleInfoComplPropo();
*)
  // C DUMAS 4701
  //Plande dév : 010/10018
  ExecuteSQLContOnExcept('UPDATE ADRESSES SET ADR_FAX="",ADR_TELEX="" WHERE ADR_FAX IS NULL');
//  ExecuteSQLNoPCL('update PIECEADRESSE set GPA_TELEPHONE="",GPA_FAX="",GPA_TELEX="" WHERE GPA_TELEPHONE IS NULL');
//  ExecuteSQLNoPCL('update EADRESSES set EAG_FAX="",EAG_TELEX="" WHERE EAG_FAX IS NULL');
//  ExecuteSQLNoPCL('update EPIECEADRESSE set EPA_TELEPHONE="",EPA_FAX="",EPA_TELEX="" WHERE EPA_TELEPHONE IS NULL');
//  ExecuteSQLNoPCL('update ADRESSESAFF set ADA_FAX="",ADA_TELEX="" WHERE ADA_FAX IS NULL');

  //M DESGOUTTE 4707
  If IsMonoOuCommune then
  ExecuteSQLContOnExcept('UPDATE DPMOYPRO SET DCA_SUPERVISE="-", DCA_PERIODEPLUS="-"');

  //P DAMINETTE 4726
  ExecuteSQLNoPCL('UPDATE IAIDSAI SET IAI_INVISIBLE = "-"');
  //P DAMINETTE 4728
  ExecuteSQLNoPCL('UPDATE IBUDGETENG SET IBE_COMPTEFOUR = ""');
  //P DAMINETTE 4729
  ExecuteSQLNoPCL('UPDATE IBUDGETLIGNE SET IBL_AUTOMATIC = "-"');
  //P DAMINETTE 4730
  ExecuteSQLNoPCL('UPDATE IBUDGETREA SET IBT_COMPTEFOUR = ""');
  //P DAMINETTE 4731
  ExecuteSQLNoPCL('UPDATE ICODGEO SET IGO_INVISIBLE = "-"');
  //P DAMINETTE 4732
  ExecuteSQLNoPCL('UPDATE ICODP00 SET I00_INVISIBLE = "-"');
  //P DAMINETTE 4733
  ExecuteSQLNoPCL('UPDATE ICODSPC SET ISP_INVISIBLE = "-"');
  //P DAMINETTE 4735
  ExecuteSQLNoPCL('UPDATE IDUREETAUX SET '
    + 'IDT_COEFDEG7= 0, IDT_DATECOEFDEG7 = "'+UsDateTime_(iDate1900)+'",'
    + ' IDT_COEFDEG8= 0, IDT_DATECOEFDEG8 = "'+UsDateTime_(iDate1900)+'",'
    + ' IDT_COEFDEG9= 0, IDT_DATECOEFDEG9 = "'+UsDateTime_(iDate1900)+'",'
    + ' IDT_COEFDEG10= 0, IDT_DATECOEFDEG10 = "'+UsDateTime_(iDate1900)+'",'
    + ' IDT_COEFDEG11= 0, IDT_DATECOEFDEG11 = "'+UsDateTime_(iDate1900)+'",'
    + ' IDT_COEFDEG12= 0, IDT_DATECOEFDEG12 = "'+UsDateTime_(iDate1900)+'",'
    + ' IDT_COEFDEG13= 0, IDT_DATECOEFDEG13 = "'+UsDateTime_(iDate1900)+'",'
    + ' IDT_COEFDEG14= 0, IDT_DATECOEFDEG14 = "'+UsDateTime_(iDate1900)+'",'
    + ' IDT_COEFDEG15= 0, IDT_DATECOEFDEG15= "'+UsDateTime_(iDate1900)+'",'
    + ' IDT_INVISIBLE="-"');

  //P DAMINETTE 4736
  ExecuteSQLNoPCL('UPDATE IECRLGN SET ILG_ENTITY = 0, ILG_DATECHEANCE="'+UsDateTime_(iDate1900)+'", '
  +'ILG_NUMECHE=0, ILG_NATUREPIECE=""');

  //P DAMINETTE 4737
  ExecuteSQLNoPCL('UPDATE IMOIMP SET '
    + 'IMP_ENTITY = 0,IMP_NUMENTITY=IMP_NUM,'
    + ' IMP_SPEC01="",IMP_SPEC02="",IMP_SPEC03="",IMP_SPEC04="",IMP_SPEC05="",IMP_SPEC06="",'
    + ' IMP_VALRES01=0, IMP_VALRES02=0,IMP_VALRES03=0,IMP_VALRES04=0,IMP_VALRES05=0,IMP_VALRES06=0,'
    + ' IMP_UONAT01="", IMP_UONAT02="", IMP_UONAT03="", IMP_UONAT04="", IMP_UONAT05="", IMP_UONAT06="",'
    + ' IMP_UOTOT01=0, IMP_UOTOT02=0, IMP_UOTOT03=0, IMP_UOTOT04=0, IMP_UOTOT05=0, IMP_UOTOT06=0,IMP_VARBASE01=0,'
    + ' IMP_VARBASE02=0, IMP_VARBASE03=0, IMP_VARBASE04=0, IMP_VARBASE05=0, IMP_VARBASE06=0,IMP_VALFIN01=0,'
    + ' IMP_VALFIN02=0, IMP_VALFIN03=0, IMP_VALFIN04=0, IMP_VALFIN05=0, IMP_VALFIN06=0');

  //P DAMINETTE 4738
  ExecuteSQLNoPCL('UPDATE IMOINTERSOC SET IRS_ENTITYORIGINE = 0, IRS_ENTITYDESTINAT = 0,'
    + ' IRS_BASEORIGINE ="", IRS_BASDESTINATION=""');

  //P DAMINETTE 4739
   ExecuteSQLNoPCL('UPDATE IMOLSNEW SET ILN_ENTITY = 0, ILN_NUMFIENTITY = ILN_NUMFI, ILN_TEG = 0');

  //P DAMINETTE 4740
   ExecuteSQLNoPCL('UPDATE IMORECUPTEMP SET IRT_ENTITY = 0, IRT_NUMENTITY = IRT_NUM');

   // R HARANG N° 4752
  (*
  StEtabDefaut := GetParamSocSecur('SO_ETABLISDEFAUT','');
  ExecuteSQLNoPCL('UPDATE HRALLOTEMENT SET HAL_ETABLISSEMENT = "'+StEtabDefaut+'"'
    +' WHERE HAL_ETABLISSEMENT IS NULL');
  ExecuteSQLNoPCL('UPDATE HRPARAMPLANNING SET HPP_ETABLISSEMENT = "'+StEtabDefaut+'"'
    +' WHERE HPP_ETABLISSEMENT IS NULL');
  ExecuteSQLNoPCL('UPDATE HRPENSION SET HFP_ETABLISSEMENT = "'+StEtabDefaut+'"'
    +' WHERE HFP_ETABLISSEMENT IS NULL');
  ExecuteSQLNoPCL('UPDATE HRRESINTERFACE SET HRF_ETABLISSEMENT = "'+StEtabDefaut+'"'
    +' WHERE HRF_ETABLISSEMENT IS NULL');
  ExecuteSQLNoPCL('UPDATE HRCONTINGENT SET HCG_ETABLISSEMENT = "'+StEtabDefaut+'"'
    +' WHERE HCG_ETABLISSEMENT IS NULL');
  ExecuteSQLNoPCL('UPDATE HRCONTRAT SET HCO_ETABLISSEMENT = "'+StEtabDefaut+'"'
    +' WHERE HCO_ETABLISSEMENT IS NULL');
  ExecuteSQLNoPCL('UPDATE HRDOSRES SET HDR_ETABLISSEMENT = "'+StEtabDefaut+'"'
    +' WHERE HDR_ETABLISSEMENT IS NULL');
  ExecuteSQLNoPCL('UPDATE HRDOSSIER SET HDC_ETABLISSEMENT ="'+StEtabDefaut+'"'
    +' WHERE HDC_ETABLISSEMENT IS NULL');
  ExecuteSQLNoPCL('UPDATE HRFAMRES SET HFR_ETABLISSEMENT ="'+StEtabDefaut+'"'
    +' WHERE HFR_ETABLISSEMENT IS NULL');
  ExecuteSQLNoPCL('UPDATE HRNBPERSONNE SET HNP_ETABLISSEMENT ="'+StEtabDefaut+'"'
    +' WHERE HNP_ETABLISSEMENT IS NULL');
  ExecuteSQLNoPCL('UPDATE HRTYPRES SET HTR_ETABLISSEMENT ="'+StEtabDefaut+'"'
    +' WHERE HTR_ETABLISSEMENT IS NULL');
  *)
  //MCD 4754
  //Init pour GIGA
//  ExecuteSQLNoPCL ('UPDATE LIGNECOMPLAFF  SET GLA_RESPONSABLE="",GLA_FONCTION = "", GLA_UNITETEMPS = "J" WHERE GLA_RESPONSABLE is Null');
//  ExecuteSQLNoPCL ('UPDATE TACHE SET ATA_DATEDEBGENER="" WHERE ATA_DATEDEBGENER IS NULL');
//  ExecuteSQLNoPCL ('UPDATE AFMODELETACHE SET AFM_DATEDEBGENER="" WHERE AFM_DATEDEBGENER IS NULL');

//  ExecuteSQLNoPCL ('UPDATE AFFAIRE SET AFF_MODEFACTPHASE="-",AFF_PHASEAFF="-" WHERE AFF_PHASEAFF is Null');



  //P DAMINETTE 4755
  ExecuteSQLNoPCL ('UPDATE IMOREF SET IRF_ENTITY = 0, IRF_NUMENTITY = IRF_NUM'
  +',IRF_PHOTO_DATE1="'+ UsDateTime_(iDate1900) + '",IRF_PHOTO_DATE2="'+ UsDateTime_(iDate1900) + '",'
  +'IRF_POINT_INFOD10="'+ UsDateTime_(iDate1900) + '",IRF_POINT_INFOD11="'+ UsDateTime_(iDate1900) + '",IRF_POINT_INFOD12="'+ UsDateTime_(iDate1900) + '",'
  +'IRF_POINT_INFOD13="'+ UsDateTime_(iDate1900) + '",IRF_POINT_INFOD14="'+ UsDateTime_(iDate1900) + '",IRF_POINT_INFOD15="'+ UsDateTime_(iDate1900) + '",'
  +'IRF_POINT_INFOD16="'+ UsDateTime_(iDate1900) + '",IRF_POINT_INFOD17="'+ UsDateTime_(iDate1900) + '",IRF_POINT_INFOD18="'+ UsDateTime_(iDate1900) + '",'
  +'IRF_POINT_INFOA10="",IRF_POINT_INFOA11="",IRF_POINT_INFOA12="",IRF_POINT_INFOA13="",'
  +'IRF_POINT_INFOA14="",IRF_POINT_INFOA15="",IRF_POINT_INFOA16="",IRF_POINT_INFOA17="", IRF_POINT_INFOA18="",IRF_POINT_INFOB10="-",IRF_POINT_INFOB11="-",IRF_POINT_INFOB12="-",'
  +'IRF_POINT_INFOB13="-",IRF_POINT_INFOB14="-",IRF_POINT_INFOB15="-",IRF_POINT_INFOB16="-",'
  +'IRF_POINT_INFOB17="-",IRF_POINT_INFOB18="-",IRF_PHOTO_TEXTE="", IRF_NATURE="",IRF_39BIS="",'
  +'IRF_INTERNE1="'+ UsDateTime_(iDate1900) + '",IRF_INTERNE2="'+ UsDateTime_(iDate1900) + '",IRF_LOCMONTANT=0,IRF_LOCPLAFOND=0,IRF_LOCPLANCHER=0,IRF_LOCBASE=0,IRF_LOCDATEDEB="'+ UsDateTime_(iDate1900) + '",IRF_LOCDATEFIN="'+ UsDateTime_(iDate1900) + '"');

  //P DAMINETTE 4756
  ExecuteSQLNoPCL ('UPDATE IMOSUBREF SET IBR_ENTITY = 0, IBR_NUMSUBENTITY = IBR_NUMSUB');
  //P DAMINETTE 4757
  ExecuteSQLNoPCL ('UPDATE INVFIC SET INV_ENTITY = 0, INV_NUMENTITY = INV_NUM');
  //P DAMINETTE 4758
  ExecuteSQLNoPCL ('UPDATE INVFICBACK SET IBK_ENTITY = 0, IBK_NUMENTITY = IBK_NUM');
  //P DAMINETTE 4759
  ExecuteSQLNoPCL ('UPDATE INVLOT SET ILO_ENTITY = 0');
  //P DAMINETTE 4760
  ExecuteSQLNoPCL ('UPDATE INVSYNTHESE SET IVY_ENTITY = 0, IVY_NUMENTITY = IVY_NUM');
  //P DAMINETTE 4761
  ExecuteSQLNoPCL ('UPDATE IPERCAL2 SET IPL_ENTITY = 0');
  //P DAMINETTE 4762
  ExecuteSQLNoPCL ('UPDATE ISCENARIO SET ISE_INVISIBLE = "-"');
  //P DAMINETTE 4763
  ExecuteSQLNoPCL ('UPDATE ISUBSCENARIO SET ISS_INVISIBLE = "-"');

{Remplacé par 4986  //M MORRETTON 4767
  ExecuteSQLNoPCL('UPDATE WPARAM SET WPA_BOOLEAN17="-", WPA_BOOLEAN18="-", WPA_BOOLEAN19="-", WPA_BOOLEAN20="-", WPA_BOOLEAN21="-", WPA_BOOLEAN22="-", WPA_BOOLEAN23="-", WPA_BOOLEAN24="-", WPA_BOOLEAN25="-" WHERE WPA_BOOLEAN17 IS NULL');
  ExecuteSQLNoPCL('UPDATE WPDRTYPE SET WRT_AVECQECO="-", WRT_AVECQPCBTARIF="-", WRT_AVECQECOTARIF="-" WHERE WRT_AVECQECO IS NULL');
  ExecuteSQLNoPCL('UPDATE WPDRTET SET WPE_AVECQECO="-", WPE_AVECQPCBTARIF="-", WPE_AVECQECOTARIF="-" WHERE WPE_AVECQECO IS NULL');}

  //M MORRETTON 4986
  (*
  ExecuteSQLNoPCL('UPDATE WPDRTET'
  +' SET WPE_AVECQECO=IIF((INSTR(WPE_LIBELLE2,";QVECO;")>0), "X", "-")'
  +' ,WPE_AVECQPCBTARIF=IIF((INSTR(WPE_LIBELLE2,";QTPCB;")>0), "X", "-")'
  +' ,WPE_AVECQECOTARIF=IIF((INSTR(WPE_LIBELLE2,";QTECO;")>0), "X", "-")'
  +' WHERE (WPE_NATUREPDR="BTH") AND ((INSTR(WPE_LIBELLE2, ";QVECO;")>0) OR (INSTR(WPE_LIBELLE2, ";QTPCB;")>0) OR (INSTR(WPE_LIBELLE2, ";QTECO;")>0))');

  ExecuteSQLNoPCL('UPDATE WPDRTYPE'
  +' SET WRT_AVECQECO=IIF((INSTR(WRT_LIBELLE,";QVECO;")>0), "X", "-")'
  +' ,WRT_AVECQPCBTARIF=IIF((INSTR(WRT_LIBELLE,";QTPCB;")>0), "X", "-")'
  +' ,WRT_AVECQECOTARIF=IIF((INSTR(WRT_LIBELLE,";QTECO;")>0), "X", "-")'
  +' WHERE (INSTR(WRT_LIBELLE, ";QVECO;")>0) OR (INSTR(WRT_LIBELLE, ";QTPCB;")>0) OR (INSTR(WRT_LIBELLE, ";QTECO;")>0)');

  ExecuteSQLNoPCL('UPDATE WPARAM'
  +' SET WPA_BOOLEAN17=IIF((INSTR(WPA_LONGVARCHAR02,";QVECO;")>0), "X", "-")'
  +' ,WPA_BOOLEAN18=IIF((INSTR(WPA_LONGVARCHAR02,";QTPCB;")>0), "X", "-")'
  +' ,WPA_BOOLEAN19=IIF((INSTR(WPA_LONGVARCHAR02,";QTECO;")>0), "X", "-")'
  +' ,WPA_BOOLEAN20="-" ,WPA_BOOLEAN21="-" ,WPA_BOOLEAN22="-" ,WPA_BOOLEAN23="-" ,WPA_BOOLEAN24="-" ,WPA_BOOLEAN25="-"'
  +' WHERE WPA_CODEPARAM="PRIXDEREVIENT" AND ((INSTR(WPA_LONGVARCHAR02, ";QVECO;")>0) OR (INSTR(WPA_LONGVARCHAR02, ";QTPCB;")>0) OR (INSTR(WPA_LONGVARCHAR02, ";QTECO;")>0))');
  *)
  //M MORRETTON 4782
  ExecuteSQLNoPCL('UPDATE EXERCICE SET EX_ENTITY=0 WHERE EX_ENTITY IS NULL');

  //P DAMINETTE 4783
  ExecuteSQLNoPCL ('update cpprorata set pa_entity=0, pa_general = "", pa_exclusion = "-",'
    +'pa_coef11=0, pa_coef12=0, pa_coef13=0,'
    +'pa_coef21=0, pa_coef22=0, pa_coef33=0,'
    +'pa_coef31=0, pa_coef32=0, pa_coef23=0');
End;

procedure MajVer951;
Begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //P DAMINETTE 4734
  ExecuteSQLNoPCL('UPDATE IDATCLO SET IDO_ENTITY = 0');
  //D SCLAVOPOULOS 4825
//  AglNettoieListes('WORDREBESL', 'GA_CODEDIM1;GA_CODEDIM2;GA_CODEDIM3;GA_CODEDIM4;GA_CODEDIM5', nil);

  //T SUBLET 4826
  //GP_20090416_TS : FQ;034;13889
  ExecuteSqlNoPCL('DELETE FROM MENU WHERE MN_1=215 AND MN_2=7 AND MN_3=11 AND MN_4>0');

  //JTR 4916
  { JTR le 29/04/2009 - Gestion éco-contribution sur les articles - TD14671-Début }
//  ExecuteSQLNoPCL('UPDATE LIGNEFRAIS SET LF_ELEMENTECO = ""');
  { JTR le 29/04/2009 - Gestion éco-contribution sur les articles - TD14671-Fin }

end;

procedure MajVer952;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //C DUMAS 4928
  ExecuteSQLContOnExcept( 'UPDATE DETABLES SET DT_LIBELLE = "Fourchette de valeurs des critères" '
  +'WHERE DT_PREFIXE = "RQF" AND DT_LIBELLE = "Fourchette de valeur des criteres"' );

End;

procedure MajVer953;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //X PERSOUYRE 4969
  ExecuteSQLContOnExcept( 'UPDATE DECOMBOS SET DO_SORT=""');

  //D KOZA 4978
  ExecuteSqlNoPCL('UPDATE USERGRP SET UG_ENVIRONNEMENT="0" WHERE (UG_ENVIRONNEMENT IS NULL) OR (UG_ENVIRONNEMENT="")');
//  ExecuteSqlNoPCL('UPDATE QGROUPE SET QGR_CTXORDO="0" WHERE (QGR_CTXORDO IS NULL) OR (QGR_CTXORDO="")');
//  ExecuteSqlNoPCL('UPDATE QSIMULATION SET QSM_CTXORDO="0" WHERE (QSM_CTXORDO IS NULL) OR (QSM_CTXORDO="")');

end;

procedure MajVer954;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  // Lek 4992
  CpInitSequenceCompta;
  // JTR 5001
  {JTR - Pour affichage des champs dans une FSL }
//  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = "LDZV" WHERE DH_PREFIXE = "GEE" AND DH_CONTROLE <> ""');
//  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = "LDZ" WHERE DH_PREFIXE = "GE1"');

  //DBR 5009
//  ExecuteSQLNoPCL( 'UPDATE PIECE SET GP_APPORTEURTIERS="" WHERE GP_APPORTEURTIERS is null' );

{  //MNG 5021
  // appel fonction de mise à jour ytradmetier
  CRMMajYTradMetier;}


End;

procedure MajVer955;
Var
  vSt : string;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //F BERGER 5027
  AglNettoieListes('WGAMMECIR', 'WGC_MAJEUR', nil);
  // G KIELWASSER 5028
  {GC GKI suite à plusieurs FQ , correction sur plusieurs libellés des champs}
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET dh_libelle="Valeur de la rupture 4"  WHERE dh_prefixe="ELI" and dh_nomchamp="ELI_QTECOND"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET dh_libelle="Article fantôme"  WHERE dh_prefixe="GA" and dh_nomchamp="GA_FANTOME"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET dh_libelle="Affichage avancé des dimensions"  WHERE dh_prefixe="GPP" and dh_nomchamp="GPP_PARAMGRILLEDIM"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET dh_libelle="Identification des derniers BL"  WHERE dh_prefixe="EMG" and dh_nomchamp="EMG_LASTBL"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET dh_libelle="Concept associé au contrôle"  WHERE dh_prefixe="GP1" and dh_nomchamp="GP1_CONCEPT"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET dh_libelle="Régime taxe"  WHERE dh_prefixe="GIC" and dh_nomchamp="GIC_REGIMETAXE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET dh_libelle="Dépôt émetteur"  WHERE dh_prefixe="GTE" and dh_nomchamp="GTE_DEPOTEMET"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET dh_libelle="Journal d''imputation règlement"  WHERE dh_prefixe="MPC" and dh_nomchamp="MPC_JALREGLE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET dh_libelle="Dépôt par défaut"  WHERE dh_prefixe="GPF" and dh_nomchamp="GPF_DEPOT"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET dh_libelle="Mouvement comptabilisé"  WHERE dh_prefixe="GSM" and dh_nomchamp="GSM_COMPTABILISE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET dh_libelle="Qualifiant coût de transport"  WHERE dh_prefixe="GA3" and dh_nomchamp="GA3_QUALIFCOUT"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET dh_libelle="Nombre d''articles par U.C"  WHERE dh_prefixe="ELI" and dh_nomchamp="ELI_QTECOND"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET dh_libelle="Date de suppression"  WHERE dh_prefixe="GCA" and dh_nomchamp="GCA_DATESUP"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET dh_libelle="Autoriser le recalcul a postériori"  WHERE dh_prefixe="GVT" and dh_nomchamp="GVT_RECALCUL"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET dh_libelle="Référence affectation"  WHERE dh_prefixe="GSM" and dh_nomchamp="GSM_REFAFFECTATION"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET dh_libelle="Nombre d''articles par U.C"  WHERE dh_prefixe="ELI" and dh_nomchamp="ELI_QTECOND"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET dh_libelle="Nombre d''articles par U.C"  WHERE dh_prefixe="GA2" and dh_nomchamp="GA2_QTEARTUC"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET dh_libelle="Nombre d''articles par U.C"  WHERE dh_prefixe="GAT" and dh_nomchamp="GAT_QTEARTUC"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Formulation (Nbr * Coef = Qté)" WHERE DH_PREFIXE="GLF" AND DH_NOMCHAMP="GLF_NBRCOEFQTE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Formulation (Nbr * Coef = Qté)" WHERE DH_PREFIXE="GAF" AND DH_NOMCHAMP="GAF_NBRCOEFQTE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Etat de la journée" WHERE DH_PREFIXE="GJC" AND DH_NOMCHAMP="GJC_ETAT"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Identifiant nomenclature GP" WHERE DH_PREFIXE="GLC" AND DH_NOMCHAMP="GLC_IDENTIFIANTWNT"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Code dépôt" WHERE DH_PREFIXE="GZI" AND DH_NOMCHAMP="GZI_DEPOT"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Code verrouillage recalcul frais" WHERE DH_PREFIXE="GPO" AND DH_NOMCHAMP="GPO_VERROU"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Modèle pour le fonds de caisse" WHERE DH_PREFIXE="GPK" AND DH_NOMCHAMP="GPK_IMPMODFDC"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Gestion du fonds de caisse" WHERE DH_PREFIXE="GPK" AND DH_NOMCHAMP="GPK_GEREFONDCAISSE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Mode bidirectionnel" WHERE DH_PREFIXE="GPK" AND DH_NOMCHAMP="GPK_PRTMODEBIDI"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Modification du fonds de caisse" WHERE DH_PREFIXE="GPK" AND DH_NOMCHAMP="GPK_MODIFFDCAIS"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Modes de paiement fonds de caisse" WHERE DH_PREFIXE="GPK" AND DH_NOMCHAMP="GPK_MDPFDCAIS"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Date de livraison" WHERE DH_PREFIXE="GZH" AND DH_NOMCHAMP="GZH_DATELIVRAISON"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Nature de la pièce" WHERE DH_PREFIXE="GZH" AND DH_NOMCHAMP="GZH_NATUREPIECEG"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Libellé article" WHERE DH_PREFIXE="GZH" AND DH_NOMCHAMP="GZH_LIBELLE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Code dépôt de la ligne de pièce" WHERE DH_PREFIXE="GZH" AND DH_NOMCHAMP="GZH_DEPOT"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Fonds de caisse à l''ouverture" WHERE DH_PREFIXE="GJM" AND DH_NOMCHAMP="GJM_FDCAISOUV"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Utilisation du bloc-notes article" WHERE DH_PREFIXE="GPP" AND DH_NOMCHAMP="GPP_BLOBLIENART"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Statut de flux imposé" WHERE DH_PREFIXE="GSN" AND DH_NOMCHAMP="GSN_SFLUXDISPATCH"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Calcul du PA à partir du PV" WHERE DH_PREFIXE="GA" AND DH_NOMCHAMP="GA_CALCULPA"');
  ExecuteSQLContOnExcept('UPDATE DETABLES SET DT_LIBELLE="Période d''application tarifs mode" WHERE DT_PREFIXE="GFP" AND DT_NOMTABLE="TARIFPER"');
  ExecuteSQLContOnExcept('UPDATE DETABLES SET DT_LIBELLE="Formule liée aux lignes de pièce" WHERE DT_PREFIXE="GLF" AND DT_NOMTABLE="LIGNEFORMULE"');
  ExecuteSQLContOnExcept('UPDATE DETABLES SET DT_LIBELLE="En-tête coût standard" WHERE DT_PREFIXE="CST" AND DT_NOMTABLE="COUTSTDTET"');
  ExecuteSQLContOnExcept('UPDATE DETABLES SET DT_LIBELLE="Lignes coût standard" WHERE DT_PREFIXE="CSL" AND DT_NOMTABLE="COUTSTDLIG"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Fonds de caisse" WHERE DH_NOMCHAMP="GPK_FDCAISSE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Fonds de caisse" WHERE DH_NOMCHAMP="GJM_FDCAISSEDEV"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="% qté réservée" WHERE DH_PREFIXE="GEP" AND DH_NOMCHAMP="GEP_POUQTERES"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Inventaire intégré" WHERE DH_PREFIXE="GIT" AND DH_NOMCHAMP="GIT_INTEGRATION"');
  ExecuteSQLContOnExcept('UPDATE DETABLES SET DT_LIBELLE="Paramétrage de la fidélité client" WHERE DT_PREFIXE="GFO" AND DT_NOMTABLE="PARFIDELITE"');
  ExecuteSQLContOnExcept('UPDATE DETABLES SET DT_LIBELLE="Pièce EDI à traiter" WHERE DT_PREFIXE="EPI" AND DT_NOMTABLE="EDIPIECE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Numéro unique" WHERE DH_PREFIXE="PRV" AND DH_NOMCHAMP="PRV_RANG"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Numéro unique" WHERE DH_PREFIXE="GZC" AND DH_NOMCHAMP="GZC_NUMERO"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Numéro ligne" WHERE DH_PREFIXE="GIC" AND DH_NOMCHAMP="GIC_NUMLIGNE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Numéro pièce" WHERE DH_PREFIXE="GIC" AND DH_NOMCHAMP="GIC_NUMERO"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Pièce de vente ou d''achat" WHERE DH_PREFIXE="GIC" AND DH_NOMCHAMP="GIC_VENTEACHAT"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Numéro de ticket" WHERE DH_PREFIXE="GZP" AND DH_NOMCHAMP="GZP_NUMERO"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Nature de la pièce" WHERE DH_PREFIXE="GTP" AND DH_NOMCHAMP="GTP_NATUREPIECEG"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Pièce transférée" WHERE DH_PREFIXE="EPI" AND DH_NOMCHAMP="EPI_TRANSFERT"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Pièce transférée" WHERE DH_PREFIXE="ELI" AND DH_NOMCHAMP="ELI_TRANSFERT"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Conditions Spéciales Tarification" WHERE DH_PREFIXE="GP" AND DH_NOMCHAMP="GP_TARIFSPECIAL"');
  ExecuteSQLContOnExcept('UPDATE DETABLES SET DT_LIBELLE="Choix qualité" WHERE DT_PREFIXE="GCQ" AND DT_NOMTABLE="GCCHOIXQUALITE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Champ critère n°3" WHERE DH_PREFIXE="GFR" AND DH_NOMCHAMP="GFR_CHAMP3"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Champ critère n°2" WHERE DH_PREFIXE="GFR" AND DH_NOMCHAMP="GFR_CHAMP2"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Champ critère n°1" WHERE DH_PREFIXE="GFR" AND DH_NOMCHAMP="GFR_CHAMP1"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="% qté affectée" WHERE DH_PREFIXE="GEP" AND DH_NOMCHAMP="GEP_POUQTEAFF"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Nature pièce" WHERE DH_NOMCHAMP="GCM_NATUREPIECEG"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Nature pièce" WHERE DH_NOMCHAMP="GIC_NATUREPIECEG"');

  //T PETETIN 5048
//  AGLNettoieListesPlus('GCTRACABILITELIG', 'GSM_DATEENTREELOT;GSM_DATEPEREMPTION;GSM_DATEDISPO', nil, True);
//  AGLNettoieListesPlus('GCTRACABILITEDISP', 'GQD_DATEPEREMPTION;GQD_DATEDISPO', nil, True);

  //MC DESSEIGNET 5051
{GM 5100  ExecuteSqlContOnExcept ('DELETE FROM LISTE WHERE LI_LISTE ="AFLISTETACHES"');}
 //C Bouet remplacement du paramsoc SO_AFQUELPLANNING par les 3 paramsocs suivants.
 (*
  vSt := GetParamsocSecur('SO_AFQUELPLANNING', 'PDC');
  if pos('DEU', vSt) > 0 then
  begin
    //ExecuteSqlContOnExcept('UPDATE PARAMSOC SET SOC_DATA = "X" WHERE SOC_NOM = "SO_AFGESTPLANNING"');
    //ExecuteSqlContOnExcept('UPDATE PARAMSOC SET SOC_DATA = "X" WHERE SOC_NOM = "SO_AFGESTPDCD"');
    SetParamSoc('SO_AFGESTPLANNING', 'X');
    SetParamSoc('SO_AFGESTPDCD', 'X');
  end
  else if pos('PDC', vSt) > 0 then
  begin
    //ExecuteSqlContOnExcept('UPDATE PARAMSOC SET SOC_DATA = "-" WHERE SOC_NOM = "SO_AFGESTPLANNING"');
    //ExecuteSqlContOnExcept('UPDATE PARAMSOC SET SOC_DATA = "X" WHERE SOC_NOM = "SO_AFGESTPDCD"');
    SetParamSoc('SO_AFGESTPLANNING', '-');
    SetParamSoc('SO_AFGESTPDCD', 'X');
  end
  else if pos('PLA', vSt) > 0 then
  begin
    //ExecuteSqlContOnExcept('UPDATE PARAMSOC SET SOC_DATA = "X" WHERE SOC_NOM = "SO_AFGESTPLANNING"');
    //ExecuteSqlContOnExcept('UPDATE PARAMSOC SET SOC_DATA = "-" WHERE SOC_NOM = "SO_AFGESTPDCD"');
    SetParamSoc('SO_AFGESTPLANNING', 'X');
    SetParamSoc('SO_AFGESTPDCD', '-');
  end;
  *)
  //G MERIEUX 5078
  // init pout GIGA
//  ExecuteSqlNoPCL('UPDATE ARTICLECOMPL SET GA2_TEMPSPAUSE="-", GA2_DUREEPAUSE=0 WHERE GA2_TEMPSPAUSE IS NULL');  // Lenotre
//  ExecuteSqlNoPCL('UPDATE AFPLANNING SET APL_QTEREALUFACT=0 WHERE APL_QTEREALUFACT IS NULL');
//  ExecuteSqlNoPCL('UPDATE ACTIVITE SET ACT_QTEREALUFACT=0,ACT_PUHTREALUFACT=0,ACT_PUHTDEVREALUFACT=0,ACT_LIBELLECOMPL = "", ACT_NUMINTER = 0, act_indice=0, ACT_ETATACTIVITE = "" WHERE ACT_QTEREALUFACT IS NULL');


//  ExecuteSqlNoPCL('UPDATE EACTIVITE SET EAC_QTEREALUFACT=0,EAC_PUHTREALUFACT=0,EAC_PUHTDEVREALUFACT=0,EAC_LIBELLECOMPL = "", EAC_NUMINTER = 0, EAC_indice=0, EAC_ETATACTIVITE = "" WHERE EAC_QTEREALUFACT IS NULL');



  //LEK CHHOEU 5081
  ExecuteSqlContOnExcept('UPDATE cpprorata set PA_GENERAL="", PA_EXCLUSION="-", PA_ENTITY=0,'+
	'PA_COEF11=IIF(PA_TAUX1=0, 0, 1), PA_COEF12=PA_TAUX1, PA_COEF13=IIF(PA_TAUX1=0, 0, 1),'+
	'PA_COEF21=IIF(PA_TAUX2=0, 0, 1), PA_COEF22=PA_TAUX2, PA_COEF23=IIF(PA_TAUX2=0, 0, 1),'+
	'PA_COEF31=IIF(PA_TAUX3=0, 0, 1), PA_COEF32=PA_TAUX3, PA_COEF33=IIF(PA_TAUX3=0, 0, 1) where PA_GENERAL is null');

End;

procedure MajVer957;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //M FAUDEL 5126
  // MSA : Initialisation des nouveaux champs de la table MSAPERIODESPE31
  ExecuteSQLContOnExcept ('UPDATE MSAPERIODESPE31 SET PE3_BOOL01 = "-",PE3_BOOL02 = "-",PE3_CARACT01 = ""'
  +',PE3_CARACT02 = "",PE3_COMBO01 = "",PE3_COMBO02 = "",PE3_DBLE01 = 0,PE3_DBLE02 = 0');
  // MSA : Initialisation des nouveaux champs de la table MSAEVOLUTIONSPE2
  ExecuteSQLContOnExcept ('UPDATE MSAEVOLUTIONSPE2 SET PE2_SAISONNIER = "",PE2_HOREQUIVALENCE = ""'
  +',PE2_COMBO01 = "",PE2_COMBO02 = "",PE2_COMBO03 = "",PE2_CARACT01 = "",PE2_CARACT02 = "",PE2_CARACT03 = ""'
  +',PE2_CARACT04 = "",PE2_CARACT05 = "",PE2_DBLE01 = 0,PE2_DBLE02 = 0,PE2_DBLE03 = 0,PE2_DBLE04 = 0,'
  +'PE2_DBLE05 = 0,PE2_BOOL01 = "-",PE2_BOOL02 = "-",PE2_BOOL03 = "-",PE2_BOOL04 = "-",PE2_BOOL05 = "-"');

  //D SCLAVOPOULOS 5167
//  AGLNettoieListes('GCMULLIGNEACH', 'GL_DATEPIECE', Nil);

  //G JUGDE
//  ExecuteSQLNoPCL ('UPDATE QWHISTORES SET QWH_MESSAGECAPT="" WHERE QWH_MESSAGECAPT IS NULL');

  //C DUMAS 5180
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Champ critère" WHERE DH_PREFIXE = "RTC" AND DH_NOMCHAMP = "RTC_CHAMP1"');

  //G KIELWASSER 5121
  ExecuteSQLcontonexcept('UPDATE DETABLES SET DT_LIBELLE="Nature pièce regroupement" WHERE DT_NOMTABLE="PARPIECEGRP"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Identifiant du paramétrage" WHERE DH_PREFIXE="GRF" AND DH_NOMCHAMP="GRF_PARFOU"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Numéro de série" WHERE DH_PREFIXE="GLS" AND DH_NOMCHAMP="GLS_IDSERIE"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Numéro de série" WHERE DH_PREFIXE="GCS" AND DH_NOMCHAMP="GCS_IDSERIE"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Dépôt destinataire" WHERE DH_PREFIXE="GP" AND DH_NOMCHAMP="GP_DEPOTDEST"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Reprendre bloc-notes tiers" WHERE DH_PREFIXE="GPP" AND DH_NOMCHAMP="GPP_BLOBTIERS"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Reprendre bloc-notes article" WHERE DH_PREFIXE="GPP" AND DH_NOMCHAMP="GPP_BLOBART"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Modifiée par l''utilisateur" WHERE DH_PREFIXE="GZQ" AND DH_NOMCHAMP="GZQ_UTILISATEUR"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Type de l''exécutable" WHERE DH_PREFIXE="GHE" AND DH_NOMCHAMP="GHE_EXECUTABLE"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Dépôt" WHERE DH_PREFIXE="GA2" AND DH_NOMCHAMP="GA2_DEPOT"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Libellé" WHERE DH_PREFIXE="DAV" AND DH_NOMCHAMP="DAV_LIBELLE"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Libellé" WHERE DH_PREFIXE="GCM" AND DH_NOMCHAMP="GCM_LIBELLE"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Libellé" WHERE DH_PREFIXE="GFM" AND DH_NOMCHAMP="GFM_LIBELLE"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Libellé" WHERE DH_PREFIXE="GFP" AND DH_NOMCHAMP="GFP_LIBELLE"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Libellé" WHERE DH_PREFIXE="GFT" AND DH_NOMCHAMP="GFT_LIBELLE"');
  ExecuteSQLcontonexcept('UPDATE DETABLES SET DT_LIBELLE="Table temporaire palmarès" WHERE DT_PREFIXE="GZA" AND DT_NOMTABLE="GCTMPPAL"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Statut de disponibilité imposé" WHERE DH_PREFIXE="GSN" AND DH_NOMCHAMP="GSN_SDISPODISPATCH"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Suite raison sociale ou prénom" WHERE DH_PReFIXE="GPA" AND DH_NOMCHAMP="GPA_LIBELLE"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Suite raison sociale ou prénom" WHERE DH_PREFIXE="EDA" AND DH_NOMCHAMP="EDA_LIBELLE2"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Date début" WHERE DH_PREFIXE="GFM" AND DH_NOMCHAMP="GFM_DATEDEBUT"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Date début" WHERE DH_PREFIXE="GFP" AND DH_NOMCHAMP="GFP_DATEDEBUT"');
  ExecuteSQLcontonexcept('UPDATE DETABLES SET DT_LIBELLE="Table temporaire statistiques" WHERE DT_PREFIXE="GZB" AND DT_NOMTABLE="GCTMPSTA"');
  ExecuteSQLcontonexcept('UPDATE DETABLES SET DT_LIBELLE="Infos complémentaires en-tête pièce" WHERE DT_PREFIXE="RDD" AND DT_NOMTABLE="RTINFOS00D"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Indice pièce" WHERE DH_PREFIXE="GIC" AND DH_NOMCHAMP="GIC_INDICE"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Compte d''imputation règlement" WHERE DH_PREFIXE="MPC" AND DH_NOMCHAMP="MPC_CPTEREGLE"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Comportement anal stock ligne" WHERE DH_PREFIXE="GPP" AND DH_NOMCHAMP="GPP_COMPSTOCKLIGNE"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Comportement anal stock pied" WHERE DH_PREFIXE="GPP" AND DH_NOMCHAMP="GPP_COMPSTOCKPIED"');
  ExecuteSQLcontonexcept('UPDATE DETABLES SET DT_LIBELLE="Infos complémentaires en-tête pièce" WHERE DT_PREFIXE="RDD" AND DT_NOMTABLE="RTINFOS00D"');
  ExecuteSQLcontonexcept('UPDATE DETABLES SET DT_LIBELLE="Articles consignés" WHERE DT_PREFIXE="GA3" AND DT_NOMTABLE="ARTICLECONSIGNES"');
  ExecuteSQLcontonexcept('UPDATE DETABLES SET DT_LIBELLE="Compteurs d''entrées établissement" WHERE DT_PREFIXE="GCE" AND DT_NOMTABLE="COMPTEURETAB"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Mode d''expédition" WHERE DH_PREFIXE="GP" AND DH_NOMCHAMP="GP_EXPEDITION"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Stk mini du dépôt" WHERE DH_PREFIXE="GZC" AND DH_NOMCHAMP="GZC_STOCKMIN"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Qté calculée" WHERE DH_PREFIXE="GZC" AND DH_NOMCHAMP="GZC_QTE"');

end;

procedure MajVer958;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //G KIELWASSER 5202
  // FQ 10258
  AGLnettoielistes('PSE_MULEXPORTS','',nil,'EDT_INTEGREAUTO');

  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Date de création" WHERE DH_PREFIXE="GNL" AND DH_NOMCHAMP="GNL_DATECREATION"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Bloc-notes" WHERE DH_PREFIXE="ETS" AND DH_NOMCHAMP="ETS_BLOCNOTE"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Bloc-notes" WHERE DH_PREFIXE="GCL" AND DH_NOMCHAMP="GCL_BLOCNOTE"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Bloc-notes" WHERE DH_PREFIXE="GHE" AND DH_NOMCHAMP="GHE_BLOCNOTES"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Bloc-notes" WHERE DH_PREFIXE="GIE" AND DH_NOMCHAMP="GIE_BLOCNOTE"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Bloc-notes" WHERE DH_PREFIXE="GJC" AND DH_NOMCHAMP="GJC_BLOCNOTE"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Bloc-notes" WHERE DH_PREFIXE="GJE" AND DH_NOMCHAMP="GJE_BLOCNOTE"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Bloc-notes" WHERE DH_PREFIXE="GSM" AND DH_NOMCHAMP="GSM_BLOCNOTE"');
  ExecuteSQLcontonexcept('UPDATE DECHAMPS SET DH_LIBELLE="Bloc-notes" WHERE DH_PREFIXE="GSE" AND DH_NOMCHAMP="GSE_BLOCNOTE"');

  //MNG 5205
    // appel fonction de mise à jour ytradmetier
// appelé en 959  CRMMajYTradMetier;

End;

procedure MajVer959;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //G KIELWASSER 5221
  ExecuteSQLcontonexcept('UPDATE paramsoc SET soc_design="G;;Lien Servantissimmo;5;15;559;271;X;-;15;9;115;24;;0;0;0;MS Sans Serif;-16777214;8;0;-;0;0;" WHERE soc_nom="sco_paramservantissimo"');
  ExecuteSQLcontonexcept('UPDATE paramsoc SET soc_design="B;;Traduction du libellé de l''article si fournisseur étranger;16;264;289;279;X;-;10;260;95;281;MS Sans Serif;-16777208;8;0;;0;0;0;X;0;0;" WHERE soc_nom="SO_TRADUCCONTREMARQUE"');
  ExecuteSQLcontonexcept('UPDATE paramsoc SET soc_design="A;200;Commande à exécuter;212;548;410;569;X;-;37;548;225;569;MS Sans Serif;-16777208;8;0;MS Sans Serif;-16777208;8;0;-;0;0;" WHERE soc_nom="SO_ECMDEXEC"');
  ExecuteSQLcontonexcept('UPDATE paramsoc SET soc_design="G;;Commandes cadencées;8;10;597;97;X;-;15;4;139;19;;0;0;0;MS Sans Serif;-2147483646;8;0;-;0;0;" WHERE soc_nom="SCO_GCCDEOUV"');
  ExecuteSQLcontonexcept('UPDATE paramsoc SET soc_design="N;;Méthode d''envoi;212;476;257;497;X;-;37;476;225;497;MS Sans Serif;-16777208;8;0;MS Sans Serif;-16777208;8;0;-;0;0;" WHERE soc_nom="SO_EMAILMETHOD"');
  //ExecuteSQLcontonexcept('UPDATE paramsoc SET soc_design="C;GCEMPLOIBLOB##;Mémo 1;261;591;375;612;X;-;204;591;256;612;MS Sans Serif;-2147483640;8;0;MS Sans Serif;-2147483640;8;0;-;0;0;-~" WHERE soc_nom="SO_GCCDMEM1"');
  //ExecuteSQLcontonexcept('UPDATE paramsoc SET soc_design="C;GCEMPLOIBLOB##;Mémo 2;261;613;375;634;X;-;204;613;256;634;MS Sans Serif;-2147483640;8;0;MS Sans Serif;-2147483640;8;0;-;0;0;-~" WHERE soc_nom="SO_GCCDMEM2"');
  //ExecuteSQLcontonexcept('UPDATE paramsoc SET soc_design="C;GCEMPLOIBLOB##;Mémo 3;261;635;375;656;X;-;204;635;256;656;MS Sans Serif;-2147483640;8;0;MS Sans Serif;-2147483640;8;0;-;0;0;-~" WHERE soc_nom="SO_GCCDMEM3"');
  //ExecuteSQLcontonexcept('UPDATE paramsoc SET soc_design="C;GCEMPLOIBLOB##;Mémo 4;261;657;375;678;X;-;204;657;256;678;MS Sans Serif;-2147483640;8;0;MS Sans Serif;-2147483640;8;0;-;0;0;-~" WHERE soc_nom="SO_GCCDMEM4"');

  //D SCLAVOPOULOS 5242
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Article dimensionné" WHERE DH_PREFIXE LIKE "W%" AND DH_LIBELLE = "Article dimentionné"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Article dimensionné de nomenclature" WHERE DH_NOMCHAMP = "GZQ_ARTICLEPERE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Qté suspecte saisie" WHERE DH_NOMCHAMP="WOP_QSUSPSAIS" AND DH_LIBELLE = "Qté supspecte saisie"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Nombre d''articles par U.C." WHERE DH_LIBELLE = "Nombre d''article par U.C."');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Code tiers (Fournisseur, Client, .)" WHERE DH_NOMCHAMP="WEV_TIERS"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Coût unitaire apporté par la phase" WHERE DH_NOMCHAMP="WPL_COUTPHASE"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Famille ressource de rattachement" WHERE DH_NOMCHAMP IN ("EGF_RESORIGINEMES","WOR_RESORIGINEMES","WGS_RESORIGINEMES")');

  //D SCLAVOPOULOS 5255
//  AglNettoieListes('WORDRELIG2', 'WOL_MAJEURWNT;WOL_MAJEURWGT');
End;

procedure MajVer960;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //MNG 5234 5270
  // appel fonction de mise à jour ytradmetier
//  CRMMajYTradMetier;

  //MCD 5277
//  ExecuteSQLNoPCL ('UPDATE ARTICLECOMPL SET GA2_DESACTCTRLQTE = "-" WHERE GA2_DESACTCTRLQTE IS NULL');

  //M FAUDEL 5282
  // FQ 16457 : le champ POG_REGROUPEMENT devient obligatoire
  AglNettoieListes('PGDUCSINIT','POG_REGROUPEMENT',nil);

  //JTR 5309
  { JTR FQ;010;17140}
  ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DESIGN = "19;Lien Servantissimmo;0;" WHERE SOC_TREE = "001;003;033;000;"');

End;

procedure MajVer961;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

{  //MNG 5321
  // appel fonction de mise à jour ytradmetier
  CRMMajYTradMetier;}


End;

procedure MajVer962;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //D SCLAVOPOULOS 5386
  ExecuteSQLContOnExcept('UPDATE COMMUN SET CO_LIBELLE="Monnaie de tenue" WHERE CO_TYPE="WTS" AND CO_CODE="M"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Reste à faire en monnaie de tenue" WHERE DH_NOMCHAMP="WLB_RESTEAFAIREM"');

  //C DUMAS 5415
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Valeur par défaut décision libre " || SubString(DH_LIBELLE,33,LEN(DH_LIBELLE))' +
  'where dh_prefixe = "RPA" and dh_libelle like "valeur par défaut booleén%"');

  //D SCLAVOPOULOS 5417
  ExecuteSQLContOnExcept('UPDATE DETABLES SET DT_LIBELLE="Arrondi conversion d''unités" WHERE DT_NOMTABLE="WCUARRONDI"');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_libelle="Article de nomenclature dimensionné" WHERE DH_NOMCHAMP IN ("WND_ARTICLEPERE","WNS_ARTICLEPERE")');

  //JTR 5420
  { GC_20090715_JTR_FQ;PGISIDE;10237 }
//  ExecuteSqlNoPCL ('UPDATE EARTICLECOMPL SET EA2_CODEIMPORT = "" WHERE EA2_CODEIMPORT IS NULL');
End;

procedure MajVer963;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //JTR 5452
  { GC_20090720_JTR_FQ;010;16997. Gestion des natures sur le contrôle des pièces d'achat }
//  ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA = "FF;FFF;" WHERE SOC_NOM = "SO_CTRLPCENATURES"');

  //DBR 5461
//  ExecuteSQLNoPCL ('UPDATE PARPIECE SET GPP_CONTEXTES = GPP_CONTEXTES || "INU;" WHERE GPP_NATUREPIECEG IN ("INV", "FRF", "SSE")');
//  ExecuteSQLNoPCL ('UPDATE PARPIECE SET GPP_CONTEXTES = "GC;" WHERE GPP_NATUREPIECEG = "FAA"');

  //RH 5487
  ExecuteSqlNoPCL('DELETE FROM MENU WHERE MN_1 =166 AND MN_2=8 AND MN_3=2 AND MN_4 IN (9,10,11,12,14,15) AND MN_TAG IN (112196,112197,999999)');

End;

procedure MajVer964;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //MNG 5497
  // MAJ de la liste GCMULPIECEACH par ajout d'un champ obligatoire FQ;034;15569
//  AglNettoieListes('GCMULPIECEACH', 'GP_DATEPIECE', nil ) ;

End;

procedure MajVer965;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //D SCLAVOPOULOS 5568
//  AGLNettoieListes('GCHISTORISEPIECE', 'GP_GUIDWOT', nil);
  //D SCLAVOPOULOS 5574
  ExecuteSQLContOnExcept ('DELETE FROM MENU WHERE MN_1=211 and mn_2=11');
  //T PETETIN 5590
//  ExecuteSQLContOnExcept('DELETE FROM LISTE WHERE LI_LISTE="GCVERIFQTERESTE"');
  //M MORRETTON 5616
//  ExecuteSQLContOnExcept('DELETE FROM MODELES WHERE MO_TYPE="E" AND MO_NATURE="TAR" AND MO_CODE IN ("TAR", "TAA")');
//  ExecuteSQLContOnExcept('DELETE FROM MODEDATA WHERE MD_CLE LIKE "ETARTAR%" OR MD_CLE LIKE "ETARTAA%"');
  //M MORRETTON 5622
  (*
  AglNettoieListesPlus('YTARIFSFSLART101', '(@LIBELLETIERS);(@LIBELLEARTICLE)');
  AglNettoieListesPlus('YTARIFSFSLART201', '(@LIBELLETIERS);(@LIBELLEARTICLE)');
  AglNettoieListesPlus('YTARIFSFSLART211', '(@LIBELLETIERS);(@LIBELLEARTICLE)');
  AglNettoieListesPlus('YTARIFSFSLART301', '(@LIBELLETIERS);(@LIBELLEARTICLE)');
  AglNettoieListesPlus('YTARIFSFSLART401', '(@LIBELLETIERS);(@LIBELLEARTICLE)');
  AglNettoieListesPlus('YTARIFSFSLART501', '(@LIBELLETIERS);(@LIBELLEARTICLE)');
  AglNettoieListesPlus('YTARIFSFSLART601', '(@LIBELLETIERS);(@LIBELLEARTICLE)');
  AglNettoieListesPlus('YTARIFSFSLTIE101', '(@LIBELLETIERS);(@LIBELLEARTICLE)');
  AglNettoieListesPlus('YTARIFSFSLTIE201', '(@LIBELLETIERS);(@LIBELLEARTICLE)');
  AglNettoieListesPlus('YTARIFSFSLTIE211', '(@LIBELLETIERS);(@LIBELLEARTICLE)');
  AglNettoieListesPlus('YTARIFSFSLTIE301', '(@LIBELLETIERS);(@LIBELLEARTICLE)');
  AglNettoieListesPlus('YTARIFSFSLTIE401', '(@LIBELLETIERS);(@LIBELLEARTICLE)');
  AglNettoieListesPlus('YTARIFSFSLTIE501', '(@LIBELLETIERS);(@LIBELLEARTICLE)');
  AglNettoieListesPlus('YTARIFSFSLTIE601', '(@LIBELLETIERS);(@LIBELLEARTICLE)');
  *)
End;

procedure MajVer966;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  // MNG 5636
  // appel fonction de mise à jour ytradmetier pour ajout expression
//  CRMMajYTradMetier;

  //JPL 5649
  (*
  ExecuteSQLNoPCL('UPDATE QBPSEQBUDGET '
             + 'SET QBD_SALARIE="", QBD_STATUTBUD="", QBD_DATESTATUT="'+UsDateTime_(IDate1900)+'", QBD_TRANSFERT="-", QBD_DATETRANSFERT="' + UsDateTime_(IDate1900) + '"'
             + ' WHERE QBD_SALARIE is null AND QBD_STATUTBUD is null AND QBD_DATESTATUT is null AND QBD_TRANSFERT is null AND QBD_DATETRANSFERT is null');
  *)
End;

procedure MajVer967;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //MCD 5691
//   InsertChoixCode('ATU', '007', 'Chiffre de missions', '', 'Chiffre d''affaires');
  //JTR 5697
   { JTR_FQ;010;17206 }
//  ExecuteSQLNoPCL('UPDATE TIERSIMPACTPIECE SET GTI_TYPENATTIERS = "FAC" WHERE GTI_ELEMENTFC = "TXN" AND GTI_TYPENATTIERS = ""');

End;

procedure MajVer968;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //JTR 5783
  { JTR - FQ;PGISIDE;10230}
  AglNettoieListes('PSE_MULEADRESSES','EAG_NUMEROADRESSE',nil);
  //N FOURNEL 5804
//  ExecuteSqlNoPCL ('UPDATE PARPIECE SET GPP_LIBELLETRANSFO = ("\""%0:s n° %1:d du " ||"\"""|| "dd/mm/yyyy" ||"\"""|| " %2:s"||"\""") WHERE GPP_LIBELLETRANSFO IS NULL');

End;

procedure MajVer969;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //RH 5811
  (*
  ExecuteSQLNoPCL('UPDATE HRCAISSE SET HRC_CATALOGUE="-", HRC_ARTICLECATA="", HRC_ARTICLECATA2="", HRC_ARTICLECATA3="", HRC_ARTICLECATA4="", HRC_ARTICLECATA5="" WHERE HRC_ARTICLECATA2 IS NULL');
  *)
  //MM 5855
//  ExecuteSQLNoPCL('DELETE FROM PARAMSOC WHERE SOC_NOM="SCO_PDRGEVALOENC"');

  //LEK 5867
  ExecuteSQLContOnExcept('update cpbonsapayer set bap_datevisa = "' + UsDateTime_(iDate1900) + '" where not bap_statutbap in ("VAL", "DEF")');
  ExecuteSQLContOnExcept('update cpbonsapayer set bap_datevisa = BAP_DATEMODIF where bap_statutbap in ("VAL", "DEF")');
  ExecuteSQLContOnExcept('update CPCIRCUIT set CCI_FERME = "-" WHERE CCI_FERME IS NULL');
  ExecuteSQLContOnExcept('update CPTYPEVISA set CTI_FERME = "-" WHERE CTI_FERME IS NULL');
  ExecuteSQLContOnExcept('update CPTYPEVISA set CTI_MAILPLUSUN = "-" WHERE CTI_MAILPLUSUN IS NULL');
End;


procedure MajVer970;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //MD 5921
  ExecuteSQLContOnExcept('DELETE FROM PARAMSOC WHERE SOC_NOM="SO_ACTIVEJEDECLARE"');
End;

procedure MajVer971;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //GM   attention SIC
//  ExecuteSQLNoPCL ('UPDATE AFFAIRE SET AFF_MODEAUGMEN="" ,AFF_TYPEAUGMEN="" WHERE AFF_MODEAUGMEN is Null');
  //LEK 6023
//  ExecuteSQLNoPCL ('update EEXBQLIG set CEL_NUMCETEBAC = 0 WHERE CEL_NUMCETEBAC IS NULL');
  ExecuteSQLNoPCL ('update GUIDE set GU_VALIDE = "-" WHERE GU_VALIDE IS NULL');

End;

procedure MajVer972;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //B DUCLOS 5918
//  AglNettoieListesPlus('AFMULFACTIERSAFF', 'T_REGIMETVA', nil, true);
//  AglNettoieListesPlus('AFMULFACTIERSAFFP', 'T_REGIMETVA', nil, true);

  //T SUBLET 6055
    //GP_20091008_TS_GP16663
//  AglNettoieListes('WCBNPREVTET', 'WPT_CODEMESSAGE;WPT_AFFAIRE', nil);
//  AglNettoieListes('WCBNPREVLIG', 'WPD_CODEPREVISION;WPD_ARTICLE;WPD_DEPOT', nil);

  // CD 6056
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE = "Contrôle de capacité" WHERE dh_prefixe = "GEM" and DH_NOMCHAMP = "GEM_CAPACITE"');

  //6060
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE = "Prix revient : Valorisation/Qté éco" WHERE dh_prefixe = "GA" and DH_NOMCHAMP = "GA_VALOQTEECOPDR"');

  // JPL 6081
  // Maj lignes des DA avec Articles
  (*
  ExecuteSQLNoPCL('UPDATE LIGNEDA SET ' +
  ' DAL_LIBREART1= (SELECT GA_LIBREART1 FROM ARTICLE WHERE GA_ARTICLE=DAL_ARTICLE), ' +
  ' DAL_LIBREART2 =(SELECT GA_LIBREART2 FROM ARTICLE WHERE GA_ARTICLE=DAL_ARTICLE), ' +
  ' DAL_LIBREART3 =(SELECT GA_LIBREART3 FROM ARTICLE WHERE GA_ARTICLE=DAL_ARTICLE), ' +
  ' DAL_LIBREART4 =(SELECT GA_LIBREART4 FROM ARTICLE WHERE GA_ARTICLE=DAL_ARTICLE), ' +
  ' DAL_LIBREART5=(SELECT GA_LIBREART5 FROM ARTICLE WHERE GA_ARTICLE=DAL_ARTICLE),  ' +
  ' DAL_LIBREART6 =(SELECT GA_LIBREART6 FROM ARTICLE WHERE GA_ARTICLE=DAL_ARTICLE), ' +
  ' DAL_LIBREART7=(SELECT GA_LIBREART7 FROM ARTICLE WHERE GA_ARTICLE=DAL_ARTICLE), ' +
  ' DAL_LIBREART8=(SELECT GA_LIBREART8 FROM ARTICLE WHERE GA_ARTICLE=DAL_ARTICLE), ' +
  ' DAL_LIBREART9=(SELECT GA_LIBREART9 FROM ARTICLE WHERE GA_ARTICLE=DAL_ARTICLE), ' +
  ' DAL_LIBREARTA=(SELECT GA_LIBREARTA FROM ARTICLE WHERE GA_ARTICLE=DAL_ARTICLE), ' +
  ' DAL_FAMILLENIV1=(SELECT GA_FAMILLENIV1 FROM ARTICLE WHERE GA_ARTICLE=DAL_ARTICLE), ' +
  ' DAL_FAMILLENIV2=(SELECT GA_FAMILLENIV2 FROM ARTICLE WHERE GA_ARTICLE=DAL_ARTICLE), ' +
  ' DAL_FAMILLENIV3=(SELECT GA_FAMILLENIV3 FROM ARTICLE WHERE GA_ARTICLE=DAL_ARTICLE) ' +
  ' where (DAL_LIBREART1 is null) and (DAL_LIBREART2 is NULL) and (DAL_LIBREART3 Is Null) ' +
  ' and (DAL_LIBREART4 is null) and (DAL_LIBREART5 is NULL) and (DAL_LIBREART6 Is Null) ' +
  ' and (DAL_LIBREART7 is null) and (DAL_LIBREART8 is NULL) and (DAL_LIBREART9 Is Null) and (DAL_LIBREARTA is Null) ' +
  ' and (DAL_FAMILLENIV1 is null) and (DAL_FAMILLENIV2 is NULL) and (DAL_FAMILLENIV3 Is Null) and (DAL_ARTICLE <> "")');
  ExecuteSQLNoPCL('UPDATE LIGNEDA SET ' +
  ' DAL_LIBREART1= "", ' +
  ' DAL_LIBREART2 ="", ' +
  ' DAL_LIBREART3 ="", ' +
  ' DAL_LIBREART4 ="", ' +
  ' DAL_LIBREART5="",  ' +
  ' DAL_LIBREART6 ="", ' +
  ' DAL_LIBREART7="", ' +
  ' DAL_LIBREART8="", ' +
  ' DAL_LIBREART9="", ' +
  ' DAL_LIBREARTA="", ' +
  ' DAL_FAMILLENIV1="", ' +
  ' DAL_FAMILLENIV2="", ' +
  ' DAL_FAMILLENIV3="" ' +
  ' where (DAL_LIBREART1 is null) and (DAL_LIBREART2 is NULL) and (DAL_LIBREART3 Is Null) ' +
  ' and (DAL_LIBREART4 is null) and (DAL_LIBREART5 is NULL) and (DAL_LIBREART6 Is Null) ' +
  ' and (DAL_LIBREART7 is null) and (DAL_LIBREART8 is NULL) and (DAL_LIBREART9 Is Null) and (DAL_LIBREARTA is Null) ' +
  ' and (DAL_FAMILLENIV1 is null) and (DAL_FAMILLENIV2 is NULL) and (DAL_FAMILLENIV3 Is Null) and (DAL_ARTICLE="")');

  // Maj lignes des DA avec Contrôle budgétaire

  ExecuteSQLNoPCL ('UPDATE LIGNEDA SET ' +
  ' DAL_CTRLBUDGET = (SELECT DA_CTRLBUDGET FROM PIECEDA WHERE DA_TYPEDA=DAL_TYPEDA AND DA_NUMERO=DAL_NUMERO AND DA_SOUCHE=DAL_SOUCHE) where DAL_CTRLBUDGET is null');
*)
  // Maj du circuit avec l'axe de validation

//  ExecuteSQLNoPCL ('UPDATE CIRCUITDA SET DAV_CODEAXE="" WHERE DAV_CODEAXE IS NULL');

  // Maj de la DA et de histo de la DA avec elements de hiérarchie
(*
  ExecuteSQLNoPCL ('UPDATE HISTODA SET DAH_CODEHIERAR="", DAH_AXEPROCHAINVAL="", DAH_VALPROCHAINVAL="" Where (DAH_CODEHIERAR is null) and ' +
  ' (DAH_AXEPROCHAINVAL is null) and (DAH_VALPROCHAINVAL is null) ');


  ExecuteSQLNoPCL ('UPDATE PIECEDA SET DA_CODEHIERAR="", DA_AXELIENHIERAR="", DA_VALHIERARCHIE="", DA_AXEPROCHAINVAL="", DA_VALPROCHAINVAL="" ' +
  ' Where (DA_CODEHIERAR is null) and (DA_AXELIENHIERAR is null) and (DA_VALHIERARCHIE is null) and (DA_AXEPROCHAINVAL is null) and (DA_VALPROCHAINVAL is null) ');

  // Maj Typeda avec les codes sessions du budget

  ExecuteSQLNoPCL ('UPDATE TYPEDA SET DAT_TYPEHIERAR = "SER", DAT_REVISION="0", DAT_REVISION1="0", ' +
  ' DAT_CODESESSION = (SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_GCCODESESSIONN"), ' +
  ' DAT_CODESESSION1 = (SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_GCCODESESSIONN1"), ' +
  ' DAT_CODEHIERAR = "" WHERE (DAT_TYPEHIERAR is null) and (DAT_REVISION is null) and (DAT_REVISION1 is null) ' +
  ' and (DAT_CODESESSION is null) and (DAT_CODESESSION1 is null)  and (DAT_CODEHIERAR is null)  ');

  // Maj des lignes de pieces pour le contrôle budgétaire

  ExecuteSQLNoPCL ('UPDATE LIGNECOMPL SET GLC_AXELIENHIERAR="", GLC_VALHIERARCHIE="", GLC_CODESESSION="", GLC_DAPRECEDENTE="" ' +
  ' Where (GLC_AXELIENHIERAR is null) and (GLC_VALHIERARCHIE is null) and (GLC_CODESESSION is null) and (GLC_DAPRECEDENTE is null) ');
  *)
  // Maj des enreg. du budget existant

//  ExecuteSQLNoPCL ('UPDATE QBPARBRE SET QBR_REVISION = "0" Where QBR_REVISION is null');


//  ExecuteSQLNoPCL ('UPDATE QBPARBREDETAIL SET QBH_REVISION = "0" Where QBH_REVISION is null ');

//  ExecuteSQLNoPCL ('UPDATE QBPDETCALENDREP SET QBE_REVISION = "0" Where QBE_REVISION is null ');

//  ExecuteSQLNoPCL ('UPDATE QBPSESSIONBP SET QBS_CODEHIERAR = "" Where QBS_CODEHIERAR is null ');

  // Maj info engagement dans la session budgétaire

//  ExecuteSQLNoPCL ('UPDATE QBPSESSIONBP SET QBS_CODESTRUCT="#PGIENGAGE" WHERE (QBS_CODESESSION = (SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_GCCODESESSIONN")) or ' +
//  ' (QBS_CODESESSION = (SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_GCCODESESSIONN1"))  or (QBS_CODESESSION in (SELECT DISTINCT DA_CODESESSION FROM PIECEDA WHERE DA_CODESESSION <> "")) ');

  // Maj noueau champ
//  ExecuteSQLNoPCL ('UPDATE QBPBIBLIOAXE SET QBX_NOMCHPPR ="", QBX_LEFTJOINRPR = "" Where (QBX_NOMCHPPR is null) and (QBX_LEFTJOINRPR is null)' );

  // Maj nouvelle table révisions

//  ExecuteSQLNoPCL ('INSERT INTO QBPREVISION (QRB_CODESESSION, QRB_REVISION, QRB_LIBELLE, QRB_DATECREATION, QRB_DATEMODIF, QRB_CREATEUR, QRB_UTILISATEUR, QRB_SOCIETE) ' +
//  ' ( SELECT DISTINCT QBR_CODESESSION,"0","Révision 0","' + UsDateTime_(V_PGI.DateEntree)+ '","' + UsDateTime_(V_PGI.DateEntree) + '","CEG","CEG","001" FROM QBPARBRE '+
//  ' WHERE NOT exists (SELECT QRB_REVISION FROM QBPREVISION WHERE QRB_CODESESSION=QBR_CODESESSION)) ');

  //D PERRAUD 6098
  //MajApres
//  ExecuteSqlNoPCL('UPDATE QAPPROS SET QA_TIERS="", QA_AFFAIRE="" WHERE QA_AFFAIRE IS NULL');
//  ExecuteSqlNoPCL('UPDATE QTRANDETD SET QTD_TIERS="", QTD_AFFAIRE="" WHERE QTD_AFFAIRE IS NULL');

  // D SCLAVOPOULOS 6114
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Qté à réceptionner prévue (saisie)" WHERE DH_NOMCHAMP IN("WOL_QARECSAIS","WOP_QARECSAIS")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE="Qté à réceptionner prévue (stock)" WHERE DH_NOMCHAMP IN("WOL_QARECSTOC","WOP_QARECSTOC")');

 { GC_20080729_JTR_01015435_Début
 // déplacé du majver917 à ici car mal traité avec updatedecoupe...
 La requête était exécutée en 917, mais pouvait "planter". Ajout de "MAX(" dans les SELECT sur la table ligne pour n'en ramener qu'un seul }
{ //JS1 17122008 on passe dazns tous les cas par le bloc DB2 if V_PGI.Driver in [dbDB2] then
  begin}
  (*
  if not IsDossierPCL then
  begin
    // CRM_20091223_MNG_FQ;010;18532
    UpDateDecoupePiece ('GP_COMPTATIERS ='
                   + ' ISNULL((SELECT ##TOP 1## T_COMPTATIERS FROM TIERS WHERE T_TIERS = GP_TIERS ORDER BY T_DATEDERNMVT DESC), ""), GP_DATECOMPTA=GP_DATEPIECE', '');
  end;
  *)
End;

procedure MajVer973;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //N FOURNEL 6117
//  ExecuteSQLContOnExcept('DELETE FROM PARAMSOC WHERE SOC_NOM = "SO_GCDELAILIVRAISON"');
  //D SCLAVOPOULOS 6140
  ExecuteSQLContOnExcept('UPDATE COMMUN SET CO_LIBELLE="Depuis les phases de production" WHERE CO_TYPE="WRS" AND CO_CODE="WOP"');
  //C DUMAS 6208
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE = "Date des cumuls non renseignée" WHERE dh_prefixe = "EMG" and DH_NOMCHAMP = "EMG_BLODATECUMUL"');

//  AglNettoieListes('AFTACHE_MUL', 'ATA_NATUREPIECEG;ATA_NUMERO', nil);

End;

procedure MajVer974;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //R HARANG 6261
//  ExecuteSQLNoPCL('UPDATE HRPARAMEMAIL SET HPE_ETABLISSEMENT="'+ GetParamSocSecur ('SO_ETABLISDEFAUT','')+'"');
//  ExecuteSQLNoPCL('UPDATE HREVENEMENT SET HEV_ETABLISSEMENT="'+ GetParamSocSecur ('SO_ETABLISDEFAUT','')+'"');
End;

procedure MajVer975;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //MNG 6302
  // appel fonction de mise à jour ytradmetier suite à suppression expression
//  ExecuteSQLNoPcL ('UPDATE WPARAMFONCTION SET WPF_BOOLEAN01="X" WHERE WPF_CODEFONCTION LIKE "QTEECO%"');
end;

procedure MajVer976;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //C PARWEZ : VALIDE PAR MAIL PAR JLS
  ExecuteSQLNoPCL( 'UPDATE DECHAMPS SET DH_LIBELLE = "Coef. production" WHERE DH_PREFIXE = "WOR" AND DH_NOMCHAMP = "WOR_COEFPROD" AND DH_LIBELLE <> "Coef. production"');
  //C DUMAS 6410
  if not IsDossierPCL then
  begin
      InsertChoixExt('GCX', 'TGCA_DATELIBRE3' , 'Date libre 3' , 'Date libre 3' , '');
      InsertChoixExt('GCW', 'TGCA_VALLIBRE3' , 'Valeur libre 3' , 'Valeur libre 3' , '');
      InsertChoixExt('GCY', 'TGCA_BOOLLIBRE3' , 'Décision libre 3' , 'Décision libre 3' , '');
      InsertChoixExt('GCZ', 'TGCA_CHARLIBRE3' , 'Texte libre 3' , 'Texte libre 3' , '');
  end;
  //C PARWEZ 6422
  ExecuteSQLNoPCL( 'DELETE PARAMSOC WHERE SOC_NOM = "SCO_WSSOIMPACTEWOL"');

  //T PETETIN 6435
//  ExecuteSQLNoPCL( 'UPDATE LIGNE SET GL_ECARTUNI=0 WHERE (GL_ECARTUNI IS NULL)' );
//  UpDateDecoupeLigne ('GL_ECARTUNI=0', ' AND (GL_ECARTUNI IS NULL)');

  // D SCLAVOPOULOS 6456
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Q" WHERE DH_PREFIXE ="WOP" AND DH_CONTROLE NOT LIKE "%Q%"'
                      + ' AND (DH_NOMCHAMP="WOP_QACCSAIS" OR DH_NOMCHAMP="WOP_QACCSTOC" OR DH_NOMCHAMP="WOP_QLANSAIS" OR DH_NOMCHAMP="WOP_QLANSTOC"'
                      + ' OR DH_NOMCHAMP="WOP_QARECSAIS" OR DH_NOMCHAMP="WOP_QARECSTOC" OR DH_NOMCHAMP="WOP_QRECSAIS" OR DH_NOMCHAMP="WOP_QRECSTOC"'
                      + ' OR DH_NOMCHAMP="WOP_QSUSPSAIS" OR DH_NOMCHAMP="WOP_QSUSPSTOC" OR DH_NOMCHAMP="WOP_QREBUSAIS" OR DH_NOMCHAMP="WOP_QREBUSTOC")');

//  if not IsDossierPCL() then
//    EdiMajVer(976).MajApres();
end;

procedure MajVer977;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //L CHHOEU 6494
  InsertChoixExt ('CCJ','1','Janvier','','');
  InsertChoixExt ('CCJ','2','Février','','');
  InsertChoixExt ('CCJ','3','Mars','','');
  InsertChoixExt ('CCJ','4','Avril','','');
  InsertChoixExt ('CCJ','5','Mai','','');
  InsertChoixExt ('CCJ','6','Juin','','');
  InsertChoixExt ('CCJ','7','Juillet','','');
  InsertChoixExt ('CCJ','8','Août','','');
  InsertChoixExt ('CCJ','9','Septembre','','');
  InsertChoixExt ('CCJ','10','Octobre','','');
  InsertChoixExt ('CCJ','11','Novembre','','');
  InsertChoixExt ('CCJ','12','Décembre','','');
  InsertChoixExt ('CCJ','A','Clôture par période','','');
  InsertChoixExt ('CCJ','B','Clôture par journal','','');
  InsertChoixExt ('CCJ','C','Déclôture par période','','');
  InsertChoixExt ('CCJ','D','Déclôture par journal','','');

  // M MORRETTON 6464
(*
  ExecuteSqlNoPCL('UPDATE COUTSINDIRECTS SET'
  +' CI_MODEGROUPEPORT=(SELECT GPO_MODEGROUPEPORT FROM PORT WHERE GPO_CODEPORT=CI_CODEPORT)'
  +',CI_REPARTITION=(SELECT GPO_REPARTITION FROM PORT WHERE GPO_CODEPORT=CI_CODEPORT)'
  +',CI_TYPEPORT=(SELECT GPO_TYPEPORT FROM PORT WHERE GPO_CODEPORT=CI_CODEPORT)'
  +',CI_TYPEFRAIS=(SELECT GPO_TYPEFRAIS FROM PORT WHERE GPO_CODEPORT=CI_CODEPORT)'
  +',CI_COEFTYPEDONNEE=(SELECT YFO_COEFTYPEDONNEE FROM YTARIFSPARAMETRES WHERE YFO_FONCTIONNALITE=CI_FONCTIONNALITE AND YFO_PROFILCOMPO="B01" AND YFO_ORIENTATION="TIE" AND YFO_CODEPORT=CI_CODEPORT)'
  +',CI_COEFSENSDONNEE=(SELECT YFO_COEFSENSDONNEE FROM YTARIFSPARAMETRES WHERE YFO_FONCTIONNALITE=CI_FONCTIONNALITE AND YFO_PROFILCOMPO="B01" AND YFO_ORIENTATION="TIE" AND YFO_CODEPORT=CI_CODEPORT)'
  +',CI_COEFSENSCALCUL=(SELECT YFO_COEFSENSCALCUL FROM YTARIFSPARAMETRES WHERE YFO_FONCTIONNALITE=CI_FONCTIONNALITE AND YFO_PROFILCOMPO="B01" AND YFO_ORIENTATION="TIE" AND YFO_CODEPORT=CI_CODEPORT)'
  +' WHERE CI_MODEGROUPEPORT IS NULL');
*)
  //D SCLAVOPOULOS 6461
//  AglNettoieListes('WORDREBESDEM', 'WOB_COEFLIEN');

End;

procedure MajVer978;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //N FOURNEL
//  ExecuteSQLNoPCL('UPDATE ARTICLECOMPL SET GA2_TOTALAUTO = "TC1;" WHERE GA2_TOTALAUTO = "MT;"');
  //N FOURNEL 6543
//  ExecuteSQLContOnExcept('DELETE FROM PARAMSOC WHERE SOC_NOM = "SO_GCDELAILIVRAISON"');

//  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "C" WHERE DH_PREFIXE ="WJA" AND DH_CONTROLE NOT LIKE "%C%"');
  //N FOURNEL
//  ExecuteSqlNoPCL ('UPDATE PARPIECE SET GPP_LIBELLETRANSFO = ("\""%0:s n° %1:d du " ||"\"""|| "dd/mm/yyyy" ||"\"""|| " %2:s"||"\""") WHERE GPP_LIBELLETRANSFO = ""');
end;

procedure MajVer979;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //M MORRETTON
//  ExecuteSQLNoPCL('UPDATE PARAMSOC SET SOC_DATA="IEC" WHERE SOC_NOM LIKE "SO_INVQTEVAL" AND SOC_DATA=""');
  //MNG
    // init oublié nouveau champ GTP_DOMAINE (socref 925)
//  ExecuteSQLNoPCL ('UPDATE TIERSPIECE SET GTP_DOMAINE="" WHERE GTP_DOMAINE IS NULL');
  //D KOZA
//  ExecuteSQLNoPCL('UPDATE WAFREVISIONBUDGET SET WRB_WBMEMO="-",WRB_BLOCNOTE="" WHERE WRB_WBMEMO IS NULL');
end;

procedure MajVer981;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //JP LAURENT
//  MajEngagement; ------> nannn domaine Q

  //D BUCHET
//  AjouteNewTarifs('221'); //GA_200912_AB_FQ;009;16739 paramètre nouveau tarif

  //D SCLAVOPOULOS
//  AglNettoieListes('WORDREPHASE','WOP_POLE;WOP_GRP');
//  AglNettoieListesPlus('WORDREPHASE','WOP_UNITEACC', nil, True);

//  ExecuteSQLContOnExcept( 'DELETE FROM PARAMSOC WHERE SOC_NOM="SO_WMESSAVANTIMPACTQTENEXTWOP"');

//  AGLNettoieListesPlus('WORDREBESL', 'WOB_QAFFSAIS', nil, True);

  //M MORRETTON
//  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_CONTROLE=DH_CONTROLE||"C"'
//  +'WHERE DH_PREFIXE="WPL"'
 // +' AND DH_NOMCHAMP IN ("WPL_VALEURPDRENT", "WPL_VALEURPDRSOR", "WPL_QTEENTREE", "WPL_COUTENTREE",'
//  +' "WPL_QTEPHASE", "WPL_COUTPHASE", "WPL_QTESORTIE", "WPL_COUTSORTIE")'
//  +' AND (NOT DH_CONTROLE LIKE "%C%") ');

  //C DUMAS
  ExecuteSQLNoPCL('update DETABLES set dt_libelle = "Paramétrage de la fidélité client" where dt_prefixe = "GFO"');

  //JP LAURENT
  //  Tables Libres pour la hiérarchie
  InsertChoixCode ('ZLH', 'HT1' ,'.- Table libre 1' , '','');
  InsertChoixCode ('ZLH', 'HT2' ,'.- Table libre 2' , '','');
  InsertChoixCode ('ZLH', 'HT3' ,'.- Table libre 3' , '','');
  InsertChoixCode ('ZLH', 'HT4' ,'.- Table libre 4' , '','');
  InsertChoixCode ('ZLH', 'HT5' ,'.- Table libre 5' , '','');
  InsertChoixCode ('ZLH', 'HT6' ,'.- Table libre 6' , '','');
  InsertChoixCode ('ZLH', 'HT7' ,'.- Table libre 7' , '','');
  InsertChoixCode ('ZLH', 'HT8' ,'.- Table libre 8' , '','');
  InsertChoixCode ('ZLH', 'HT9' ,'.- Table libre 9' , '','');
  InsertChoixCode ('ZLH', 'HTA' ,'.- Table libre 10' , '','');
End;

procedure MajVer982;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //N FOURNEL
  //ExecuteSqlNoPCL('UPDATE PARPIECE SET GPP_GESTARTFERME="CRE;DUP;" WHERE GPP_GESTARTFERME IS NULL');
  //ExecuteSqlNoPCL('UPDATE PARPIECE SET GPP_GESTDATEFERMEE="MES" WHERE GPP_GESTDATEFERMEE IS NULL');

  //LEK
  ExecuteSqlNoPCL('UPDATE GENERAUX SET G_TYPECPTTVA="DIV" WHERE G_NATUREGENE="DIV" OR G_NATUREGENE IS NULL OR G_TYPECPTTVA IS NULL');
  ExecuteSqlNoPCL('UPDATE GENERAUX SET G_TYPECPTTVA="TVA" WHERE G_GENERAL IN (SELECT TV_CPTEACH FROM TXCPTTVA WHERE TV_CPTEACH=G_GENERAL)');
  ExecuteSqlNoPCL('UPDATE GENERAUX SET G_TYPECPTTVA="TVA" WHERE G_GENERAL IN (SELECT TV_CPTEVTE FROM TXCPTTVA WHERE TV_CPTEVTE=G_GENERAL)');
  ExecuteSqlNoPCL('UPDATE GENERAUX SET G_TYPECPTTVA="TVA" WHERE G_GENERAL IN (SELECT TV_ENCAISACH FROM TXCPTTVA WHERE TV_ENCAISACH=G_GENERAL)');
  ExecuteSqlNoPCL('UPDATE GENERAUX SET G_TYPECPTTVA="TVA" WHERE G_GENERAL IN (SELECT TV_ENCAISVTE FROM TXCPTTVA WHERE TV_ENCAISVTE=G_GENERAL)');
  ExecuteSqlNoPCL('UPDATE GENERAUX SET G_TYPECPTTVA="HT"  WHERE G_NATUREGENE IN ("CHA","PRO","IMO")');
  ExecuteSqlNoPCL('UPDATE GENERAUX SET G_TYPECPTTVA="TTC" WHERE G_NATUREGENE IN ("COC","COD","COF","COS","TIC","TID")');

  //D SCLAVOPOULOS
//  AglNettoieListes('WLIGNESCSP','GL_GUIDWOL', nil,'GL_IDENTIFIANTWOL');

  //M MORRETTON
(*
  ExecuteSQLNoPCL('UPDATE PARAMSOC SET SOC_DATA="A"'
  +' WHERE SOC_NOM="SO_PDRORRQTE" AND (SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_PDRORRORIGINE")="P"'
  +' OR SOC_NOM="SO_PDRORRQTECDE" AND (SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_PDRORRORIGINECDE")="P"'
  +' OR SOC_NOM="SO_PDRORRQTESTRACH" AND (SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_PDRORRORIGINECDE")="P"');
  ExecuteSQLNoPCL('UPDATE PARAMSOC SET SOC_DATA="O" WHERE SOC_NOM="SO_PDRORRORIGINE" OR SOC_NOM="SO_PDRORRORIGINECDE"'
  +' OR SOC_NOM="SO_PDRORRORIGINESTRACH"');
*)

  //B DUCLOS
//  AglNettoieListesPlus('AFMULAFFAIREMULTI', 'AFF_TYPEAFFAIRE', nil, true);

End;

procedure MajVer983;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //XP
  ExecuteSQLContOnExcept('UPDATE UTILISAT SET US_INTERACTIF="X" WHERE US_INTERACTIF IS NULL OR (US_INTERACTIF NOT IN ("-", "X"))');

  //SAULNIER
  (*
  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Préfixe du code à barres" where dh_nomchamp="GCB_PREFIXECAB"');
  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Code de la série" where dh_nomchamp="GAM_CODESERIE"');
  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Colonnes à déstocker" where dh_nomchamp="GPP_QTEMOINS"');
  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Comportement dimensions en saisie",dh_explication="Comportement dimensions en saisie" where dh_nomchamp="GPP_DIMSAISIE"');
  ExecuteSQLContOnExcept('update dechamps set dh_explication="Comportement anal stock ligne" where dh_nomchamp="GPP_COMPSTOCKLIGNE"');
  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Initialisation qté en création" where dh_nomchamp="GPP_INITQTECRE"');
  ExecuteSQLContOnExcept('update detables set dt_libelle="Commission représentant" where dt_nomtable="COMMISSION"');
  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Constante prix psychologique",dh_explication="Constante prix psychologique" where dh_nomchamp="GAR_CONSTANTE"');
  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Créateur de la fiche" where dh_nomchamp="GZQ_CREATEUR" or dh_nomchamp="WND_CREATEUR"');
  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Délai moyen de livraison" where dh_nomchamp="GAT_DELAIMOYEN"');
  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Identifiant fournisseur choisi" where dh_nomchamp="GZC_TIERS"');
  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Intitulé",dh_explication="Intitulé" where dh_nomchamp="GDP_LIBELLE"');
  ExecuteSQLNoPCL ('UPDATE PARPIECE SET GPP_CONTEXTES = GPP_CONTEXTES || "INU;" WHERE GPP_NATUREPIECEG="CCH"');
  *)
  //JLS Le 24/12/2009
//  ExecuteSqlNoPCL('UPDATE WORDREPHASE SET WOP_QREBPREVSAIS=0, WOP_QREBPREVSTOC=0 WHERE WOP_QREBPREVSAIS IS NULL');

  //T PETETIN
  (*
  If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "YEB" AND CO_CODE = "W16"') then
     ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) '
     +'values("YEB", "W16", "WFORMCONVVARDEF", "TAB", "G01")');
  If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "YEB" AND CO_CODE = "W17"') then
     ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) '
     +'values("YEB", "W17", "WFORMCONV", "TAB", "G01")');
  If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "YEB" AND CO_CODE = "W18"') then
     ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) '
     +'values("YEB", "W18", "WFORMCONVVAR", "TAB", "G01")');
  If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "YEB" AND CO_CODE = "W19"') then
     ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) '
     +'values("YEB", "W19", "WFORMCONV", "TTE", "G01")');
  If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "YEB" AND CO_CODE = "W20"') then
     ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) '
     +'values("YEB", "W20", "WFORMCONVGA", "TTE", "G01")');
  If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "YEB" AND CO_CODE = "W21"') then
     ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) '
     +'values("YEB", "W21", "WFORMCONVVAR", "TTE", "G01")');
  *)
//   AglNettoieListes('GCGROUPEPIECEACH', 'GP_DOMAINE');

   { JTR - FQ;010;17230 }
   ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA = "TCH" WHERE SOC_NOM = "SO_TYPETAUXDEVISE" AND SOC_DATA = ""');

End;

procedure MajVer984;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  //TS GP_20100105_TS_GP16927
//  if not IsDossierPCL() then
//    EdiMajVer(984).MajApres();
 //D SCLAVOPOULOS
//   AglNettoieListes('WORDREBES2', 'QAFF;WOB_QBESSTOC;WOB_QCONSTOC', nil);
end;

procedure MajVer985;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  (*
  ExecuteSQLNoPCL('UPDATE PARPIECE SET GPP_LISTEAFFAIRE = "AFSAISIEFAC" WHERE GPP_NATUREPIECEG IN ("AVA", "FAA") '
    + 'AND ((GPP_LISTEAFFAIRE = "") OR (GPP_LISTEAFFAIRE IS NULL))');
  *)
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE||"C" WHERE '
  +'(DH_NOMCHAMP="WPE_DATEMODIF" OR DH_NOMCHAMP="WPE_UTILISATEUR") AND (NOT DH_CONTROLE LIKE "%C%")');
  ExecuteSQLNoPCL('DELETE FROM MENU WHERE MN_TAG LIKE "21186%"');
  ExecuteSQLNoPCL('UPDATE MENU SET MN_TAG=211861 WHERE MN_TAG="-21186"');
//  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Famille de taxe à appliquer" where dh_nomchamp="GEE_FAMILLETAXE"');
(*
  ExecuteSQLContOnExcept('update DECHAMPS set dh_libelle="Soumis à franco" where dh_nomchamp="GPT_FRANCO"');
  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Séparateur de champs" where dh_nomchamp="grf_separateur"');
  ExecuteSQLContOnExcept('update DECHAMPS set dh_libelle="Numéro ligne lu" where dh_nomchamp="ELI_NUMLIGNELU"');
  ExecuteSQLContOnExcept('update DECHAMPS set dh_libelle="Numéro de ligne" where dh_nomchamp="gzg_compteur"');
  ExecuteSQLContOnExcept('update dechamps set dh_libelle = "Nbre de tickets en attente détruits" where dh_nomchamp = "GJC_NBTICATTSUP"');
  ExecuteSQLContOnExcept('update DECHAMPS set dh_libelle ="Montant minimum pour validation" where dh_nomchamp="DAT_MONTANTMINI"');
  ExecuteSQLContOnExcept('update detables set DT_LIBELLE = "Ligne des numéros de série"  where dt_nomtable="LIGNESERIE"');
  ExecuteSQLContOnExcept('update DECHAMPS set dh_libelle="Laisser un pourcentage du stock" where dh_nomchamp="GTE_TXSTOCK"');
*)
      // GA_201001_JVO_GA15581
  (*
  if not IsDossierPCL() then
  begin
    if ExisteSQL('SELECT 1 FROM ARTICLE INNER JOIN ARTICLECOMPL ON (GA_ARTICLE=GA2_ARTICLE)' +
                 ' WHERE GA_TYPEARTICLE="PRE" AND (GA2_TYPEPLANIF="FOR" OR GA2_TYPEPLANIF="UNI")' +
                 ' AND (GA_QUALIFUNITEVTE <> GA_QUALIFUNITEACT) AND GA_QUALIFUNITEVTE <> ""') then
          SetParamSoc('SO_AFPBUNITEACTVTE', 'X')
    else
          SetParamSoc('SO_AFPBUNITEACTVTE', '-');
  end;
  *)
End;

procedure MajVer987;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;

  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "type nomenclature" WHERE DH_LIBELLE = "type nomemclature"');
//  ExecuteSQLContOnExcept('DELETE FROM PARAMSOC WHERE SOC_NOM="SO_GCCATALOGUEVENTE"');

end;

procedure MajVer988;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
    // JS on supprime les cptx / bobs arrivés par erreur suite à moulinage via majver
  ExecuteSQLContOnExcept('delete from ymybobs where yb_cptx in (select yc_guid '+
      'from ymycptx where yc_filename not like "%SOCREF%")');
  ExecuteSQLContOnExcept('delete from ymycptx where yc_filename not like "%SOCREF%"');

End;

procedure MajVer990;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  (*
  if not IsDossierPCL then
  begin
    UpDateDecoupeLigne ('GLC_MODFACTECHE="-", GLC_NUMFACTECHE=0, GLC_COMPTAARTICLE=""', 'AND GLC_MODFACTECHE is Null', 'GLC');
  end;
  *)
//  ExecutesqlNoPCL ('UPDATE ARTICLE SET GA_QPCBCONSO=0, GA_FAMILLECO = "" WHERE GA_QPCBCONSO IS NULL');
  // MC DESSEIGNET 3499
  //MCD 4479
   //GIGA CB Le typeplanif est maintenant obligatoire si articlecompl exist
  { GC_20090715_JTR_FQ;PGISIDE;10237 }
//  ExecuteSQLNoPCL ('UPDATE ARTICLECOMPL SET GA2_RESSOURCE="", GA2_CODEIMPORT = "" WHERE GA2_RESSOURCE IS NULL ');

  //M C DESSEIGNET le 11/06/2008 Version 9.0.912.3 Demande n°2686
  //A faire dans toutes les bases. cette table fait partie aussi des bases réduites
  ExecuteSQLContOnExcept('update contact set c_guidperanl ="", c_numeroadresse=0 where c_guidperanl is null');

  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Nombre d''unités d''éléments" '
  +'WHERE DH_LIBELLE = "Nombre d''unité d''éléments"');
//  ExecuteSQLContOnExcept('DELETE FROM PARAMSOC WHERE SOC_NOM = "SCO_GBCFG"');
  // JTR le 01/03/2010 suite FQ;010;18720
  ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA = "-" WHERE SOC_NOM = "SO_CPTASTKEXTPCEPREC"');

End;

procedure MajVer995;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  // D SCLAVOPOULOS
//  ExecuteSQLContOnExcept( 'DELETE FROM PARAMSOC WHERE SOC_NOM="SO_WMESSAVANTIMPACTQTENEXTWOP"');

  // P COAN
//  ExecuteSQLContOnExcept( 'UPDATE DECHAMPS SET DH_LIBELLE = "Qui fournit la matière" WHERE DH_PREFIXE = "WOB" AND DH_NOMCHAMP = "WOB_QUIFOURNI" AND DH_LIBELLE = "Qui fourni la matière"');

  // T PETETIN
  // FQ;034;16804 (Suite demande 6810)
(*
  If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "YEB" AND CO_LIBRE = "Y01" AND CO_CODE = "GP8"') then
    ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) values("YEB", "GP8", "SO_GCGESTUNITEMODE", "PAR", "Y01")');
*)
  //M MORRETTON
//  ExecuteSQLContOnExcept('DELETE FROM MODEDATA WHERE MD_CLE LIKE "ETARTAR%"');
//  ExecuteSQLContOnExcept('DELETE FROM MODELES WHERE MO_NATURE="TAR" AND MO_CODE="TAR"');

  //MN GARNIER
  ExecuteSQLContOnExcept('DELETE FROM MODELES WHERE MO_TYPE = "E" AND MO_NATURE = "AEC" AND MO_CODE = "EC1" ');
  ExecuteSQLContOnExcept('DELETE FROM MODEDATA WHERE MD_CLE LIKE "EAECEC1%"  ');

  //J TRIFILIEFF
  ExecuteSQLNoPCL('UPDATE ETABLES SET EDT_USEQUERY = "-" WHERE (EDT_USEQUERY <> "X") OR (EDT_USEQUERY IS NULL)');

  //M DEMEULEMEESTER
  ExecuteSQLContOnExcept('UPDATE dechamps SET DH_LIBELLE = "Désactive le contrôle quantité" WHERE DH_LIBELLE="Désactive le controle quantité"');

  //P COAN
//  ExecuteSQLNoPcl('UPDATE WCFGCHOIXTET SET WXT_TYPEAFFICHAGE = "CHK" WHERE WXT_TYPEAFFICHAGE IS NULL');

  //M MORRETTON
  ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_CONTROLE=DH_CONTROLE||"C" WHERE DH_PREFIXE="WPL" AND (NOT DH_CONTROLE LIKE "%C%") ');

//  ExecuteSQLNoPCL('UPDATE YTARIFSTYPE SET YTV_PRIORITEQUIM="", YTV_PRIORITEDATEM="'+UsDateTime_(IDate1900)+'" WHERE YTV_PRIORITEQUIM IS NULL');

//  ExecuteSQLNoPCL('UPDATE PORT SET GPO_FRAISINDPR="SOC" WHERE GPO_FRAISINDPR=""');
(*
  ExecuteSQLNoPCL('UPDATE LIGNEFRAIS SET LF_FRAISINDPR=(SELECT IIF(GPO_FRAISINDPR="SOC", '
  +'(SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM = "SO_FRAISINDPR"), GPO_FRAISINDPR) FROM '+
  'PORT WHERE GPO_CODEPORT=LF_CODEPORT)WHERE LF_TYPEFRAIS="501" AND (LF_FRAISINDPR IS NULL OR LF_FRAISINDPR="")');
*)
  //F BERGER
//  AGLNettoieListes('WNOMELIGMODIFLOT','WNL_GUID',nil);

  //T PITIOT
//  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Valeur de la rupture 4" where dh_nomchamp="GZB_VALRUPT4"');
//  ExecuteSQLContOnExcept('update dechamps set DH_LIBELLE="Vérification client fermé" where DH_NOMCHAMP="GEA_VERIFCLIFERME"');

  //N FOURNEL
//  ExecuteSQLNoPCL('UPDATE LIGNEFORMULE SET GLF_NUMLIGNE = (SELECT GL_NUMORDRE FROM LIGNE WHERE GL_NATUREPIECEG=GLF_NATUREPIECEG AND GL_SOUCHE=GLF_SOUCHE AND'
//  +' GL_NUMERO=GLF_NUMERO AND GL_INDICEG=GLF_INDICEG AND GL_NUMLIGNE=GLF_NUMLIGNE)');

  ExecuteSQLContOnExcept('update menu set mn_accesgrp=(select mn_accesgrp from menu where mn_1=26 and mn_2=3 and mn_3=8 and mn_4=1) where mn_1=26 and mn_2=9 and mn_3=1');
  ExecuteSQLContOnExcept('update menu set mn_accesgrp=(select mn_accesgrp from menu where mn_1=26 and mn_2=3 and mn_3=8 and mn_4=2) where mn_1=26 and mn_2=9 and mn_3=2');
  ExecuteSQLContOnExcept('update menu set mn_accesgrp=(select mn_accesgrp from menu where mn_1=26 and mn_2=3 and mn_3=8 and mn_4=3) where mn_1=26 and mn_2=9 and mn_3=3');

  ExecuteSQLContOnExcept('delete from menu where mn_2 = 3 and mn_tag in (26115,26114,26073,26113)');
  ExecuteSQLContOnExcept('delete from menu where mn_1=26 and mn_2 = 3 and mn_3=8 and mn_4 = 0');

  { GC_20100315_JTR_FQ;PGISIDE;10445 }
  ExecuteSQLContOnExcept('DELETE FROM COMMUN WHERE CO_TYPE = "EC3" AND CO_CODE IN ("RI1", "RI2")');

  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Suite raison sociale ou prénom" where dh_libelle="Suite raison social ou prénom"');
  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Modifiée par l''utilisateur" where dh_libelle="Modifée par l''utilisateur"');
  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Numéro de série" where dh_libelle="Numero de série"');
  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Dépôt destinataire" where dh_libelle="Dépot destinataire"');
  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Référence affectation" where dh_libelle="Réfrence affectation"');
  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Numéro de série" where dh_libelle="Numero de série"');
  ExecuteSQLContOnExcept('update dechamps set dh_libelle="Gestion de la contre marque" where dh_nomchamp="gpp_contremarque"');
  ExecuteSQLContOnExcept('update detables set dt_libelle="Exception Pièce / Domaine" where dt_nomtable="domainepiece"');

  //C DUMAS
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "D" WHERE DH_PREFIXE = "T" AND '
  +'DH_CONTROLE NOT LIKE "%D%" AND DH_NOMCHAMP NOT IN ("T_SOCIETE","T_CREERPAR","T_EXPORTE","T_SCORECLIENT"'
  +',"T_PAYEURECLATEMENT","T_DEBRAYEPAYEUR","T_PASSWINTERNET","T_CLETELEPHONE")');
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "D" WHERE DH_PREFIXE = "RSU" AND'
  +' DH_CONTROLE NOT LIKE "%D%"');

  //T SUBLET
  (*
   { Paramsoc à présent dépendants du SO_GEREFICHELOT }
    SetParamSoc(
      'SO_STKMESSAGELOTUNIQUE',
      GetParamSocSecur('SO_STKMESSAGELOTUNIQUE', False) and GetParamSocSecur('SO_GEREFICHELOT', False)
    );
    SetParamSoc(
      'SO_STKBLOCAGELOTUNIQUE',
      GetParamSocSecur('SO_STKBLOCAGELOTUNIQUE', False) and GetParamSocSecur('SO_GEREFICHELOT', False)
    );
    SetParamSoc(
      'SO_STKBLOCAGELOTEXT',
      GetParamSocSecur('SO_STKBLOCAGELOTEXT', False) and GetParamSocSecur('SO_GEREFICHELOT', False)
    );
  *)
    { Paramsoc à présent dépendants du SO_GEREFICHESERIE }
  (*
    SetParamSoc(
      'SO_STKMESSAGESERUNIQUE',
      GetParamSocSecur('SO_STKMESSAGESERUNIQUE', False) and GetParamSocSecur('SO_GEREFICHESERIE', False)
    );
    SetParamSoc(
      'SO_STKBLOCAGESERUNIQUE',
      GetParamSocSecur('SO_STKBLOCAGESERUNIQUE', False) and GetParamSocSecur('SO_GEREFICHESERIE', False)
    );
   *)
    //JP LAURENT
    // Maj des nouvelles zones concernant le paramétrage des indicateurs
//    ExecuteSQLNoPCL('UPDATE QBPINDICATEUR SET QBZ_ORIGINEREAL="REA", QBZ_CONDITIONREAL=""'
//    +' WHERE (QBZ_ORIGINEREAL is Null) and (QBZ_ORDREPERE<>0) AND (QBZ_VALAFFR<>"")');

//    ExecuteSQLNoPCL('UPDATE QBPINDICATEUR SET QBZ_ORIGINEREAL="SAI", QBZ_CONDITIONREAL=""'
//    +' WHERE (QBZ_ORIGINEREAL is Null) and (QBZ_ORDREPERE<>0) AND (QBZ_VALAFFR="")');

//    ExecuteSQLNoPCL('UPDATE QBPINDICATEUR SET QBZ_ORIGINEREAL="",QBZ_CONDITIONREAL=""  WHERE QBZ_ORDREPERE=0');

//    ExecuteSQLNoPCL('UPDATE QBPLIENINDIC SET QBW_TYPEINDIC="BUD" WHERE QBW_TYPEINDIC is null');

    // Maj nouvelle zone concernant la date de naissance de la session fille
//    ExecuteSQLNoPCL('UPDATE QBPLIENSIMUL SET QBK_DATEREALISE="'+UsDateTime_(IDate1900)+'" WHERE QBK_DATEREALISE is Null');

    // Maj nouvelle zone concernant le type de la séquence
//    ExecuteSQLNoPCL('UPDATE QBPSEQBUDGET SET QBD_ORIGINE="BUD" WHERE QBD_ORIGINE is Null');

    //L GUIPPE
//    ExecuteSQLNoPCL('UPDATE DECHAMPS SET DH_LIBELLE="Code" WHERE DH_NOMCHAMP="GCO_CODECOND"');

    //D KOZA
      { SCM : Maj QDE_PHASE (QDETCIRC) car si pas renseignée pose des problèmes dans placement manuel planning }
    if not IsDossierPCL then
    begin
      (*
      if ExisteSQL('SELECT 1 FROM QDETCIRC WHERE (QDE_PHASE IS NULL) OR (TRIM(QDE_PHASE)="")') then
        ExecuteSqlNoPCL('UPDATE QDETCIRC SET QDE_PHASE=(SELECT QP_PHASE FROM QPHASEITI,QCIRCUIT'
                      +                                ' WHERE QP_CTX=QDE_CTX AND QCI_CTX=QDE_CTX'
                      +' AND QCI_CIRCUIT=QDE_CIRCUIT AND QP_CODITI=QCI_CODITI AND QP_OPEITI=QDE_OPECIRC)'
                      + ' WHERE (QDE_PHASE IS NULL) OR (TRIM(QDE_PHASE)="")');
        *)
    end;

  //J TRIFILIEFF (FQ 18877)
//  ExecuteSQLNoPCL('UPDATE LIGNEFORMULE SET GLF_ND = 1 WHERE (GLF_ND = 0) OR (GLF_ND IS NULL)');
end;

procedure MajVer996;
var
  i: Integer;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //7559
//  ExecuteSQLContOnExcept('UPDATE PARAMSOC SET SOC_DATA="" WHERE SOC_NOM="SO_STKTRFCHEMIN"'
//  +' AND ((SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_STKTRF")<>"X")');
  //7579
  ExecuteSQLContOnExcept ('UPDATE DECHAMPS SET DH_LIBELLE = "Commercial 1" WHERE DH_NOMCHAMP = "GP_REPRESENTANT"');
  ExecuteSQLContOnExcept ('UPDATE DECHAMPS SET DH_LIBELLE = "Commercial 2" WHERE DH_NOMCHAMP = "GP_REPRESENTANT2"');
  ExecuteSQLContOnExcept ('UPDATE DECHAMPS SET DH_LIBELLE = "Commercial 3" WHERE DH_NOMCHAMP = "GP_REPRESENTANT3"');
  //7591
  ExecuteSQLContOnExcept( 'UPDATE DECHAMPS SET DH_CONTROLE = "" WHERE DH_NOMCHAMP = "ADR_CONTACT"' );
//  AglNettoieListes( 'GCADRESSES', '', Nil, 'ADR_CONTACT' );
  //7614
  //FQ;034;15271
  //Mise à jour paramétrage de la pièce TRV
  (*
  if not IsDossierPCL then
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_HISTORIQUE="X",GPP_NATURETIERS="CLI;FOU;" WHERE GPP_NATUREPIECEG = "TRV"');
  // JTR le 10/05/2010 - FQ;010;19070
  // Mise à jour paramètres comptables des pièces qui gère les FAR/FAE
    ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_TYPEECRCPTA = "NOR", GPP_TYPEPASSCPTA = "REE", GPP_TYPEPASSCPTAR = "REE", GPP_TYPEPASSACC = "REE", GPP_TYPEPASSACCR = "REE" WHERE GPP_FAR_FAE <> ""');
  *)
  if not IsDossierPCL then
  begin
    for i := 1 to 2 do
    begin
      InsertChoixExt('GS1', 'TGCS_CHARLIBRE' + IntToStr(i), 'Texte libre '    + IntToStr(i), 'Texte libre '    + IntToStr(i), '');
      InsertChoixExt('GS2', 'TGCS_VALLIBRE'  + IntToStr(i), 'Valeur libre '   + IntToStr(i), 'Valeur libre '   + IntToStr(i), '');
      InsertChoixExt('GS3', 'TGCS_DATELIBRE' + IntToStr(i), 'Date libre '     + IntToStr(i), 'Date libre '     + IntToStr(i), '');
      InsertChoixExt('GS4', 'TGCS_BOOLLIBRE' + IntToStr(i), 'Décision libre ' + IntToStr(i), 'Décision libre ' + IntToStr(i), '');
    end;
  end
end;

procedure MajVer997;
var
  iPERSO: Integer;
  Sql : string;
  TobSidos: Tob; { MajAvant997 et MajVer997: Mémorisation de la tablette (SELECT * FROM COMMUN WHERE CO_TYPE="TRA" AND CO_LIBRE<>"---")}
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  //COMPTA
  UpdateColRTYPEIDBQ;
  ExecuteSQLContOnExcept('delete from swvignettes where swv_codevignette in ("CPVIGNETTEMULBAP", "CPVIGSUIVIBAP")');
  //DOMG
  //FQ;034;15271
  //Création des mouvements "ATD" sur les dépôts de destination
  (*
  if not IsDossierPCL then
  begin
    if  (GetParamSocSecur('SO_GCTRV', False)) and
        (ExisteSql('SELECT 1 FROM PIECE WHERE GP_NATUREPIECEG="TRV"')) and
        (not ExisteSql('SELECT 1 FROM STKMOUVEMENT WHERE GSM_STKTYPEMVT="ATT" AND GSM_QUALIFMVT="ATD"')) then
    begin
      //Création des mouvements "ATD" sur les dépôts de destination
      ExecuteSQL('INSERT INTO STKMOUVEMENT'
              + ' (GSM_STKTYPEMVT, GSM_QUALIFMVT, GSM_GUID, GSM_DATECREATION, GSM_DATEMODIF, GSM_CREATEUR, GSM_UTILISATEUR, GSM_NUMORDRE'
              + ', GSM_ETATMVT, GSM_DATEMVT, GSM_PREFIXEORI, GSM_NATUREORI, GSM_SOUCHEORI, GSM_NUMEROORI, GSM_INDICEORI, GSM_OPECIRCORI, GSM_NUMLIGNEORI, GSM_CODELISTEORI'
              + ', GSM_NUMPREPAORI, GSM_GUIDORI, GSM_DEPOT, GSM_ARTICLE, GSM_TIERS'
              + ', GSM_STATUTDISPO, GSM_STATUTFLUX, GSM_EMPLACEMENT, GSM_LOTEXTERNE, GSM_LOTINTERNE, GSM_SERIEEXTERNE, GSM_SERIEINTERNE, GSM_DATEPEREMPTION, GSM_DATEENTREELOT'
              + ', GSM_DATEDISPO, GSM_INDICEARTICLE, GSM_MARQUE, GSM_CHOIXQUALITE, GSM_QPREVUE, GSM_QPREPA, GSM_QRUPTURE, GSM_PHYSIQUE, GSM_QTEFACT'
              + ', GSM_DATEPREVUE, GSM_IDACTION, GSM_GUIDACTION, GSM_DPA, GSM_DPR, GSM_PMAP, GSM_PMRP, GSM_PSTD, GSM_MONTANT, GSM_MONTANTACTU, GSM_MONTANTCPTA'
              + ', GSM_PRIXSAISIS, GSM_MOTIFMVT, GSM_REFAFFECTATION, GSM_TIERSPROP, GSM_ETATTRANSFERT, GSM_TENUESTOCK'
              + ', GSM_AFFAIRE, GSM_COMPTABILISE, GSM_CONTREMARQUE, GSM_FOURNISSEUR, GSM_REFERENCE, GSM_WBMEMO, GSM_BLOCNOTE)'

              + '(SELECT "ATT" AS GSM_STKTYPEMVT, "ATD" AS GSM_QUALIFMVT, PGIGUID AS GSM_GUID, GSM_DATECREATION, GSM_DATEMODIF, GSM_CREATEUR, GSM_UTILISATEUR, GSM_NUMORDRE'
              + ', "FER" AS GSM_ETATMVT, GSM_DATEMVT, GSM_PREFIXEORI, GSM_NATUREORI, GSM_SOUCHEORI, GSM_NUMEROORI, GSM_INDICEORI, GSM_OPECIRCORI, GSM_NUMLIGNEORI, GSM_CODELISTEORI'
              + ', GSM_NUMPREPAORI, GSM_GUIDORI, GP_DEPOTDEST AS GSM_DEPOT, GSM_ARTICLE, IIF((GDE_TIERS<>""),GDE_TIERS,(SELECT TRIM(SOC_DATA) FROM PARAMSOC WHERE SOC_NOM="SO_GCTIERSDEFAUT")) AS GSM_TIERS'
              + ', GSM_STATUTDISPO, GSM_STATUTFLUX, GSM_EMPLACEMENT, GSM_LOTEXTERNE, GSM_LOTINTERNE, GSM_SERIEEXTERNE, GSM_SERIEINTERNE, GSM_DATEPEREMPTION, GSM_DATEENTREELOT'
              + ', GSM_DATEDISPO, GSM_INDICEARTICLE, GSM_MARQUE, GSM_CHOIXQUALITE, GSM_PHYSIQUE AS GSM_QPREVUE, GSM_QPREPA, GSM_QRUPTURE, 0 AS GSM_PHYSIQUE, GSM_QTEFACT'
              + ', GL_DATELIVRAISON AS GSM_DATEPREVUE, GSM_IDACTION, GSM_GUIDACTION, GSM_DPA, GSM_DPR, GSM_PMAP, GSM_PMRP, GSM_PSTD, GSM_MONTANT, GSM_MONTANTACTU, GSM_MONTANTCPTA'
              + ', GSM_PRIXSAISIS, GSM_MOTIFMVT, GSM_REFAFFECTATION, GSM_TIERSPROP, GSM_ETATTRANSFERT, GSM_TENUESTOCK'
              + ', GSM_AFFAIRE, GSM_COMPTABILISE, GSM_CONTREMARQUE, GSM_FOURNISSEUR, GSM_REFERENCE, GSM_WBMEMO, GSM_BLOCNOTE'
              + ' FROM STKMOUVEMENT, PIECE, LIGNE, DEPOTS'
              + ' WHERE GSM_PREFIXEORI = "GL" AND GSM_STKTYPEMVT="PHY" AND GSM_QUALIFMVT="ETR" AND GSM_NATUREORI="TRV" AND GP_VIVANTE="X" AND GSM_PHYSIQUE > 0'
              + ' AND GSM_NATUREORI=GP_NATUREPIECEG AND GSM_SOUCHEORI=GP_SOUCHE AND GSM_NUMEROORI=GP_NUMERO AND GSM_INDICEORI=GP_INDICEG'
              + ' AND GSM_NATUREORI=GL_NATUREPIECEG AND GSM_SOUCHEORI=GL_SOUCHE AND GSM_NUMEROORI=GL_NUMERO AND GSM_INDICEORI=GL_INDICEG AND GSM_NUMLIGNEORI=GL_NUMORDRE'
              + ' AND GP_DEPOTDEST=GDE_DEPOT)');

      //Recalcul des attendus sur les dépôts de destination
      ExecuteSQL('UPDATE DISPO'
              + ' SET GQ_RESERVEFOU=ISNULL((SELECT SUM(ABS(GSM_QPREVUE)-ABS(GSM_QPREPA)-ABS(GSM_PHYSIQUE))'
              + ' FROM STKMOUVEMENT'
              + ' WHERE GSM_ARTICLE=GQ_ARTICLE AND GSM_DEPOT=GQ_DEPOT AND GSM_ETATMVT<>"SOL"'
              + ' AND ((GSM_STKTYPEMVT="ATT" AND GSM_QUALIFMVT="AAC") OR (GSM_STKTYPEMVT="ATT" AND GSM_QUALIFMVT="ACC") OR (GSM_STKTYPEMVT="ATT" AND GSM_QUALIFMVT="ACP")'
              + ' OR (GSM_STKTYPEMVT="ATT" AND GSM_QUALIFMVT="ADM") OR (GSM_STKTYPEMVT="ATT" AND GSM_QUALIFMVT="APF") OR (GSM_STKTYPEMVT="ATT" AND GSM_QUALIFMVT="APR")'
              + ' OR (GSM_STKTYPEMVT="ATT" AND GSM_QUALIFMVT="ASA") OR (GSM_STKTYPEMVT="ATT" AND GSM_QUALIFMVT="ATD"))), 0)'
              + ' WHERE GQ_DEPOT IN (SELECT GSM_DEPOT FROM STKMOUVEMENT WHERE GSM_STKTYPEMVT="ATT" AND GSM_QUALIFMVT="ATD" GROUP BY GSM_DEPOT) AND GQ_CLOTURE="-"');
    end;
  end;
  *)
  { JTR suite FQ;010;16638 }
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "DPR modifié" WHERE DH_NOMCHAMP = "ELC_MODIFPRIXACHAT"');
  { JTR suite FQ;010;16638 }
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "DPR modifié" WHERE DH_NOMCHAMP = "GLC_MODIFPRIXACHAT"');
  ExecuteSQLNoPcl('UPDATE  DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Z" WHERE DH_PREFIXE = "WAN" AND DH_NOMCHAMP IN("WAN_CODITI","WAN_ACTIF") AND DH_CONTROLE NOT LIKE "%Z%"');
//  AglNettoieListesPlus('GCMULLIGNEFRAIS', 'GL_DPA;GL_DPR;GL_PMAP;GL_PMRP;GL_PSTD', nil, false);
  //AglNettoieListes('WLIGNESCSP','GL_GUIDWOL', nil,'GL_IDENTIFIANTWOL');
  ExecuteSQLContOnExcept('DELETE FROM MODELES WHERE MO_TYPE="E" AND MO_NATURE="TAR" AND MO_CODE IN ("TAR", "TAA")');
  ExecuteSQLContOnExcept('DELETE FROM MODEDATA WHERE MD_CLE LIKE "ETARTAR%" OR MD_CLE LIKE "ETARTAA%"');
  ExecuteSQLNoPcl('DELETE FROM PARAMSOC WHERE SOC_NOM  = "SCO_WSSOIMPACTEWOL"');
  // MNG : traduction métier
  ExecuteSQLContOnExcept('UPDATE YTABLEALERTES SET YTA_LIBELLE = "Opérations commerciales" WHERE YTA_PREFIXE = "ROP"');
  ExecuteSQLContOnExcept('UPDATE YTABLEALERTES SET YTA_LIBELLE = "Propositions commerciales" WHERE YTA_PREFIXE = "RPE"');
  InsertChoixExt('WRO', 'SOUS-TRAIT ACHAT' , 'Sous-traitance d''achat' , 'Sous-trait. achat' , '');
  InsertChoixExt('WRO', 'SOUS-TRAIT PHASE' , 'Sous-traitance de phase' , 'Sous-trait. phase' , '');
//  AGLNettoieListesPlus('YTARIFSMAJ101', 'YTS_PROFILCOMPO', nil, False);
//  AGLNettoieListesPlus('YTARIFSMAJ201', 'YTS_PROFILCOMPO', nil, False);
//  AGLNettoieListesPlus('YTARIFSMAJ211', 'YTS_PROFILCOMPO', nil, False);
//  AGLNettoieListesPlus('YTARIFSMAJ301', 'YTS_PROFILCOMPO', nil, False);
//  AGLNettoieListesPlus('YTARIFSMAJ401', 'YTS_PROFILCOMPO', nil, False);
//  AGLNettoieListesPlus('YTARIFSMAJ501', 'YTS_PROFILCOMPO', nil, False);
//  AGLNettoieListesPlus('YTARIFSMAJ601', 'YTS_PROFILCOMPO', nil, False);
//  ExecuteSQLContOnExcept('update DECHAMPS set DH_LIBELLE="Nombre de niveaux" where DH_PREFIXE="GHP" and DH_NUMCHAMP=85');
  ExecuteSQLContOnExcept('update DETABLES set DT_LIBELLE="Entête paramétrage éco-contribution" where DT_NOMTABLE="ECOENT"');
  ExecuteSQLContOnExcept('Update DECHAMPS set DH_LIBELLE="Gestion de la contremarque" where DH_NOMCHAMP="GPP_CONTREMARQUE"');
  // suppression des 3 champs inexistants dans la liste
{CRM_20100622_CD_012;11219}
//  AglNettoieListes('RBCONNAISSANCE','',Nil,'BCO_RATTACHEMENT1;BCO_RATTACHEMENT2;BCO_RATTACHEMENT3');
  // Suppression des champs WPN_AFFAIRE et WPN_REFLIGNEAFF dans la modif en série des éléments parc
//  ExecuteSQLContOnExcept('Update DECHAMPS set DH_CONTROLE="LDC" where DH_NOMCHAMP="WPN_AFFAIRE" or DH_NOMCHAMP="WPN_REFLIGNEAFF"');
  ExecuteSQLContOnExcept('DELETE FROM MODELES WHERE MO_NATURE="LIV" AND MO_CODE="Z01"');
  ExecuteSQLContOnExcept('DELETE FROM MODEDATA WHERE MD_CLE LIKE "ELIVZ01%"');
  { JTR suite FQ;010;16638 }
//  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "DPR modifié" WHERE DH_NOMCHAMP = "ELC_MODIFPRIXACHAT"');
  { JTR suite FQ;010;16638 }
//  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "DPR modifié" WHERE DH_NOMCHAMP = "GLC_MODIFPRIXACHAT"');
  ExecuteSQLNoPcl('UPDATE DECHAMPS SET DH_CONTROLE = DH_CONTROLE || "Z" WHERE DH_PREFIXE = "WAN" AND DH_NOMCHAMP IN ("WAN_CODEITI","WAN_ACTIF") AND DH_CONTROLE NOT LIKE "%Z%"');
  // M DEMEULEMEESTER
  ExecuteSQLContOnExcept('UPDATE DECHAMPS SET DH_LIBELLE = "Libellé en génération" WHERE DH_LIBELLE = "Libellé de transformation"') ;

// GP Le 08/06/2010 : Fiche N° 27495 : Passage auto en pointage avancé.
  try
  CPActivationPointageAvance ;
  CPInitialisationEtablissementEtAgencesBancaires() ;
  finally
    end;

// BIDOUILLE POUR PERSONNALISATION
// on recharge la TOB mémorisée en majavant997
// 2eme version BIDOUILLE POUR PERSONNALISATION
// on charge la table tempo tmpcommun dans une tob pour restaurer tablette commun
  if TableExiste('TMPCOMMUN') then
  begin
    TobSidos := Tob.Create('TOB SIDOS', nil, -1);
    Sql := 'SELECT * FROM TMPCOMMUN';
    TobSidos.LoadDetailDBFromSql('COMMUN', Sql);
    if Assigned(TobSidos) then
    begin
      for iPERSO := 0 to TobSidos.Detail.Count-1
      do
        TobSidos.Detail[iPerso].InsertDB(nil);

      FreeAndNil(TobSidos);
    end;
  end;

  if not ExisteSQL ('SELECT CO_LIBELLE FROM COMMUN WHERE CO_TYPE="TRA" AND CO_CODE="000"') then
    ExecuteSqlContOnExcept ('INSERT INTO COMMUN (CO_TYPE, CO_CODE, CO_LIBELLE, CO_ABREGE, CO_LIBRE) ' +
                              'VALUES ("TRA", "000", "Langue personnalisée", "", "---")');

end;

procedure MajVer998;
var Q : TQuery;
    TPUBLICOTIS : TOB;
    SQl : String;
begin
  J.Position := J.Position + 1;
  Application.ProcessMessages;
  // traitement des anomalies zarb...
  //ExecuteSQLContOnExcept('UPDATE WNATURETRAVAIL SET WNA_FONCTION="" WHERE WNA_FONCTION IS NULL');
  //ExecuteSQLContOnExcept('UPDATE WORDRELIG SET WOL_GUID="" WHERE WOL_GUID IS NULL');
  // T PETETIN
  // 7894: FQ;034;16804 (Suite 6810+7145) Partage de la tablette WUNITE
  If not ExisteSQL('SELECT CO_TYPE FROM COMMUN WHERE CO_TYPE = "YEB" AND CO_LIBRE = "Y01" AND CO_CODE = "W22"') then
    ExecuteSQLContOnExcept('INSERT INTO COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) values("YEB", "W22", "WUNITE", "TTE", "Y01")');
  //
  ExecuteSQLContOnExcept ('UPDATE PORT SET GPO_TYPEFRAIS="501" WHERE (GPO_TYPEFRAIS IS NULL) OR (GPO_TYPEFRAIS="")');
  //
  ExecuteSQLContOnExcept ('UPDATE BDETETUDE SET BDE_FOURNISSEUR="" WHERE BDE_FOURNISSEUR IS NULL');
  ExecuteSQLContOnExcept ('UPDATE CODECPTA SET GCP_FAMILLETAXE="" WHERE GCP_FAMILLETAXE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE DECISIONACHLIG SET BAD_PABASE=0 WHERE BAD_PABASE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE DECISIONACHLIG SET BAD_COEFUAUS=0 WHERE BAD_COEFUAUS IS NULL');
  ExecuteSQLContOnExcept ('UPDATE DECISIONACHLFOU SET BDF_COEFUAUS=0 WHERE BDF_COEFUAUS IS NULL');
  UpdateDecoupeLigneOuvPlat('BOP_MONTANTPR=0, BOP_COEFFG=0, BOP_COEFFR=0, BOP_COEFFC=0',' AND (BOP_COEFFC IS NULL)');
  UpdateDecoupeLigneOuvPlat('BOP_NATURETRAVAIL="", BOP_FOURNISSEUR=""',' AND (BOP_NATURETRAVAIL IS NULL)');
  ExecuteSQLContOnExcept ('UPDATE AFFAIRE SET AFF_FACTURE=(SELECT ##TOP 1## T_FACTURE FROM TIERS WHERE T_TIERS=AFF_TIERS AND T_NATUREAUXI IN ("CLI","PRO")) WHERE AFF_FACTURE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE AFFAIRE SET AFF_MANDATAIRE="",AFF_CODEBQ="",AFF_TYPEPAIE="",AFF_BQMANDATAIRE="",AFF_IDSPIGAO=0 WHERE AFF_MANDATAIRE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE FACTAFF SET AFA_DATEDEBUTFAC='+DatetoStr(IDate1900)+',AFA_DATEFINFAC='+DateToStr(Idate1900)+' WHERE AFA_DATEDEBUTFAC IS NULL');
  ExecuteSQLContOnExcept ('UPDATE HRPARAMPLANNING SET HPP_REGCOL1="-", HPP_REGCOL2="-", HPP_REGCOL3="-" WHERE HPP_REGCOL1 IS NULL');
  ExecuteSQLContOnExcept ('UPDATE PARCTIERS SET BP1_DATEFINSERIA="20991231" WHERE BP1_DATEFINSERIA IS NULL');
  ExecuteSQLContOnExcept ('UPDATE PIEDECHE SET GPE_CODE="",GPE_TIMBRE=0 WHERE GPE_CODE IS NULL');
  UpdateDecoupeLigne		 ('GL_IDSPIGAO=""',' AND GL_IDSPIGAO IS NULL');
  UpdateDecoupeLigneBase ('BLB_FOURNISSEUR="",BLB_BASEACHAT=0,BLB_VALEURACHAT=0,BLB_TYPEINTERV=""',' AND BLB_FOURNISSEUR IS NULL');
  UpDateDecoupeLigneCompl('GLC_VOIRDETAIL="-"',' AND GLC_VOIRDETAIL IS NULL');
  UpDateDecoupeLigneOUV  ('BLO_UNIQUEBLO=0,BLO_NATURETRAVAIL=""',' AND BLO_UNIQUEBLO IS NULL');
  UpDateDecoupeLignefac  ('BLF_NATURETRAVAIL="",BLF_FOURNISSEUR="",BLF_UNIQUEBLO=0,BLF_TOTALTTCDEV=0',' AND BLF_NATURETRAVAIL IS NULL');
  UpDateDecoupePiece     ('GP_UNIQUEBLO=0',' AND GP_UNIQUEBLO IS NULL');
  ExecuteSQLContOnExcept ('UPDATE PIECETRAIT SET BPE_REGLSAISIE="-" WHERE BPE_REGLSAISIE IS NULL');
  UpDateDecoupePiedbase  ('GPB_FOURN="",GPB_BASEACHAT=0,GPB_VALEURACHAT=0,GPB_DELTAACHAT=0,GPB_TYPEINTERV=""',' AND GPB_FOURN IS NULL');
  UpDateDecoupePieceRG   ('PRG_FOURN=""',' AND PRG_FOURN IS NULL');
  ExecuteSQLContOnExcept ('UPDATE AFFAIRE SET AFF_CONSOMME=0 WHERE AFF_CONSOMME IS NULL');
  ExecuteSQLContOnExcept ('UPDATE PIECERG SET PRG_APPLICABLE="X",PRG_FOURN="",PRG_NUMERORIB=0 WHERE PRG_APPLICABLE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE PIEDBASERG SET PBR_FOURN="" WHERE PBR_FOURN IS NULL');
  // -- 998.5 BTP
  ExecuteSQLContOnExcept ('UPDATE BTTYPEAFFAIRE SET BTY_PRIOCONTRAT="" WHERE BTY_PRIOCONTRAT IS NULL');
  ExecuteSQLContOnExcept ('UPDATE ARTICLE SET GA_GEREDEMPRIX="X" WHERE GA_GEREDEMPRIX IS NULL');
  // -- mise à jour des champs apportés via les BOB dans les cptx
  // cdm_8185
	ExecuteSqlContOnExcept('update banquecp set bq_ebicsactif = "-",BQ_AUXILIAIRE="" where BQ_AUXILIAIRE is null');
  // Cdm_8518
	ExecuteSqlContOnExcept('UPDATE COMMUN SET CO_LIBELLE="Ecritures avec les à nouveaux" WHERE CO_TYPE="RBA" AND CO_CODE="TOU"');
	ExecuteSqlContOnExcept('UPDATE DECOMBOS SET DO_LIBELLE="Liste des cycles de révision actifs" WHERE DO_COMBO="CREVCYCLEACTIF"');
  // Cdm_8519
	ExecuteSqlContOnExcept('UPDATE DECOMBOS SET DO_LIBELLE="Champ de gestion des natures d''immo" WHERE DO_COMBO="ICHAMPIASNATURE"');
	ExecuteSqlContOnExcept('UPDATE DECOMBOS SET DO_LIBELLE="Liste des champs pour édition" WHERE DO_COMBO="ICHAMPPOUREDI"');
  // Cdm_8743
  ExecuteSqlContOnExcept('Update dechamps set DH_LIBELLE="Responsable absences" where DH_PREFIXE="ISQ" and DH_NUMCHAMP=20');
	ExecuteSqlContOnExcept('Update dechamps set DH_LIBELLE="Responsable absences" where DH_PREFIXE="ISW" and DH_NUMCHAMP=20');
	// Cdm_8776
  ExecuteSqlContOnExcept('UPDATE PARAMSOC SET SOC_DATA = (SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_CPJALENCADECA") WHERE SOC_NOM = "SO_TRJALSUBSTITUTION"');
	ExecuteSqlContOnExcept('Update dechamps set DH_LIBELLE="Régime TVA" where DH_PREFIXE="CPT" and DH_NUMCHAMP=13');
  // Cdm_7889
  ExecuteSQLContOnExcept ('UPDATE RHCOMPETENCES SET PCO_SAISISSVALID = "X"');
  ExecuteSQLContOnExcept ('UPDATE RHCOMPETRESSOURCE SET PCH_DEGREMAITRISEC = ""');
  ExecuteSQLContOnExcept ('UPDATE STAGEOBJECTIF SET POS_DEGREMAITRISEC = ""');
  ExecuteSQLContOnExcept ('UPDATE REGLESALERTES SET PLR_MENU = "",PLR_LIEUMENU = "",PLR_MODEALERTES = "",PLR_MESSAGEALERTE = "",PLR_GRPDESTIALERTE = "",PLR_APPLIALERTE = ""');
  ExecuteSQLContOnExcept ('UPDATE HISTOALERTES SET PLH_MESSAGEBLOB = PLH_MESSAGE, PLH_APPLIALERTE="", PLH_MODEALERTES="", PLH_EMAILENVOYE=""');
  ExecuteSQLContOnExcept ('UPDATE PARAMSOC SET SOC_DATA = (SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_PGECABGESTIONMAIL" ) WHERE SOC_NOM="SO_PGGESTIONALERTES"');
  ;
  ExecuteSQLContOnExcept( 'Update DEPORTSAL SET PSE_ISPCSEMPLOI= ""');
  ;
  { Génération des éléments de la tablette des groupes de destinataires - Module Alerte }
  PgInitTabletteTypeDesti;
  { Récupération du paramétrage resp-assist-adjoint tables services et Deportsal }
  PGRecupAffectRoleRh;
  { Génération des suivis de demandes du module absences ERH }
  PGGenereProcessusAbs(False,'','PSN_TYPPROCESSDDE="ABS" AND PSN_ETATSUIVIDDE<>"ATT"');
  ;
  ExecuteSQLContOnExcept ('UPDATE FORMATIONS SET PFO_GUID=PGIGUID WHERE (PFO_GUID IS NULL OR PFO_GUID="")');
  ExecuteSQLContOnExcept ('UPDATE INSCFORMATION SET PFI_GUID=PGIGUID WHERE (PFI_GUID IS NULL OR PFI_GUID="")');
  ;
  { Génération des suivis de demandes du module Formations }
  PGGenereProcessusFor(False,'');
  PGGenereProcessusInscFor(False,'');

  ExecuteSQLContOnExcept ('UPDATE CONTRATTRAVAIL SET PCI_ORDREDECLAED = 0');
  ExecuteSQLContOnExcept ('UPDATE DADSPERIODES SET PDE_ORDREDECLAED = 0');
  Q:= OpenSQL ('SELECT * FROM PUBLICOTIS', True);
  try
     if not Q.EOF then
        begin
        TPUBLICOTIS:= Tob.Create ('PUBLICOTIS', nil, -1);
        TPUBLICOTIS.LoadDetailDB ('PUBLICOTIS', '', '', Q, False);
        TPUBLICOTIS.PutValueAllFille ('PUO_TYPEAFFECT', 'DAS');
        try
           BeginTrans;
           if (Not (ExisteSQL ('SELECT PUO_TYPEAFFECT FROM PUBLICOTIS WHERE PUO_TYPEAFFECT IS NOT NULL AND PUO_TYPEAFFECT<>"DAS"'))) then
              begin
              ExecuteSQL ('DELETE FROM PUBLICOTIS WHERE PUO_TYPEAFFECT IS NULL OR PUO_TYPEAFFECT="DAS"');
              TPUBLICOTIS.InsertDBByNivel (False, 1, 1);
              end;
           CommitTrans;
        except
           RollBack;
           end;
        end;
  finally
     Ferme(Q);
     end;

  ExecuteSQLContOnExcept ('UPDATE PUBLICOTIS SET PUO_LIBELLERUBR = (SELECT PRM_LIBELLE FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_RUBRIQUE = PUO_RUBRIQUE) WHERE PUO_NATURERUB = "REM"');
  ExecuteSQLContOnExcept ('UPDATE PUBLICOTIS SET PUO_LIBELLERUBR = (SELECT PCT_LIBELLE FROM COTISATION WHERE ##PCT_PREDEFINI## PCT_RUBRIQUE = PUO_RUBRIQUE) WHERE PUO_NATURERUB = "COT"');
  ExecuteSQLContOnExcept ('UPDATE PUBLICOTIS SET PUO_LIBELLERUBR = (SELECT PCL_LIBELLE FROM CUMULPAIE WHERE ##PCL_PREDEFINI## PCL_CUMULPAIE = PUO_RUBRIQUE) WHERE PUO_NATURERUB = "CUM"');
  ExecuteSQLContOnExcept ('UPDATE PUBLICOTIS SET PUO_LIBELLESEGMENT = (SELECT SUBSTRING(CO_LIBELLE, 1,85) FROM COMMUN WHERE CO_CODE = PUO_UTILSEGMENT AND CO_TYPE = "PDG") WHERE PUO_TYPEAFFECT = "DAS"');
  ExecuteSQLContOnExcept ('UPDATE SALARIESCOMPL SET PSZ_ZZ1="", PSZ_ZZLIB1="",'+
                          ' PSZ_ZZ2="", PSZ_ZZLIB2="", PSZ_ZZ3="", PSZ_ZZLIB3="",'+
                          ' PSZ_ZZ4="", PSZ_ZZLIB4="", PSZ_ZZ5="", PSZ_ZZLIB5="",'+
                          ' PSZ_ZZ6="", PSZ_ZZLIB6="", PSZ_ZZ7="", PSZ_ZZLIB7="",'+
                          ' PSZ_ZZ8="", PSZ_ZZLIB8="", PSZ_ZZ9="", PSZ_ZZLIB9="",'+
                          ' PSZ_ZZ10="", PSZ_ZZLIB10="", PSZ_ZZ11="",'+
                          ' PSZ_ZZLIB11="", PSZ_ZZ12="", PSZ_ZZLIB12="",'+
                          ' PSZ_ZZ13="", PSZ_ZZLIB13="", PSZ_ZZ14="",'+
                          ' PSZ_ZZLIB14="", PSZ_ZZ15="", PSZ_ZZLIB15="",'+
                          ' PSZ_ZZ16="", PSZ_ZZLIB16="", PSZ_ZZ17="",'+
                          ' PSZ_ZZLIB17="", PSZ_ZZ18="", PSZ_ZZLIB18="",'+
                          ' PSZ_ZZ19="", PSZ_ZZLIB19="", PSZ_ZZ20="",'+
                          ' PSZ_ZZLIB20="", PSZ_COMPLPCS="", PSZ_DEMATBULL = "-"');
  ExecuteSqlContOnExcept ('UPDATE CONTRATTRAVAIL SET PCI_DNATYPEPREAVI1="",'+
                          ' PCI_DNATYPEPREAVI2="", PCI_DNACODUNITEHOR="",'+
                          ' PCI_DEBPREAVIS1="'+UsDateTime_(Idate1900)+'",'+
                          ' PCI_DEBPREAVIS2="'+UsDateTime_(Idate1900)+'",'+
                          ' PCI_FINPREAVIS1="'+UsDateTime_(Idate1900)+'",'+
                          ' PCI_FINPREAVIS2="'+UsDateTime_(Idate1900)+'"');

  ExecuteSqlContOnExcept ('UPDATE DUCSENTETE SET PDU_EFFMOYTR=0');
  ExecuteSqlContOnExcept ('UPDATE DADS2SALARIES SET PD2_SOMMEEXOCET=0');
  ExecuteSQLContOnExcept( 'DELETE FROM decombos WHERE do_combo="PGSALARIE0"');
  // -- Cdm_9413 --
  ExecuteSQLContOnExcept('UPDATE ABSENCESALARIE SET PCN_TYPEIMPUTE="REL" WHERE PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="PRI" AND PCN_PERIODEPAIE LIKE "Reliquat%" AND PCN_TYPEIMPUTE<>"REL"');
  ExecuteSQLContOnExcept('UPDATE ETABCOMPL SET ETB_CPMODEARRONDI="ESU", ETB_CPMODEARRSLDC="IDE" WHERE ETB_CONGESPAYES="X"');
  ExecuteSQLContOnExcept('UPDATE ETABCOMPL SET ETB_CPMODEARRONDI="ANE", ETB_CPMODEARRSLDC="ANE" WHERE ETB_CONGESPAYES="-"');
  ExecuteSQLContOnExcept('INSERT INTO SALARIESCOMPL (PSZ_SALARIE) SELECT PSA_SALARIE FROM SALARIES WHERE PSA_SALARIE NOT IN (SELECT PSZ_SALARIE FROM SALARIESCOMPL)');
  ExecuteSQLContOnExcept('UPDATE SALARIESCOMPL SET PSZ_TYPMODEARRONDI="ETB",PSZ_TYPMODEARRSLDC="ETB",PSZ_CPMODEARRONDI="ESU", PSZ_CPMODEARRSLDC="IDE" WHERE PSZ_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONGESPAYES="X")');
  ExecuteSQLContOnExcept('UPDATE SALARIESCOMPL SET PSZ_TYPMODEARRONDI="PER",PSZ_TYPMODEARRSLDC="PER",PSZ_CPMODEARRONDI="ANE", PSZ_CPMODEARRSLDC="ANE" WHERE PSZ_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONGESPAYES="-")');
  ExecuteSQLContOnExcept('UPDATE ETABCOMPL SET ETB_CPANTICIPE="-", ETB_CPANTNBMAXI=0, ETB_CPANTEXCLUSAL="-", ETB_CPANTEXCLUNBJ=0');
  ExecuteSQLContOnExcept('UPDATE ABSENCESALARIE SET PCN_MVTANTICIPE="-"');
  ExecuteSQLContOnExcept('UPDATE CONTRATTRAVAIL SET PCI_PERPAIEMENTSAL="16", PCI_INTITULCONTRAT=""');
  // Hors recup CDM
  (* Domaine C *)
  ExecuteSQLContOnExcept('UPDATE ANALYTIQ SET Y_DATPER="'+usDateTime_(IDate1900)+'",Y_ENTITY=0,Y_REFGUID=""');
  ExecuteSQLContOnExcept('UPDATE ECRCOMPL SET EC_ENTITY=0,EC_ID=0');
  ExecuteSQLContOnExcept('UPDATE ECRITURE SET E_DATPER="'+usDateTime_(IDate1900)+'",E_ENTITY=0,E_REFGUID=""');
  ExecuteSQLContOnExcept('UPDATE GENERAUX SET G_GUIDASSOCIER="",G_INVISIBLE="-",G_NONTAXABLE="X",G_RESTRICTIONA1="",G_RESTRICTIONA2="",G_RESTRICTIONA3="",G_RESTRICTIONA4="",G_RESTRICTIONA5="",G_TYPECPTTVA=""');
  ExecuteSQLContOnExcept('UPDATE GENERAUXREF SET GER_NONTAXABLE="-"');
  ExecuteSQLContOnExcept('UPDATE JOURNAL SET J_CONTREPARTIEAUX="",J_INVISIBLE="-",J_LIBELLEAUTO="",J_TVACTRL="-"');
  ExecuteSQLContOnExcept('UPDATE SECTION SET S_INVISIBLE="-"');
  ExecuteSQLContOnExcept('UPDATE TXCPTTVA SET TV_DESTACH="",TV_DESTVTE="" WHERE TV_DESTACH IS NULL');
  //
	ExecuteSQLContOnExcept ('UPDATE GENERAUX SET G_TYPECPTTVA="TTC" WHERE G_NATUREGENE IN ("COC","COF","COS","COD","TIC","TID")');
	ExecuteSQLContOnExcept ('UPDATE GENERAUX SET G_TYPECPTTVA="DIV" WHERE G_NATUREGENE IN ("BQE","CAI","DIV","EXT")');
	ExecuteSQLContOnExcept ('UPDATE GENERAUX SET G_TYPECPTTVA="HT" WHERE G_NATUREGENE IN ("IMO","CHA","PRO")');
	ExecuteSQLContOnExcept ('UPDATE GENERAUX SET G_TYPECPTTVA="TVA" WHERE G_GENERAL LIKE "445%"');
  // Edition 7 Paie
  ExecuteSQLContOnExcept ('UPDATE DEPORTSAL SET PSE_PRECOMPTEIJSS="-" , PSE_TYPMETIJSSMAL="DOS" , PSE_METHODIJSSMAL="" , PSE_TYPMETIJSSAT = "DOS", PSE_METHODIJSSAT=""');
  ExecuteSQLContOnExcept ('UPDATE ABSENCESALARIE SET PCN_MTIJSSBRUTE=0, PCN_MTIJSSBRUTE2=0, PCN_MTIJSSNETTE=0, PCN_MTIJSSNETTE2=0 , PCN_NBABSPRECOMPT=0, PCN_NBABSPRECOMPT2=0');
  ExecuteSQLContOnExcept ('UPDATE REGLTIJSS SET PRI_PRECOMPTE="-", PRI_REGULARISATION ="-"');
  ExecuteSQLContOnExcept ('UPDATE DECLARANTATTEST SET PDA_EMAIL=""');
  ExecuteSQLContOnExcept ('UPDATE ATTESTATIONS SET PAS_ATTESTEDI="-" , PAS_ATTESTRECTIFIC="-" , PAS_TYPEREPRISE="" , PAS_NUMDECLARANT = ""');
  ExecuteSQLContOnExcept ('UPDATE EMETTEURSOCIAL SET PET_NOMATTESTEDI="" , PET_PRENATTESTEDI="" , PET_EMAILATTESTEDI=""');
  ExecuteSQLContOnExcept ('UPDATE METHCALCULSALMOY SET PSM_TYPECALC="IDR"');
  ExecuteSQLContOnExcept ('UPDATE ABSENCESALARIE SET PCN_TYPEIMPUTE="REL" WHERE PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="PRI" AND PCN_PERIODEPAIE LIKE "Reliquat%" AND PCN_TYPEIMPUTE<>"REL" AND (PCN_TYPEIMPUTE IS NULL OR PCN_TYPEIMPUTE="")');
  // ---- V8 - V9 BTP ----
  ExecuteSQLContOnExcept('UPDATE BSITUATIONS SET BST_VIVANTE="X",BST_INDICESIT=0 WHERE BST_VIVANTE IS NULL');
  SQL := 'UPDATE BTETAT SET BTA_TYPEACTION="INT",BTA_GESTIONCONSO="-",BTA_AFFAIRE="",BTA_TIERS="",'+
         'BTA_RESSOURCE="",BTA_VALORISE="-",BTA_AFFECTCHANTIER="-",BTA_AFFECTIERS="-",BTA_AFFECTRESS="-" '+
         'WHERE BTA_TYPEACTION IS NULL';
  ExecuteSQLContOnExcept(SQL);
  ExecuteSQLContOnExcept ('UPDATE HISTOBULLETIN SET PHB_DATEDEBRUB="'+UsDateTime (Idate1900)+'", '+
                          'PHB_DATEFINRUB="'+UsDateTime (Idate1900)+'" WHERE PHB_DATEDEBRUB IS NULL');
  ExecuteSQLContOnExcept ('UPDATE PGHISTODETAIL SET PHD_CONTRATTRAV=0 WHERE PHD_CONTRATTRAV IS NULL');
  ExecuteSQLContOnExcept ('UPDATE PIECEINTERV SET BPI_DATECONTRAT=BPI_DATECREATION WHERE BPI_DATECONTRAT IS NULL');
  ExecuteSQLContOnExcept ('UPDATE NATUREPREST SET BNP_SECTION="" WHERE BNP_SECTION IS NULL');
  ExecuteSQLContOnExcept ('UPDATE DEPOTS SET GDE_SECTION="" WHERE GDE_SECTION IS NULL');
  ExecuteSQLContOnExcept ('UPDATE ARTICLEDEMPRIX SET BDP_DPA=0,BDP_PUHTDEV=0,BDP_SELECTIONNE="-" WHERE BDP_DPA IS NULL');
  if ISColumnExists ('ARS_SECTION','RESSOURCE') then ExecuteSQLContOnExcept ('UPDATE RESSOURCE SET ARS_SECTION="" WHERE ARS_SECTION IS NULL');
  ExecuteSQLContOnExcept ('UPDATE NATUREPREST SET BNP_SECTION="" WHERE BNP_SECTION IS NULL');
  ExecuteSQLContOnExcept ('UPDATE BTDOMAINEACT SET BTD_COEFFG_APPEL=BTD_COEFFG,BTD_COEFMARG_APPEL=BTD_COEFMARG WHERE BTD_COEFFG_APPEL IS NULL');
  ExecuteSQLContOnExcept ('UPDATE DETAILDEMPRIX SET BD0_DPA=0,BD0_PUHTDEV=0 WHERE BD0_DPA IS NULL');
  ExecuteSQLContOnExcept ('UPDATE ACOMPTES SET GAC_DATEECHEANCE=GAC_DATEECR,GAC_FOURNISSEUR="" WHERE GAC_DATEECHEANCE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE PGEXCEPTIONS SET PEN_ISRUBPERIOD="-", PEN_PERIODERUB="100", PEN_NATUREPRIME="", '+
              'PEN_APPLICNATURE="'+UsDateTime (Idate1900)+'", '+
              'PEN_MODIFNATURE="'+UsDateTime (Idate1900)+'",PEN_FERME="-" WHERE PEN_ISRUBPERIOD IS NULL');
  ExecuteSQLContOnExcept ('UPDATE ETABCOMPL SET ETB_NUMCAISSECP="",ETB_DSNPOINTDEPOT="GEN", ETB_REMUNEXPATRIE="" WHERE ETB_DSNPOINTDEPOT IS NULL');
  ExecuteSQLContOnExcept ('UPDATE ETABCOMPL SET ETB_DSNPOINTDEPOT="MSA" WHERE ETB_MSAACTIVITE <> "" OR ETB_MSASECTEUR  <> "" OR ETB_MSAUNITEGES <> ""');
  ExecuteSQLContOnExcept ('UPDATE CONTRATPREV SET POP_PREVOYANCEOPT="", POP_PREVOYANCEPOP="" WHERE POP_PREVOYANCEOPT IS NULL');
  ExecuteSQLContOnExcept ('UPDATE CONTRATTRAVAIL SET PCI_DATESIGNRUPTC="'+UsDateTime(idate1900)+'",PCI_TYPECDD="",PCI_SALAIREREF=0, PCI_ETABLIEUTRAV=PCI_ETABLISSEMENT, '+
              'PCI_TRANSACTION="NON", PCI_PORTABPREV="NON", PCI_MOTIFEXCLUDSN = "", PCI_MOTIFSUSPAIE  = "", '+
              'PCI_QUOTITE=0 WHERE PCI_SALAIREREF IS NULL');
  ExecuteSQLContOnExcept ('UPDATE CONTRATTRAVAIL SET PCI_REGIMELOCALDSN = "01" WHERE PCI_ETABLISSEMENT IN (SELECT ETB_ETABLISSEMENT From ETABCOMPL WHERE  ETB_REGIMEALSACE  = "X")');
  ExecuteSQLContOnExcept ('UPDATE CONTRATTRAVAIL SET PCI_REGIMELOCALDSN = "99" WHERE PCI_ETABLISSEMENT IN (SELECT ETB_ETABLISSEMENT From ETABCOMPL  WHERE  ETB_REGIMEALSACE  = "-")');
  ExecuteSQLContOnExcept('UPDATE DECLARATIONAED SET PDN_NBANNULEREMP=0 WHERE PDN_TYPEDECL="51" AND PDN_NBANNULEREMP IS NULL');
  ExecuteSQLContOnExcept('UPDATE DECLARATIONAED SET PDN_NBANNULEREMP=1 WHERE PDN_TYPEDECL="59" AND PDN_NBANNULEREMP IS NULL');
  ExecuteSQLContOnExcept('UPDATE PARPIECE SET GPP_NUMMOISPIECE="-",GPP_NBCARAN=2,GPP_SEPNUMMENSUEL="-",GPP_CARSEPNUM="",GPP_STOCKSSDETAIL= "-",GPP_NUMAUTOLIG= "-" WHERE GPP_NUMMOISPIECE IS NULL');
  ExecuteSQLContOnExcept('UPDATE CODECPTA SET GCP_VENTEACHAT="" WHERE GCP_VENTEACHAT IS NULL');
  ExecuteSQLContOnExcept('UPDATE GENERAUX SET G_TYPECPTTVA="TTC" WHERE G_NATUREGENE IN ("COC","COF","COS","COD","TIC","TID")');
  ExecuteSQLContOnExcept('UPDATE GENERAUX SET G_TYPECPTTVA="DIV" WHERE G_NATUREGENE IN ("BQE","CAI","DIV","EXT")');
  ExecuteSQLContOnExcept('UPDATE GENERAUX SET G_TYPECPTTVA="HT" WHERE G_NATUREGENE IN ("IMO","CHA","PRO")');
  ExecuteSQLContOnExcept('UPDATE GENERAUX SET G_TYPECPTTVA="TVA" WHERE G_GENERAL LIKE "445%"');
  ExecuteSQLContOnExcept('UPDATE PARFOU SET GRF_MULTIFOU="-" WHERE GRF_MULTIFOU IS NULL');
  ExecuteSQLContOnExcept('UPDATE DEPORTSAL SET PSE_PRECOMPTEIJSS="-" , PSE_TYPMETIJSSMAL="DOS" , PSE_METHODIJSSMAL="" , PSE_TYPMETIJSSAT = "DOS", PSE_METHODIJSSAT="" WHERE PSE_PRECOMPTEIJSS IS NULL');
  ExecuteSQLContOnExcept ('UPDATE ABSENCESALARIE SET PCN_MTIJSSBRUTE=0, PCN_MTIJSSBRUTE2=0, PCN_MTIJSSNETTE=0, PCN_MTIJSSNETTE2=0 , PCN_NBABSPRECOMPT=0, PCN_NBABSPRECOMPT2=0 WHERE PCN_MTIJSSBRUTE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE ABSENCESALARIE SET PCN_ASPROLONGATION = "-" WHERE PCN_ASPROLONGATION IS NULL');
  ExecuteSQLContOnExcept ('UPDATE ABSENCESALARIE SET PCN_DARRETINITIAL = "' + UsDateTime(Idate1900) + '" WHERE PCN_DARRETINITIAL IS NULL');
  ExecuteSQLContOnExcept ('UPDATE ABSENCESALARIE SET PCN_TYPEIMPUTE="REL" WHERE PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="PRI" AND PCN_PERIODEPAIE LIKE "Reliquat%" AND PCN_TYPEIMPUTE<>"REL" AND (PCN_TYPEIMPUTE IS NULL OR PCN_TYPEIMPUTE="")');
  ExecuteSQLContOnExcept ('update absencesalarie SET PCN_DSN="-"');
  ExecuteSQLContOnExcept('UPDATE SALARIESCOMPL SET PSZ_CTRLCIVILNIR = "-" WHERE PSZ_CTRLCIVILNIR IS NULL');
  ExecuteSQLContOnExcept ('UPDATE REGLTIJSS SET PRI_PRECOMPTE="-", PRI_REGULARISATION ="-" WHERE PRI_PRECOMPTE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE DECLARANTATTEST SET PDA_EMAIL="",PDA_NOMPERSONNE="" WHERE PDA_EMAIL IS NULL');
  ExecuteSQLContOnExcept ('UPDATE ATTESTATIONS SET PAS_ATTESTEDI="-" , PAS_ATTESTRECTIFIC="-" , PAS_TYPEREPRISE="" , PAS_NUMDECLARANT = "" WHERE PAS_ATTESTEDI IS NULL');
  ExecuteSQLContOnExcept ('UPDATE EMETTEURSOCIAL SET PET_NOMATTESTEDI="" , PET_PRENATTESTEDI="" , PET_EMAILATTESTEDI="" WHERE PET_NOMATTESTEDI IS NULL');
  ExecuteSQLContOnExcept ('UPDATE EMETTEURSOCIAL SET PET_CONTACTDSN=PET_CONT1DUDS, PET_DOMAINEDSN=PET_DOMAINEDUDS1, '+
              'PET_TELDSN=PET_TEL1DADSU, PET_FAXDSN=PET_FAX1DADSU, PET_MAILDSN=PET_APPEL1DUDS '+
              'WHERE PET_CONTACTDSN IS NULL');
  ExecuteSQLContOnExcept ('UPDATE EMETTEURSOCIAL SET PET_CIVILDSN="01" WHERE PET_CIVIL1DADSU="MR"');
  ExecuteSQLContOnExcept ('UPDATE EMETTEURSOCIAL SET PET_CIVILDSN="02" WHERE PET_CIVIL1DADSU="MLE" OR PET_CIVIL1DADSU="MME"');
  ExecuteSQLContOnExcept ('UPDATE EMETTEURSOCIAL SET PET_CIVILDSN="" WHERE PET_CIVILDSN IS NULL');
  ExecuteSQLContOnExcept ('UPDATE EMETTEURSOCIAL SET PET_CONTACTDSN2=PET_CONT2DUDS, PET_DOMAINEDSN2=PET_DOMAINEDUDS2, PET_TELDSN2=PET_TEL2DADSU, PET_FAXDSN2=PET_FAX2DADSU, PET_MAILDSN2=PET_APPEL2DUDS WHERE PET_CONTACTDSN2 IS NULL');
  ExecuteSQLContOnExcept ('UPDATE EMETTEURSOCIAL SET PET_CIVILDSN2="01" WHERE PET_CIVIL2DADSU="MR"');
  ExecuteSQLContOnExcept ('UPDATE EMETTEURSOCIAL SET PET_CIVILDSN2="02" WHERE PET_CIVIL2DADSU="MLE" OR PET_CIVIL2DADSU="MME"');
  ExecuteSQLContOnExcept ('UPDATE EMETTEURSOCIAL SET PET_CIVILDSN2="" WHERE PET_CIVILDSN2 IS NULL');
  ExecuteSQLContOnExcept ('UPDATE EMETTEURSOCIAL SET PET_CONTACTDSN3=PET_CONT3DUDS, PET_DOMAINEDSN3=PET_DOMAINEDUDS3, PET_TELDSN3=PET_TEL3DADSU, PET_FAXDSN3=PET_FAX3DADSU, PET_MAILDSN3=PET_APPEL3DUDS WHERE PET_CONTACTDSN3 IS NULL');
  ExecuteSQLContOnExcept ('UPDATE EMETTEURSOCIAL SET PET_CIVILDSN3="01" WHERE PET_CIVIL3DADSU="MR"');
  ExecuteSQLContOnExcept ('UPDATE EMETTEURSOCIAL SET PET_CIVILDSN3="02" WHERE PET_CIVIL3DADSU="MLE" OR PET_CIVIL3DADSU="MME"');
  ExecuteSQLContOnExcept ('UPDATE EMETTEURSOCIAL SET PET_CIVILDSN3="" WHERE PET_CIVILDSN3 IS NULL');
  ExecuteSQLContOnExcept ('UPDATE METHCALCULSALMOY SET PSM_TYPECALC="IDR" WHERE PSM_TYPECALC IS NULL');
  ExecuteSQLContOnExcept ('UPDATE BTTYPEAFFAIRE SET BTY_PRIOCONTRAT="" WHERE BTY_PRIOCONTRAT IS NULL');
  ExecuteSQLContOnExcept ('UPDATE ARTICLE SET GA_GEREDEMPRIX="X",GA_SECTION="",GA_GEREANAL="X",GA_QUALIFUNITEACH="",GA_PAUA=0, GA_COEFCONVQTEACH=0 WHERE GA_GEREDEMPRIX IS NULL');
  ExecuteSQLContOnExcept ('UPDATE BDETETUDE SET BDE_FOURNISSEUR="" WHERE BDE_FOURNISSEUR IS NULL');
  ExecuteSQLContOnExcept ('UPDATE CODECPTA SET GCP_FAMILLETAXE="" WHERE GCP_FAMILLETAXE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE DECISIONACHLIG SET BAD_PABASE=0,BAD_COEFUAUS=0,BAD_COEFUSUV=0  WHERE BAD_PABASE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE DECISIONACHLFOU SET BDF_COEFUAUS=0 WHERE BDF_COEFUAUS IS NULL');
  UpdateDecoupeLigneOuvPlat('BOP_MONTANTPR=0, BOP_COEFFG=0, BOP_COEFFR=0, BOP_COEFFC=0,BOP_NATURETRAVAIL="",BOP_FOURNISSEUR=""',' AND (BOP_COEFFC IS NULL)');
  UpdateDecoupeLigneOuvPlat('BOP_TENUESTOCK=(select ga_tenuestock from article where ga_article=bop_article),BOP_COEFCONVQTE=0,BOP_COEFCONVQTEVTE=0',' AND (bop_tenuestock IS NULL)');
  ExecuteSQLContOnExcept ('UPDATE AFFAIRE SET AFF_FACTURE=(SELECT ##TOP 1## T_FACTURE FROM TIERS WHERE T_TIERS=AFF_TIERS AND T_NATUREAUXI IN ("CLI","PRO")) WHERE AFF_FACTURE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE AFFAIRE SET AFF_MANDATAIRE="",AFF_CODEBQ="",AFF_TYPEPAIE="",AFF_BQMANDATAIRE="",AFF_IDSPIGAO=0 WHERE AFF_MANDATAIRE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE AFFAIRE SET AFF_MODELEWORD="" WHERE AFF_MODELEWORD IS NULL');
  ExecuteSQLContOnExcept ('UPDATE AFFAIRE SET AFF_SSTRAITANCE="-", AFF_DATESSTRAIT="'+USDATETIME(IDate1900)+'" where aff_sstraitance is null');
  ExecuteSQLContOnExcept ('UPDATE AFFAIRE SET AFF_DEMANDEREGLE="-",AFF_TVAREDUITE="-" WHERE AFF_DEMANDEREGLE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE FACTAFF SET AFA_DATEDEBUTFAC="'+USDATETIME(IDate1900)+'",AFA_DATEFINFAC="'+USDATETIME(Idate1900)+'" WHERE AFA_DATEDEBUTFAC IS NULL');
  ExecuteSQLContOnExcept ('UPDATE HRPARAMPLANNING SET HPP_REGCOL1="-", HPP_REGCOL2="-", HPP_REGCOL3="-",HPP_AFFEVTMAT="-" WHERE HPP_REGCOL1 IS NULL');
  ExecuteSQLContOnExcept ('UPDATE PARCTIERS SET BP1_DATEFINSERIA="20991231" WHERE BP1_DATEFINSERIA IS NULL');
  ExecuteSQLContOnExcept ('UPDATE PIEDECHE SET GPE_CODE="",GPE_TIMBRE=0,GPE_FOURNISSEUR="" WHERE GPE_CODE IS NULL');
  UpdateDecoupeLigne('GL_IDSPIGAO="",GL_COEFCONVQTEVTE=0',' AND GL_IDSPIGAO IS NULL');
  UpdateDecoupeLigneBase('BLB_FOURNISSEUR="",BLB_BASEACHAT=0,BLB_VALEURACHAT=0,BLB_TYPEINTERV=""',' AND BLB_FOURNISSEUR IS NULL');
  UpDateDecoupeLigneCompl('GLC_VOIRDETAIL="-",GLC_CODEMATERIEL="",GLC_BTETAT="",GLC_IDEVENTMAT=0,GLC_NUMEROTATION="",GLC_NUMFORCED="-"  ',' AND GLC_CODEMATERIEL IS NULL');
  UpDateDecoupeLigneOUV('BLO_UNIQUEBLO=0,BLO_NATURETRAVAIL="",BLO_COEFCONVQTE=0,BLO_COEFCONVQTEVTE=0',' AND BLO_UNIQUEBLO IS NULL');
  UpDateDecoupeLignefac('BLF_NATURETRAVAIL="",BLF_FOURNISSEUR="",BLF_UNIQUEBLO=0,BLF_TOTALTTCDEV=0,BLF_MTPRODUCTION=0',' AND BLF_NATURETRAVAIL IS NULL');
  UpDateDecoupePiece('GP_UNIQUEBLO=0,GP_CODEMATERIEL="",GP_BTETAT=""',' AND GP_UNIQUEBLO IS NULL');
  ExecuteSQLContOnExcept ('UPDATE PIECETRAIT SET BPE_REGLSAISIE="-" WHERE BPE_REGLSAISIE IS NULL');
  UpDateDecoupePiedbase('GPB_FOURN="",GPB_BASEACHAT=0,GPB_VALEURACHAT=0,GPB_DELTAACHAT=0,GPB_TYPEINTERV=""',' AND GPB_FOURN IS NULL');
  ExecuteSQLContOnExcept('UPDATE PIECERG SET PRG_APPLICABLE="X",PRG_FOURN="",PRG_NUMERORIB=0 WHERE PRG_APPLICABLE IS NULL');
  ExecuteSQLContOnExcept ('UPDATE REMUNERATION SET PRM_CONDIUTILIS1="",PRM_CONDIUTIL2="",PRM_VALELTDYNTAB="" WHERE PRM_CONDIUTILIS1 IS NULL');
  ExecuteSQLContOnExcept ('UPDATE REMUNERATION SET PRM_PERIODERUB="100", PRM_ISRUBPERIOD="-", PRM_NATUREPRIME="", '+
              'PRM_APPLICNATURE="'+UsDateTime (Idate1900)+'", '+
              'PRM_MODIFNATURE="'+UsDateTime (Idate1900)+'", '+
              'PRM_AUTREELEMENT = "", PRM_APPLICAUTREELE = "'+UsDateTime (Idate1900)+'", '+
              'PRM_MODIFAUTREELEM = "'+UsDateTime (Idate1900)+'" WHERE PRM_PERIODERUB IS NULL');
  ExecuteSQLContOnExcept ('UPDATE DECLARATIONS SET PDT_CIRCACC7BIS="",'+
              ' PDT_ACCIDENTTRAV="-", PDT_ACCIDENTTRAJET="-",'+
              ' PDT_CONTRATCDI="-", PDT_CONTRATCDD="-",'+
              ' PDT_CONTRATAPP="-", PDT_CONTRATINT="-",'+
              ' PDT_CONTRATAUTRE="-", PDT_BOOLTEMOIN="-",'+
              ' PDT_PERSONNEAVISEE="-" WHERE PDT_CIRCACC7BIS IS NULL');
  ExecuteSQLContOnExcept ('INSERT INTO SALARIESCOMPL (PSZ_SALARIE) SELECT PSA_SALARIE '+
              'FROM SALARIES WHERE PSA_SALARIE NOT IN (SELECT PSZ_SALARIE FROM SALARIESCOMPL)');
  ExecuteSQLContOnExcept ('UPDATE SALARIESCOMPL SET PSZ_TYPMODEARRONDI = "ETB" '+
              'WHERE (PSZ_TYPMODEARRONDI IS NULL OR PSZ_TYPMODEARRONDI="") AND PSZ_SALARIE<>""');
  ExecuteSQLContOnExcept ('UPDATE SALARIESCOMPL SET PSZ_TYPMODEARRSLDC = "ETB" '+
              'WHERE (PSZ_TYPMODEARRSLDC IS NULL OR PSZ_TYPMODEARRSLDC="") AND PSZ_SALARIE<>""');
  ExecuteSQLContOnExcept ('UPDATE SALARIESCOMPL SET PSZ_CPMODEARRSLDC = (SELECT ETB_CPMODEARRSLDC FROM ETABCOMPL '+
              'WHERE ETB_ETABLISSEMENT =(SELECT PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE = PSZ_SALARIE)) '+
              'WHERE (PSZ_CPMODEARRSLDC IS NULL OR PSZ_CPMODEARRSLDC="") AND PSZ_SALARIE<>""');
  ExecuteSQLContOnExcept ('UPDATE SALARIESCOMPL SET PSZ_CPMODEARRONDI = (SELECT ETB_CPMODEARRONDI FROM ETABCOMPL '+
              'WHERE ETB_ETABLISSEMENT =(SELECT PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE = PSZ_SALARIE)) '+
              'WHERE (PSZ_CPMODEARRONDI IS NULL OR PSZ_CPMODEARRONDI="") AND PSZ_SALARIE<>""');
  ExecuteSQLContOnExcept('UPDATE SALARIESCOMPL SET PSZ_DIGITALELECTRO="" WHERE PSZ_DIGITALELECTRO IS NULL');
  ExecuteSQLContOnExcept('UPDATE SALARIESCOMPL SET PSZ_DIGITALPAPER="" WHERE PSZ_DIGITALPAPER IS NULL');
  ExecuteSQLContOnExcept('UPDATE SALARIESCOMPL SET PSZ_DIGITALEXCLU="" WHERE PSZ_DIGITALEXCLU IS NULL');
  ExecuteSQLContOnExcept ('UPDATE CONSOMMATIONS SET BCO_LINKEQUIPE="",BCO_FAMILLETAXE1="",BCO_DATETRAVAUX=BCO_DATEMOUV+" 08:00:00" WHERE BCO_LINKEQUIPE IS NULL');
  if ISColumnExists ('T_CODEBARRE','TIERS') and ISColumnExists ('T_QUALIFCODEBARRE','TIERS') then
  begin
  ExecuteSQLContOnExcept ('UPDATE TIERS SET T_CODEBARRE="",T_QUALIFCODEBARRE="" WHERE T_CODEBARRE IS NULL');
  end;
  ExecuteSQLContOnExcept ('UPDATE BTFAMILLEMAT SET BFM_NONGEREPLANNING="X" WHERE BFM_NONGEREPLANNING IS NULL');
  ExecuteSQLContOnExcept ('UPDATE HRPARAMPLANNING SET HPP_FAMMATGERE="" WHERE HPP_FAMMATGERE IS NULL');
  if ISColumnExists ('SH_NUMMOISPIECE','SOUCHE') then
  begin
//  ExecuteSQLContOnExcept ('update SOUCHE SET SH_NUMMOISPIECE="-" WHERE SH_NUMMOISPIECE IS NULL');
  ExecuteSQLContOnExcept ('update SOUCHE SET SH_NUMMOISPIECE=(select ##TOP 1## GPP_NUMMOISPIECE FROM PARPIECE where GPP_SOUCHE=SH_SOUCHE) WHERE SH_TYPE="GES"');
  end;
  ExecuteSQLContOnExcept ('update BTPARDOC SET BPD_IMPDETAVE="-", BPD_IMPCUMULAVE="-" WHERE BPD_IMPDETAVE IS NULL');
  // ------------------------------
  ExecuteSQLContOnExcept ('UPDATE MENU SET MN_ACCESGRP="----------------------------------------------------------------------------------------------------" WHERE mn_accesgrp is null');
end;


// =====   MAJAPRES ================================================

function MAJHalleyApres(VSoc: Integer; MajLab1, MajLab2: TLabel; MajJauge: TProgressBar): boolean;
var savFCurrentAlias:string;
    savEnableDeshare:boolean;
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

  // WARNING !!!!! JS1 : 12102007 : ON FORCE LA DESACTIVATION DU DESHARE AU DEBUT DU MAJVER SUITE BUG OBSERVE SUR SIC !!!!!!!
  V_PGI.EnableDeShare := False;

  //if TOB(V_PGI.TOBSOC)<>Nil then BEGIN TOB(V_PGI.TOBSOC).Free ; V_PGI.TOBSOC:=NIL ; END ;
  //V_PGI.DEDejaCharge:= False;
  //ChargeTablePrefixe(FALSE, TRUE);  // ajout PCS pour s'assurer du chargement en mémoire du nouveau dictionnaire pour utiliser les BOBs
  // ----- V9 ----------------------
  If VSoc <  900 Then MajVer900; //Nouvelle SOCREF v900
  If VSoc <  901 Then MajVer901;
  If VSoc <  902 Then MajVer902;
  If VSoc <  903 Then MajVer903;
  If VSoc <  904 Then MajVer904;
  If VSoc <  905 Then MajVer905;
  If VSoc <  906 Then MajVer906;
  If VSoc <  907 Then MajVer907;
  If VSoc <  909 Then MajVer909;
  If VSoc <  910 Then MajVer910;
  If VSoc <  911 Then MajVer911;
  If VSoc <  912 Then MajVer912;
  If VSoc <  913 Then MajVer913;
  If VSoc <  914 Then MajVer914;
  If VSoc <  915 Then MajVer915;
  If VSoc <  917 Then MajVer917;
  If VSoc <  918 Then MajVer918;
  If VSoc <  919 Then MajVer919;
  If VSoc <  920 Then MajVer920;
  If VSoc <  921 Then MajVer921;
  If VSoc <  922 Then MajVer922;
  If VSoc <  923 Then MajVer923;
  If VSoc <  924 Then MajVer924;
  If VSoc <  925 Then MajVer925;
  If VSoc <  926 Then MajVer926;
  If VSoc <  927 Then MajVer927;
  If VSoc <  928 Then MajVer928;
  If VSoc <  929 Then MajVer929;
  If VSoc <  930 Then MajVer930;
  If VSoc <  931 Then MajVer931;
  If VSoc <  932 Then MajVer932;
  If VSoc <  933 Then MajVer933;
  If VSoc <  934 Then MajVer934;
  If VSoc <  935 Then MajVer935;
  If VSoc <  936 Then MajVer936;
  If VSoc <  937 Then MajVer937;
  If VSoc <  938 Then MajVer938;
  If VSoc <  939 Then MajVer939;
  If VSoc <  940 Then MajVer940;
  If VSoc <  941 Then MajVer941;
  If VSoc <  942 Then MajVer942;
  If VSoc <  943 Then MajVer943;
  If VSoc <  944 Then MajVer944;
  If VSoc <  945 Then MajVer945;
  If VSoc <  946 Then MajVer946;
  If VSoc <  947 Then MajVer947;
  If VSoc <  948 Then MajVer948;
  If VSoc <  949 Then MajVer949;
  If VSoc <  951 Then MajVer951;
  If VSoc <  952 Then MajVer952;
  If VSoc <  953 Then MajVer953;
  If VSoc <  954 Then MajVer954;
  If VSoc <  955 Then MajVer955;
  If VSoc <  957 Then MajVer957;
  If VSoc <  958 Then MajVer958;
  If VSoc <  959 Then MajVer959;
  If VSoc <  960 Then MajVer960;
  If VSoc <  961 Then MajVer961;
  If VSoc <  962 Then MajVer962;
  If VSoc <  963 Then MajVer963;
  If VSoc <  964 Then MajVer964;
  If VSoc <  965 Then MajVer965;
  If VSoc <  966 Then MajVer966;
  If VSoc <  967 Then MajVer967;
  If VSoc <  968 Then MajVer968;
  If VSoc <  969 Then MajVer969;
  If VSoc <  970 Then MajVer970;
  If VSoc <  971 Then MajVer971;
  If VSoc <  972 Then MajVer972;
  If VSoc <  973 Then MajVer973;
  If VSoc <  974 Then MajVer974;
  If VSoc <  975 Then MajVer975;
  If VSoc <  976 Then MajVer976;
  If VSoc <  977 Then MajVer977;
  If VSoc <  978 Then MajVer978;
  If VSoc <  979 Then MajVer979;
  If VSoc <  981 Then MajVer981;
  If VSoc <  982 Then MajVer982;
  If VSoc <  983 Then MajVer983;
  If VSoc <  984 Then MajVer984;
  If VSoc <  985 Then MajVer985;
  If VSoc <  987 Then MajVer987;
  If VSoc <  988 Then MajVer988;
  If VSoc <  990 Then MajVer990; //V9Ed1
  If VSoc <  995 Then MajVer995;
  If VSoc <  996 Then MajVer996;
  If VSoc <  997 Then MajVer997;
  If VSoc <  998 Then MajVer998; //V9Ed2

  if not IsDossierPCL then
  begin
    AddNewNaturesGC;
//    GPMajSTKNature;
  end;
//  MajSouche;
  MAJStandardPaie;
  MAJListeProduits;
  MajBTMAJSTRUCTURES; // --- pour BTP
  ForceConfidentialiteMenu;
  Result := True;
  if V_PGI.SAV then LogAGL('Fin Mise à jour à ' + DateTimeToStr(Now));
  if (V_PGI.ModePCL='1') then       //js1 traitements pour pgimajver batché expert
  begin
    // js1 on loggue la fin de maj
    LogPCL('FIN=' + DateTimeToStr(Now),FicLogPCL);
    savEnableDeshare:=V_PGI.EnableDeshare;
    V_PGI.EnableDeshare:=True;
    ExecuteSQLContOnExcept('UPDATE DOSSIER SET DOS_VERROU="'+ stSavVerrou+
    '", DOS_UTILISATEUR="'+ stSavUser +'" WHERE DOS_NODOSSIER="'+V_PGI.NoDossier+'"');
    V_PGI.EnableDeshare:=savEnableDeshare;

  // js1 100706 Compression de la base
    if okShrinkPCL then begin
        LogPCL('DEBUT SHRINK=' + DateTimeToStr(Now),FicLogPCL);
        savFCurrentAlias:=V_PGI.FCurrentAlias;
        V_PGI.FCurrentAlias:=GetDb0name;
        ShrinkDatabaseMsSql('DB'+V_PGI.NoDossier) ;
        V_PGI.FCurrentAlias:=savFCurrentAlias;
        // js1 on loggue la fin de maj
        LogPCL('FIN SHRINK=' + DateTimeToStr(Now),FicLogPCL);
    end;
  end;
  MajStructure(True);
  // BTP
  Deverrouille;
  // --
  //JS1 24102007 on loggue la fin de la mise a jour dans le jnalevent FQ12251
  MAJJnalEvent ('MAV', 'OK', 'Mise à jour de version ' + V_PGI.NumVersion, '');
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
    ExecuteSQLContOnExcept('DELETE FROM ELTNATIONAUX WHERE PEL_PREDEFINI="STD" AND PEL_MONTANT=0 AND PEL_MONTANTEURO=0  AND PEL_THEMEELT="RET" AND PEL_DATEVALIDITE = "' +
      UsDateTime(DD) + '"');
    TSoc := THTable.Create(Application);
    TSoc.DatabaseName := DBSoc.DatabaseName;
    TSoc.Tablename := 'ELTNATIONAUX';
    TSoc.Indexname := 'PEL_CLE1';
    TSoc.Open;
    TRef := OpenTableRef('ELTNATIONAUX', 'PEL_CLE1');
    while not TRef.EOF do
    begin
      if (not TSoc.Findkey([TRef.FindField('PEL_PREDEFINI').AsString,
                            TRef.FindField('PEL_NODOSSIER').AsString,
                            TRef.FindField('PEL_CODEELT').AsString,
                            TRef.FindField('PEL_CONVENTION').AsString,
                            TRef.FindField('PEL_DATEVALIDITE').AsDateTime])) and
                            (TRef.FindField('PEL_PREDEFINI').AsString = 'STD') then
        AddParamGC(TSoc, TRef);
      TRef.Next;
    end;
    Ferme(TRef);
    TSoc.Close;
    TSoc.Free;
  end;
end;

//Ludovic MONTAVON Demande N° 1418
procedure BasculeExpertEtConseil;
  //Réplication des liens experts comptables et conseils fiscaux dans des liens signataires
  //de déclaration de la v8 (anl_declaration<>'')
    procedure BasculeExpertEtConseil_ (tip:string) ;
    var anl_ : Tob ;
        q : TQuery ;
    begin

      q := OpenSQL ('select * from ANNULIEN ' +
                    'where ANL_FONCTION="DIV" and ANL_TYPEPER="' +tip+ '"', True);
      anl_ := tob.create('ANNULIEN', nil, -1);
      if not q.eof then
      begin
        while not q.Eof do
        begin
          anl_.selectDb('', q);
          anl_.putValue('ANL_FONCTION', 'SGN') ;
          anl_.putValue('ANL_TYPEPER', '') ;
          if tip='EXPCPT' then anl_.putValue('ANL_DECLARATION', 'EXP')
          else if tip='CFIS' then anl_.putValue('ANL_DECLARATION', 'CSL') ;
          anl_.putValue('ANL_NOADMIND', ''); //indépendant par défaut
          anl_.insertOrUpdateDb;
          q.Next;
        end ;
      end ;
      Ferme(q);
      anl_.Free ;
    end ;
begin
  If IsMonoOuCommune Then
  begin
    BasculeExpertEtConseil_('EXPCPT') ;
    BasculeExpertEtConseil_('CFIS') ;
  end ;
end ;


//DESSEIGNET Marie Christine Demande N° 2069
//Voir Celine BOUET si questions sur ce code
Procedure ModifPlanningGA;
Var
  vSt    : String;
  vStDeb : String;
  vStFin : String;

//modif Planning GA
Begin
  If isMssql Then
  Begin
    DateTimeToString( vStDeb, 'hh:nn:ss', GetParamsocSecur('SO_AFPMDEBUT', '14:00:00') );
    DateTimeToString( vStFin, 'hh:nn:ss', GetParamsocSecur('SO_AFPMFIN', '18:00:00') );
    vStDeb := '"' + vStDeb + '"';
    vStFin := '"' + vStFin + '"';
  End
  Else If isOracle Then
  Begin
    DateTimeToString( vStDeb, 'hh', GetParamsocSecur('SO_AFPMDEBUT', '14:00:00') );
    DateTimeToString( vStFin, 'hh', GetParamsocSecur('SO_AFPMFIN', '18:00:00') );
    vStDeb := vStDeb + '/24';
    vStFin := vStFin + '/24';
  End
  Else If isDB2 Then
  Begin
    DateTimeToString( vStDeb, 'hh', GetParamsocSecur('SO_AFPMDEBUT', '14:00:00') );
    DateTimeToString( vStFin, 'hh', GetParamsocSecur('SO_AFPMFIN', '18:00:00') );
    vStDeb := vStDeb + ' hours ';
    vStFin := vStFin + ' hours ';
  End;

  vSt := 'UPDATE AFPLANNING SET apl_heurefin_pla = apl_datefinpla + ' + vStFin + ' where apl_heurefin_pla = apl_datefinpla + ' + vStDeb + ' and apl_typepla = "PLA"';
  ExecuteSQLNoPCL( vSt );
  vSt := 'UPDATE AFPLANNING SET apl_heurefin_real = apl_datefinreal + ' + vStFin + ' where apl_heurefin_real = apl_datefinreal + ' + vStDeb + ' and apl_typepla = "PLA"';
  ExecuteSQLNoPCL( vSt );

  If isMssql Then
  Begin
    DateTimeToString( vStDeb, 'hh:nn:ss', GetParamsocSecur('SO_AFAMDEBUT', '8:00:00') );
    DateTimeToString( vStFin, 'hh:nn:ss', GetParamsocSecur('SO_AFAMFIN', '12:00:00') );
    vStDeb := '"' + vStDeb + '"';
    vStFin := '"' + vStFin + '"';
  End
  Else If isOracle Then
  Begin
    DateTimeToString( vStDeb, 'hh', GetParamsocSecur('SO_AFPMDEBUT', '8:00:00') );
    DateTimeToString( vStFin, 'hh', GetParamsocSecur('SO_AFPMFIN', '12:00:00') );
    vStDeb := vStDeb + '/24';
    vStFin := vStFin + '/24';
  End
  Else If isDB2 Then
  Begin
    DateTimeToString( vStDeb, 'hh', GetParamsocSecur('SO_AFPMDEBUT', '8:00:00') );
    DateTimeToString( vStFin, 'hh', GetParamsocSecur('SO_AFPMFIN', '12:00:00') );
    vStDeb := vStDeb + ' hours ';
    vStFin := vStFin + ' hours ';
  End;
  ExecuteSQLNoPCL( 'UPDATE AFPLANNING SET apl_heurefin_pla = apl_datefinpla + ' + vStFin + 'where apl_heurefin_pla = apl_datefinpla + ' + vStDeb + ' and apl_typepla = "PLA"');
  ExecuteSQLNoPCL( 'UPDATE AFPLANNING SET apl_heurefin_real = apl_datefinreal + ' + vStFin + 'where apl_heurefin_real = apl_datefinreal + ' + vStDeb + ' and apl_typepla = "PLA"');
End;


//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
//============= MAJ PENDANT ===================================================
//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


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
End;


//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
//============= CONTROLES ===================================================
//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

function MAJHalleyIsPossible(VSoc: Integer; MajLab1, MajLab2: TLabel; MajJauge: TProgressBar): boolean;
Begin

  Result:=True;
  if IsOracle then //JS112062008 Test du jeu de caractères : on sort si KO
  begin
        Result:=okJeuCarOra;
  end;
  if Vsoc < 850 then
  begin
        Result:=False ;
        PGIBox('La version de la base ('+intToStr(vsoc)+') est inférieure à 850. ' + #13#10
             + 'Vous devez au préalable faire une mise à jour intermédiaire en structure 850 (Version 8) .','Attention');
  end;
  If (V_PGI.Driver=dbINTRBASE) then
  begin
        Result:=False ;
        PGIBox('Le SGBDR Interbase n''est plus supporté ','Attention');
  end;
  If (V_PGI.Driver=dbORACLE7) then
  begin
        Result:=False ;
        PGIBox('Le SGBDR Oracle 7 n''est plus supporté ','Attention');
  end;
  if not result then exit;

  if (GetParamSocSecur('SO_DISTRIBUTION','000')<>'014')  then
  begin
        Result:=False ;
        PGIBox('Vous ne pouvez pas traiter cette base ' + #13#10
             + '( Code distribution '+ GetParamSoc('SO_DISTRIBUTION')+ ' )'+ #13#10
             + 'avec ce programme. ','Attention');
  end;


  if not ISUtilisable Then
  begin
    Result:=False ;
    PGIBox('Un traitement de mise à jour est dèjà en cours sur cette base' + #13#10,'Attention');
  end;

End;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 23/03/2007
Modifié le ... :   /  /
Description .. : Juste pour éviter d'avoir toujours à tester :
Suite ........ : if not IsDossierPCL then ...
Mots clefs ... :
*****************************************************************}
function ExecuteSQLNoPCL(const Sql: WideString): Integer;
begin
  if not IsDossierPCL then
    Result := ExecuteSQLContOnExcept(Sql)
  else
    Result := 0
end;

//===============================================================
//
// ExecuteSQL Continue On Exception en mode SAV
// 05/04/07. ALR & LA
//===============================================================
function ExecuteSQLContOnExcept(const Sql: WideString): Integer;
begin
  result := -1;

  try
    result := ExecuteSQL(sql);
  except
    on e: exception do
    begin
      if  (v_pgi.SAV) then
        PgiError(E.message);
      // nb: si pas  sav ... raise=> pgierror + haut
      trace.traceError('SQL','Exception '+e.message );
      trace.traceError('SQL',sql);
      // ALR 05/04/07
      if  (not v_pgi.SAV) then
        raise;
    end;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Abélard
Créé le ...... : 18/04/2007
Modifié le ... : 18/04/2007
Description .. : Mise en Forme d'une commande CL avec la
Suite ........ : syntaxe de la fonction QCMDEXC()
Exemple ...... : CALL QSYS.QCMDEXC('CHGJOB INQMSGRPY(*DFT)',0000000022.00000)
Mots clefs ... : OS400
*****************************************************************}
Function OS400_QCMDEXC( CommandeCL: String ): String;
Var
  LongCommande: String;
Begin
  LongCommande := Format( '%.10d', [Length(CommandeCL)] );
  Result := 'CALL QSYS.QCMDEXC("' + CommandeCL + '",' + LongCommande + '.00000)';
End;


end.
