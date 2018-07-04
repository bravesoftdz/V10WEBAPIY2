{***********UNITE*************************************************
Auteur  ...... :                                                                            
Créé le ...... : 05/01/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : TIERS360 ()                                    
Mots clefs ... : TOF;TIERS360
*****************************************************************}
Unit Tiers360_tof ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,FE_Main,
{$else}
     eMul,MainEagl,
{$ENDIF}
{$ifdef GIGI}
  DicoAf,
{$ENDIF}
{$ifdef AFFAIRE}
  AffaireUtil,
{$ENDIF}

     Vierge,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,M3Fp,
     UTOF,Utob,HTB97, ParamDbg, HSysMenu,HPanel,EntGC,AglInit ;

Type
  TOF_TIERS360 = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose              ; override;
  Private
    stTiers,stAuxi,Action,SArgument : String;
    HTrad1,HTrad2,HTrad3,HTrad4,HTrad5,HTrad6 : THSystemMenu;
    numContacts,numAdresses,numActions,numPropositions,numPieces,numParc,numAffaires,numChainages,numInterventions : integer;
    TobContacts,TobAdresses ,TobActions,TobPropositions,TobPieces,TobParc,TobAffaires,TobChainages,TobInterventions : tob;
    stTable : array[1..6] of string;
    ListeAlim : array[1..6] of string;
    EtatRisque: string ;
    TobMere, TobTiers : tob;
    widthOri,heightOri,widthBase1Ori : integer ; { largeur et hauteur native de la fiche }
    ficheTiers : boolean;
    { variables cbs }
    numcbs : array[1..6] of integer;
    tablecbs : array[1..6] of string;
    CodeCBS : array[1..6] of string;
    ListeAlimCBS : array[1..6] of string;
    titrecbs : array[1..6] of string;
    TobCBS : array[1..6] of tob;

    Procedure BParam1OnClick( Sender : tObject);
    Procedure BParam2OnClick( Sender : tObject);
    Procedure BParam3OnClick( Sender : tObject);
    Procedure BParam4OnClick( Sender : tObject);
    Procedure BParam5OnClick( Sender : tObject);
    Procedure BParam6OnClick( Sender : tObject);
    Procedure BListe1OnClick( Sender : tObject);
    Procedure BListe2OnClick( Sender : tObject);
    Procedure BListe3OnClick( Sender : tObject);
    Procedure BListe4OnClick( Sender : tObject);
    Procedure BListe5OnClick( Sender : tObject);
    Procedure BListe6OnClick( Sender : tObject);
    Procedure DoClickParam( numGrid : integer );
    procedure ResizeGrid ( numGrid : integer );
    procedure AlimTobs(numOrdre : integer);
    Procedure Grid1OnDblClick( Sender : tObject);
    Procedure Grid2OnDblClick( Sender : tObject);
    Procedure Grid3OnDblClick( Sender : tObject);
    Procedure Grid4OnDblClick( Sender : tObject);
    Procedure Grid5OnDblClick( Sender : tObject);
    Procedure Grid6OnDblClick( Sender : tObject);
    Procedure BFerme1OnClick( Sender : tObject);
    Procedure BFerme2OnClick( Sender : tObject);
    Procedure BFerme3OnClick( Sender : tObject);
    Procedure BFerme4OnClick( Sender : tObject);
    Procedure BFerme5OnClick( Sender : tObject);
    Procedure BFerme6OnClick( Sender : tObject);
    Procedure BValide1OnClick( Sender : tObject);
    Procedure BValide2OnClick( Sender : tObject);
    Procedure BValide3OnClick( Sender : tObject);
    Procedure BValide4OnClick( Sender : tObject);
    Procedure BValide5OnClick( Sender : tObject);
    Procedure BValide6OnClick( Sender : tObject);

    Procedure ClickListe(numOrdre : integer);
    Procedure BPARAMVIGN_OnClick(Sender : tObject);
    Procedure BTAILLE_OnClick(Sender : tObject);
    Procedure BTIERS_OnClick(Sender : tObject);
    Procedure ClickGrid(numGrid : integer);
    Procedure AffecteClick;
    Function RecupCol(Grid : THGrid) : String;
    procedure RecupLesPosVignettes;
    procedure CreatEtAfficheMere;
    Procedure AlimTobContacts ;
    Procedure AlimTobActions ;
    Procedure AlimTobPropositions ;
    Procedure AlimTobAdresses ;
    Procedure AlimTobPieces ;
    Procedure AlimTobChainages ;
{$IFDEF SAV}
    Procedure AlimTobParc ;
    Procedure AlimTobInterventions ;
{$ENDIF}
{$IFDEF AFFAIRE}
    Procedure AlimTobAffaires ;
    Procedure AffecteGridAffaire(Grid : THGrid);
{$ENDIF AFFAIRE}
    Procedure AlimTobCBS(i : integer) ;
    procedure CacheLesDates ( numGrid : integer; Montre : boolean) ;
    Procedure AppliqueDate(numOrdre : integer);
    procedure RecupDates ( numGrid : integer; CodeT : String) ;
    procedure SauveTailleEcran;
    procedure RecupTailleEcran;
    Function CBSRecupSql(CodeCBS, CodeTiers, CodeAuxiliaire,DateDebut,DateFin : String) : string;
    Function CBSRecupTable(CodeCBS : String) : string;
    Function CBSClickListe(CodeCBS, CodeTiers, CodeAuxiliaire,stArg : String)  : string;
    Function CBSClickGrid(CodeCBS,valeur,stArg : String)  : string;
    Function CBSTestChamp(CodeCBS : String)  : string;
    Function CBSRecupCode(number : integer) : string;
    Function CBSCacherDate(CodeCBS : String) : boolean;
    Function CBSRecupListe(CodeCBS : String)  : string;
    function PosVign(Code : String; Liste : array of string) : integer;
    Procedure FEUOnClick( Sender : tObject);
    Procedure VALIDETOUTOnClick( Sender : tObject);
    Procedure RefreshTiers;
  end ;

  const
  nb_vignettes : integer = 6;
  ListeAlimContacts : String = 'RTCONTACTS360';
  ListeAlimAdresses : String = 'RTADRESSES360';
  ListeAlimActions : String = 'RTACTIONS360';
  ListeAlimPropositions : String = 'RTPROPOSITIONS360';
  ListeAlimPieces : String = 'RTPIECES360';
  ListeAlimParc : String = 'RTPARC360';
  ListeAlimAffaires : String = 'RTAFFAIRE360';
  ListeAlimChainages : String = 'RTCHAINAGES360';
  ListeAlimInterventions : String = 'RTINTERVENTION360';
  TitreContacts: String = 'Contacts';
  TitreAdresses : String = 'Adresses';
  TitreActions : String = 'Actions';
  TitrePropositions : String = 'Propositions';
  TitrePieces : String = 'Pièces';
  TitreParc : String = 'Parc';
  TitreAffaires : String = 'Affaires';
  TitreChainages : String = 'Chaînages';
  TitreInterventions : String = 'Interventions';
  TableC : String = 'CONTACT';
  TableADR : String = 'ADRESSES';
  TableRAC : String = 'ACTIONS';
  TableRPE : String = 'PERSPECTIVES';
  TableGP : String = 'PIECE';
  TableWPC : String = 'WPARCNOME';
  TableAFF : String = 'AFFAIRE';
  TableWIV : String = 'WINTERVENTION';
  TableRCH : String = 'ACTIONSCHAINEES';
  ListeChampsDate : String = 'DATE;LDATE';
	PositionOriVign: array[1..6] of string 	= (
          {1}        'C'
          {2}       ,'A'
          {3}       ,'R'
          {4}       ,'P'
          {5}       ,'G'
          {6}       ,'W'
                     );

Implementation
uses
  UtilPgi,UtilRT
{$IFDEF SAV}
  ,wparc,wparcnome
{$ENDIF}
{$IFDEF AFFAIRE}
  ,ConfidentAffaire
{$ENDIF AFFAIRE}
  ,FactUtil,GCOuvertureGEP,TiersUtil,UtilConfid;

procedure TOF_TIERS360.OnArgument (S : String ) ;
begin
  Inherited ;
  RecupTailleEcran;

  AffecteClick;

  stAuxi := GetArgumentValue(S, 'AUXILIAIRE');
  stTiers := GetArgumentValue(S, 'TIERS');
  Action := GetArgumentValue(S, 'ACTION');
  ficheTiers:=false;
  if Pos('FICHETIERS',S) <> 0 then
    begin
    ficheTiers:=true;
    if Assigned(GetControl('BTIERS')) then
      SetControlVisible('BTIERS',false);
    end;

  TobTiers:=nil;
  SArgument:=S;
  TobTiers := Tob(GetArgumentInteger(S, 'TOBTIERS'   ));
  if not Assigned(TobTiers) then
    TobTiers:=TOB.Create ('TIERS', Nil, -1);
  RefreshTiers ;
  RecupLesPosVignettes;
  if Assigned(GetControl('FEUVERT')) then
     TToolBarButton97(GetControl('FEUVERT')).OnClick := FEUOnClick;
  if Assigned(GetControl('FEUROUGE')) then
     TToolBarButton97(GetControl('FEUROUGE')).OnClick := FEUOnClick;
  if Assigned(GetControl('FEUORANGE')) then
     TToolBarButton97(GetControl('FEUORANGE')).OnClick := FEUOnClick;
  if Assigned(GetControl('VALIDETOUT')) then
     TToolBarButton97(GetControl('VALIDETOUT')).OnClick := VALIDETOUTOnClick;
  AppliquerConfidentialite(Ecran,TobTiers.GetString('T_NATUREAUXI'));
end;

Procedure TOF_TIERS360.AlimTobContacts ;
var sql : string;
begin
  if Assigned(TobContacts) then exit;
  sql:= 'SELECT * FROM CONTACT WHERE C_TYPECONTACT="T" AND C_AUXILIAIRE = "'+stAuxi+'" AND C_FERME <> "X"';
  TobContacts:=TOB.Create (TableC, nil, -1);
  try
    TobContacts.LoadDetailDBFromSQL(TableC, sql);
  finally
  end;
end;

Procedure TOF_TIERS360.AlimTobAdresses ;
var sql : string;
begin
  if Assigned(TobAdresses) then exit;
  sql:= 'SELECT * from ADRESSES Where ADR_TYPEADRESSE="TIE" AND ADR_REFCODE = "'+stTiers+'"';
  TobAdresses:=TOB.Create (TableADR, nil, -1);
  try
    TobAdresses.LoadDetailDBFromSQL(TableADR, sql) ;
  finally
  end;
end;

Procedure TOF_TIERS360.AlimTobActions ;
var sql : string;
begin
  if Assigned(TobActions) then exit;
  sql:= 'SELECT * FROM ACTIONS WHERE RAC_AUXILIAIRE = "'+stAuxi+'" AND RAC_DATEACTION <= "'+
      UsDateTime(StrToDate(GetControlText('DATE'+IntToStr(numActions)+'_')))+
      '" AND RAC_DATEACTION >= "'+UsDateTime(StrToDate(GetControlText('DATE'+IntToStr(numActions))))+'"' ;
  TobActions:=TOB.Create (TableRAC, nil, -1);
  try
    TobActions.LoadDetailDBFromSQL(TableRAC, sql);
  finally
  end;
end;

Procedure TOF_TIERS360.AlimTobChainages ;
var sql : string;
begin
  if Assigned(TobChainages) then exit;
  sql:= 'SELECT * FROM ACTIONSCHAINEES WHERE RCH_TIERS= "'+stTiers+'" AND RCH_DATEDEBUT <= "'+
      UsDateTime(StrToDate(GetControlText('DATE'+IntToStr(numChainages)+'_')))+
      '" AND RCH_DATEDEBUT >= "'+UsDateTime(StrToDate(GetControlText('DATE'+IntToStr(numChainages))))+
      '" AND RCH_TERMINE = "-"' ;
  TobChainages:=TOB.Create (TableRCH, nil, -1);
  try
    TobChainages.LoadDetailDBFromSQL(TableRCH, sql);
  finally
  end;
end;

Procedure TOF_TIERS360.AlimTobPropositions ;
var sql : string;
begin
  if Assigned(TobPropositions) then exit;
  sql:= 'SELECT * FROM PERSPECTIVES WHERE RPE_AUXILIAIRE = "'+stAuxi+'" AND RPE_DATEREALISE <= "'+
  UsDateTime(StrToDate(GetControlText('DATE'+IntToStr(numPropositions)+'_')))+
  '" AND RPE_DATEREALISE >= "'+UsDateTime(StrToDate(GetControlText('DATE'+IntToStr(numPropositions))))+'"' ;
  TobPropositions:=TOB.Create (TableRPE, nil, -1);
  try
    TobPropositions.LoadDetailDBFromSQL(TableRPE, sql);
  finally
  end;
end;

Procedure TOF_TIERS360.AlimTobPieces;
var sql : string;
begin
  if Assigned(TobPieces) then exit;
  sql:= 'SELECT * FROM PIECE WHERE GP_TIERS = "'+stTiers+'" AND GP_DATEPIECE <= "'+
      UsDateTime(StrToDate(GetControlText('DATE'+IntToStr(numPieces)+'_')))+
      '" AND GP_DATEPIECE >= "'+UsDateTime(StrToDate(GetControlText('DATE'+IntToStr(numPieces))))+'"'
      + ' AND GP_NATUREPIECEG <>"'+VH_GC.AFNatAffaire+'" AND GP_NATUREPIECEG <>"'+VH_GC.AFNatProposition+'"'  //mcd 04/03/07 ne pas prendre en compte peice/affaire et propositon affaire
       ;
  TobPieces:=TOB.Create (TableGP, nil, -1);
  try
    TobPieces.LoadDetailDBFromSQL(TableGP, sql) ;
  finally
  end;
end;
{$IFDEF SAV}
Procedure TOF_TIERS360.AlimTobParc ;
var sql : string;
begin
  if Assigned(TobParc) then exit;
  sql:= 'SELECT * FROM WPARC LEFT JOIN WPARCNOME ON WPC_IDENTIFIANT=WPN_IDENTPARC WHERE WPC_TIERS = "'+stTiers+'"'+
    ' AND WPC_ETATPARC="ES" AND WPN_ETATPARC="ES"';
  TobParc:=TOB.Create ('--Parc--', nil, -1);
  try
    TobParc.LoadDetailDBFromSQL('--Parc--', sql);
  finally
  end;
end;

Procedure TOF_TIERS360.AlimTobInterventions ;
var sql : string;
begin
  if Assigned(TobInterventions) then exit;
  sql:= 'SELECT * FROM '+TableWIV+' WHERE WIV_TIERS = "'+stTiers+'"'+
    ' AND WIV_ETATINTERV="ENC" AND WIV_DATEDEBUT <= "'+
      UsDateTime(StrToDate(GetControlText('DATE'+IntToStr(numInterventions)+'_')))+
      '" AND WIV_DATEDEBUT >= "'+UsDateTime(StrToDate(GetControlText('DATE'+IntToStr(numInterventions))))+'"';
  TobInterventions:=TOB.Create ('--Interventions--', nil, -1);
  try
    TobInterventions.LoadDetailDBFromSQL('--Interventions--', sql);
  finally
  end;
end;

{$ENDIF}
{$IFDEF AFFAIRE}
Procedure TOF_TIERS360.AlimTobAffaires ;
var sql : string;
begin
  if Assigned(TobAffaires) then exit;
  sql:= 'SELECT * FROM AFFAIRE WHERE AFF_TIERS = "'+stTiers+'"'+' AND AFF_DATEDEBUT <= "'+
  UsDateTime(StrToDate(GetControlText('DATE'+IntToStr(numAffaires)+'_')))+
  '" AND AFF_DATEDEBUT >= "'+UsDateTime(StrToDate(GetControlText('DATE'+IntToStr(numAffaires))))+'"' ;
  TobAffaires:=TOB.Create (TableAFF, nil, -1);
  try
    TobAffaires.LoadDetailDBFromSQL(TableAFF, sql);
  finally
  end;
end;
{$ENDIF AFFAIRE}

Procedure TOF_TIERS360.AlimTobCBS(i : integer) ;
var sql,datedebut,datefin : string;
begin
  if Assigned(TobCBS[i]) then exit;
  datedebut:='';
  datefin:='';
  if CBSCacherDate(CodeCBS[i]) then
     begin
     datedebut:=GetControlText('DATE'+IntToStr(numcbs[i]));
     datefin:=GetControlText('DATE'+IntToStr(numcbs[i])+'_');
     end;
  sql:= CBSRecupSql(CodeCBS[i],stTiers,stAuxi,datedebut,datefin);
  if sql = '' then exit;
  TobCBS[i]:=TOB.Create (tablecbs[i], nil, -1);
  try
    TobCBS[i].LoadDetailDBFromSQL(tablecbs[i], sql);
  finally
  end;
end;

{$IFDEF AFFAIRE}
Procedure TOF_TIERS360.AffecteGridAffaire (Grid:ThGrid) ;
var ii : integer;
begin   //mcd 02/05/07 les colonnes affaire doivent être cachée ne fct du paramétrage
  //fait en similaire des fct de AffaireUtil : ModifColAffaireGrid
  for ii:=0 to Grid.ColCount-1 Do
    if Grid.FColNames[ii] = 'AFF_AFFAIRE' then
    begin
      Grid.ColLengths[ii]:=-1;//colonne pas visible
      Grid.ColWidths[ii] := -1;
    end
    else if Grid.FColNames[ii] = 'AFF_AFFAIRE1' then
    begin
     Grid.cells[ii,0] := VH_GC.CleAffaire.Co1Lib; //libellé personnalisé de la colonne
    end
    else if Grid.FColNames[ii] = 'AFF_AFFAIRE2' then
    begin
     if not VH_GC.CleAffaire.Co2Visible then
     begin
      Grid.ColLengths[ii]:=-1; //si pas visible
      Grid.ColWidths[ii] := -1;
     end;
     Grid.cells[ii,0] := VH_GC.CleAffaire.Co2Lib; //libellé personnalisé de la colonne
     if (VH_GC.AFFORMATEXER <> 'AUC')  then
       Grid.ColFormats[ii]:=FormatZoneExercice(VH_GC.AFFORMATEXER, false); //pour format dans grid mais non OK ???
    end
    else if Grid.FColNames[ii] = 'AFF_AFFAIRE3' then
    begin
     if not VH_GC.CleAffaire.Co3Visible then
     begin
      Grid.ColLengths[ii]:=-1; //si pas visible
      Grid.ColWidths[ii] := -1;
     end;
     Grid.cells[ii,0] := VH_GC.CleAffaire.Co3Lib; //libellé personnalisé de la colonne
    end
    else if Grid.FColNames[ii] = 'AFF_AVENANT' then
    begin
      if not (VH_GC.CleAffaire.GestionAvenant) then
      begin
        Grid.ColLengths[ii]:=-1; //si pas visible
        Grid.ColWidths[ii] := -1;
      end;
    end
end;
{$ENDIF AFFAIRE}

Procedure TOF_TIERS360.BParam1OnClick( Sender : tObject);
begin
  DoClickParam(1);
end;

Procedure TOF_TIERS360.BParam2OnClick( Sender : tObject);
begin
  DoClickParam(2);
end;

Procedure TOF_TIERS360.BParam3OnClick( Sender : tObject);
begin
  DoClickParam(3);
end;

Procedure TOF_TIERS360.BParam4OnClick( Sender : tObject);
begin
  DoClickParam(4);
end;

Procedure TOF_TIERS360.BParam5OnClick( Sender : tObject);
begin
  DoClickParam(5);
end;

Procedure TOF_TIERS360.BParam6OnClick( Sender : tObject);
begin
  DoClickParam(6);
end;

Procedure TOF_TIERS360.FEUOnClick( Sender : tObject);
begin
  CalculSoldesAuxi(TOBTiers.GetValue('T_AUXILIAIRE')) ;
  TobTiers.LoadDB(True) ; // Pour recharger les données calculées
  TheTob:=TobTiers;
  AglLanceFiche('GC', 'GCENCOURS','','','ACTION=CONSULTATION');
end;

Procedure TOF_TIERS360.VALIDETOUTOnClick( Sender : tObject);
var i : integer;
begin
  TobTiers.InitValeurs;
  RefreshTiers;
  if Assigned(TobActions) then
      FreeAndNil(TobActions);
  if Assigned(TobChainages) then
      FreeAndNil(TobChainages);
  if Assigned(TobPropositions) then
    FreeAndNil(TobPropositions);
  if Assigned(TobPieces) then
    FreeAndNil(TobPieces);
  if Assigned(TobContacts) then
      FreeAndNil(TobContacts);
  if Assigned(TobAdresses) then
    FreeAndNil(TobAdresses);
  if Assigned(TobParc) then
    FreeAndNil(TobParc);
  if Assigned(TobInterventions) then
    FreeAndNil(TobInterventions);
  for i:=1 to nb_vignettes Do
    if Assigned(TobCBS[i]) then
      FreeAndNil(TobCBS[i]);
  RecupLesPosVignettes;
end;

Procedure TOF_TIERS360.BFerme6OnClick( Sender : tObject);
begin
  THPanel(GetControl('P6')).Visible:=false;
  SaveSynRegKey('RTTiersVign6','X',TRUE);
end;

Procedure TOF_TIERS360.BFerme5OnClick( Sender : tObject);
begin
  THPanel(GetControl('P5')).Visible:=false;
  SaveSynRegKey('RTTiersVign5','X',TRUE);
end;

Procedure TOF_TIERS360.BFerme4OnClick( Sender : tObject);
begin
  THPanel(GetControl('P4')).Visible:=false;
  SaveSynRegKey('RTTiersVign4','X',TRUE);
end;

Procedure TOF_TIERS360.BFerme3OnClick( Sender : tObject);
begin
  THPanel(GetControl('P3')).Visible:=false;
  SaveSynRegKey('RTTiersVign3','X',TRUE);
end;

Procedure TOF_TIERS360.BFerme2OnClick( Sender : tObject);
begin
  THPanel(GetControl('P2')).Visible:=false;
  SaveSynRegKey('RTTiersVign2','X',TRUE);
end;

Procedure TOF_TIERS360.BFerme1OnClick( Sender : tObject);
begin
  THPanel(GetControl('P1')).Visible:=false;
  SaveSynRegKey('RTTiersVign1','X',TRUE);
end;

Procedure TOF_TIERS360.BValide1OnClick( Sender : tObject);
begin
  AppliqueDate(1);
end;
Procedure TOF_TIERS360.BValide2OnClick( Sender : tObject);
begin
  AppliqueDate(2);
end;
Procedure TOF_TIERS360.BValide3OnClick( Sender : tObject);
begin
  AppliqueDate(3);
end;
Procedure TOF_TIERS360.BValide4OnClick( Sender : tObject);
begin
  AppliqueDate(4);
end;
Procedure TOF_TIERS360.BValide5OnClick( Sender : tObject);
begin
  AppliqueDate(5);
end;
Procedure TOF_TIERS360.BValide6OnClick( Sender : tObject);
begin
  AppliqueDate(6);
end;



Procedure TOF_TIERS360.Grid1OnDblClick( Sender : tObject);
begin
  ClickGrid(1);
end;

Procedure TOF_TIERS360.Grid2OnDblClick( Sender : tObject);
begin
  ClickGrid(2);
end;

Procedure TOF_TIERS360.Grid3OnDblClick( Sender : tObject);
begin
  ClickGrid(3);
end;

Procedure TOF_TIERS360.Grid4OnDblClick( Sender : tObject);
begin
  ClickGrid(4);
end;

Procedure TOF_TIERS360.Grid5OnDblClick( Sender : tObject);
begin
  ClickGrid(5);
end;

Procedure TOF_TIERS360.Grid6OnDblClick( Sender : tObject);
begin
  ClickGrid(6);
end;

Procedure TOF_TIERS360.BListe1OnClick( Sender : tObject);
begin
  ClickListe(1);
end;

Procedure TOF_TIERS360.BListe2OnClick( Sender : tObject);
begin
  ClickListe(2);
end;

Procedure TOF_TIERS360.BListe3OnClick( Sender : tObject);
begin
  ClickListe(3);
end;

Procedure TOF_TIERS360.BListe4OnClick( Sender : tObject);
begin
  ClickListe(4);
end;

Procedure TOF_TIERS360.BListe5OnClick( Sender : tObject);
begin
  ClickListe(5);
end;

Procedure TOF_TIERS360.BListe6OnClick( Sender : tObject);
begin
  ClickListe(6);
end;

Procedure TOF_TIERS360.AppliqueDate(numOrdre : integer);
var i : integer;
begin
  if ListeAlim[numOrdre] =ListeAlimActions then
    begin
    FreeAndNil(TobActions);
    AlimTobActions;
    SaveSynRegKey('RTTiersDateVR',GetControlText('DATE'+IntToStr(numOrdre))+';'+
      GetControlText('DATE'+IntToStr(numOrdre)+'_'),TRUE);
    end
  else if ListeAlim[numOrdre] =ListeAlimChainages then
    begin
    FreeAndNil(TobChainages);
    AlimTobChainages;
    SaveSynRegKey('RTTiersDateVH',GetControlText('DATE'+IntToStr(numOrdre))+';'+
      GetControlText('DATE'+IntToStr(numOrdre)+'_'),TRUE);
    end
  else if ListeAlim[numOrdre] =ListeAlimPropositions then
    begin
    FreeAndNil(TobPropositions);
    AlimTobPropositions;
    SaveSynRegKey('RTTiersDateVP',GetControlText('DATE'+IntToStr(numOrdre))+';'+
      GetControlText('DATE'+IntToStr(numOrdre)+'_'),TRUE);
    end
  else if ListeAlim[numOrdre] =ListeAlimPieces then
    begin
    FreeAndNil(TobPieces);
    AlimTobPieces;
    SaveSynRegKey('RTTiersDateVG',GetControlText('DATE'+IntToStr(numOrdre))+';'+
      GetControlText('DATE'+IntToStr(numOrdre)+'_'),TRUE);
    end
{$IFDEF AFFAIRE}
  else if ListeAlim[numOrdre] =ListeAlimAffaires then
    begin
    FreeAndNil(TobAffaires);
    AlimTobAffaires;
    SaveSynRegKey('RTTiersDateVF',GetControlText('DATE'+IntToStr(numOrdre))+';'+
      GetControlText('DATE'+IntToStr(numOrdre)+'_'),TRUE);
    end
{$ENDIF AFFAIRE}
{$IFDEF SAV}
  else if ListeAlim[numOrdre] =ListeAlimInterventions then
    begin
    FreeAndNil(TobInterventions);
    AlimTobInterventions;
    SaveSynRegKey('RTTiersDateVI',GetControlText('DATE'+IntToStr(numOrdre))+';'+
      GetControlText('DATE'+IntToStr(numOrdre)+'_'),TRUE);
    end
{$ENDIF AFFAIRE}
  else
  For i:=1 to nb_vignettes Do
    if ListeAlim[numOrdre] =ListeAlimCBS[i] then
      begin
      FreeAndNil(TobCBS[i]);
      AlimTobCBS(i);
      SaveSynRegKey('RTTiersDateVI',GetControlText('DATE'+IntToStr(numOrdre))+';'+
        GetControlText('DATE'+IntToStr(numOrdre)+'_'),TRUE);
      break;
      end
  ;

  TobMere.Detail[numOrdre-1].free;
  AlimTobs(numOrdre);
  TobMere.Detail[numOrdre-1].PutGridDetailOnListe(THGrid(GetControl('GRID'+IntToStr(numOrdre))), ListeAlim[numOrdre]);
  ResizeGrid(numOrdre);
end;

Procedure TOF_TIERS360.ClickListe(numOrdre : integer);
var stArg : String;
    i : integer;
begin
  if Action = 'CONSULTATION' then
    stArg:='ACTION=CONSULTATION;'
  else
    stArg:='ACTION=MODIFICATION;';
  if ListeAlim[numOrdre] =ListeAlimContacts then
    begin
        //mcd 28/05/07 14139 ?? à faire idem GC ???  if not ((ctxAffaire in V_PGI.PGIContexte) or (ctxScot in V_PGI.PGIContexte)) then
      AglLanceFiche('YY','YYCONTACT','T;'+stAuxi,'',stArg+'TYPE=T;'+'TYPE2='+GetControlText('T_NATUREAUXI')+';PART='+GetControlText('T_PARTICULIER')+';TITRE='+GetControlText ('T_LIBELLE')+';TIERS='+stTiers+';ALLCONTACT' );
    end
  else if ListeAlim[numOrdre] =ListeAlimAdresses then
          AglLanceFiche('GC','GCADRESSES','ADR_TYPEADRESSE=TIE;ADR_NATUREAUXI='+GetControlText('T_NATUREAUXI')+';ADR_REFCODE='+stTiers,'',
          stArg+'TYPEADRESSE=TIE;TITRE='+GetControlText('T_LIBELLE')+';PART='+GetControlText('T_PARTICULIER')+';CLI='+stTiers+';TIERS='+stTiers+';NATUREAUXI='+GetControlText('T_NATUREAUXI'))
  else if ListeAlim[numOrdre] =ListeAlimActions then
    AglLanceFiche('RT','RTACTIONS_TIERS','RAC_TIERS='+stTiers,'',stArg+'NOCHANGEPROSPECT;PREVU')
  else if ListeAlim[numOrdre] =ListeAlimChainages then
    AglLanceFiche('RT','RTCHAINAGES_TIERS','RCH_TIERS='+stTiers,'',stArg+'NOCHANGEPROSPECT;TIERS='+stTiers)
  else if ListeAlim[numOrdre] =ListeAlimPropositions then
    AglLanceFiche('RT','RTPERSP_MUL_TIERS','RPE_TIERS='+stTiers,'',stArg+'NOCHANGEPROSPECT')
  else if ListeAlim[numOrdre] =ListeAlimPieces then
    begin
    if (ctxAffaire in V_PGI.PGIContexte) or (ctxScot in V_PGI.PGIContexte) then
      AglLanceFiche('AFF','AFPIECECOURS_TIE', 'GP_TIERS=' +stTiers,'',stArg+'NATUREAUXI=CLI')
    else
      AglLanceFiche('GC','GCPIECECOURS_TIER', 'GP_TIERS=' +stTiers,'',stArg);
    end
  else if ListeAlim[numOrdre] =ListeAlimParc then
    AglLanceFiche ('W', 'WPARCELEM_MUL', 'WPC_TIERS=' + stTiers, '',stArg+'WPC_TIERS')
  else if ListeAlim[numOrdre] =ListeAlimInterventions then
    AglLanceFiche ('W', 'WINTERVMUL_TIERS', 'WIV_TIERS='+stTiers, '',stArg+'FICHETIERS;WIV_TIERS;')
{$IFDEF AFFAIRE}
  else if ListeAlim[numOrdre] =ListeAlimAffaires then
    begin
    If AGLJaiLeDroitFiche (['AFFAIRE',stArg,'PRO'],3)  then
    AGLLanceFiche('AFF','AFFAIRE_MUL','AFF_STATUTAFFAIRE=AFF;AFF_TIERS=' +stTiers,'','STATUT=AFF;NOFILTRE;NOCHANGESTATUT;'+stArg+
      'AFF_TIERS=' +stTiers);
    end
{$ENDIF AFFAIRE}
  else
    for i:=1 to nb_vignettes Do
      if ListeAlim[numOrdre] =ListeAlimCBS[i] then
        begin
        CBSClickListe(CodeCBS[i],stTiers,stAuxi,stArg);
        break;
        end;
  ;
end;

Procedure TOF_TIERS360.ClickGrid(numGrid : integer);
var tobgrid : tob;
    Grid : THGrid;
{$IFDEF SAV}
    GetCleWPC : tCleWPC;
{$ENDIF}
    colonnes,stmess : string;
  stArg : String;
  i : integer ;
begin
  if Action = 'CONSULTATION' then
    stArg:='ACTION=CONSULTATION'
  else
    stArg:='ACTION=MODIFICATION';
  tobgrid:=TOB.Create (stTable[numGrid], TobMere, numGrid-1);
  Grid:=THGrid(GetControl('GRID'+IntToStr(numGrid)));
  colonnes:=RecupCol(Grid);
  tobgrid.GetLigneGrid(Grid,Grid.Row,colonnes) ;
  if tobgrid.FieldExists('C_AUXILIAIRE') then
    begin
    if tobgrid.GetInteger('C_NUMEROCONTACT') =0 then
       PgiInfo ('Vous devez afficher la colonne n° contact pour accéder à la fiche','Contact')
    else
      AglLanceFiche('YY','YYCONTACT','T;'+stAuxi,tobgrid.GetString('C_NUMEROCONTACT'),stArg+';ALLCONTACT');
    end
  else if tobgrid.FieldExists('ADR_TYPEADRESSE') then
    begin
    if tobgrid.GetInteger('ADR_NUMEROADRESSE') =0 then
       PgiInfo ('Vous devez afficher la colonne n° adresse pour accéder à la fiche','Adresse')
    else
      AglLanceFiche('GC','GCADRESSES','TIE;'+tobgrid.GetString('ADR_NUMEROADRESSE'),'',stArg+
           ';TIERS='+tobgrid.GetString('ADR_REFCODE'));
    end
  else if tobgrid.FieldExists('RAC_AUXILIAIRE') then
    begin
    if tobgrid.GetInteger('RAC_NUMACTION') =0 then
       PgiInfo ('Vous devez afficher la colonne n° action pour accéder à la fiche','Action')
    else
      AglLanceFiche('RT','RTACTIONS','',stAuxi+';'+tobgrid.GetString('RAC_NUMACTION'),stArg+';NOCHANGEPROSPECT;FICHETIERS;RAC_TIERS='+stTiers);
    end
  else if tobgrid.FieldExists('RCH_NUMERO') then
    begin
    if tobgrid.GetInteger('RCH_NUMERO') =0 then
       PgiInfo ('Vous devez afficher la colonne n° chaînage pour accéder à la fiche','Chaînage')
    else
      AGLLanceFiche('RT','RTCHAINAGES','','','GRC;ORIGINETIERS;RCH_NUMERO='+tobgrid.getString('RCH_NUMERO'));
    end
  else if tobgrid.FieldExists('RPE_PERSPECTIVE') then
    begin
    if tobgrid.GetInteger('RPE_PERSPECTIVE') =0 then
       PgiInfo ('Vous devez afficher la colonne n° proposition pour accéder à la fiche','Proposition')
    else
      AglLanceFiche('RT','RTPERSPECTIVES','',tobgrid.GetString('RPE_PERSPECTIVE'),stArg+';NOCHANGEPROSPECT;FICHETIERS;RAC_TIERS='+stTiers);
    end
  else if tobgrid.FieldExists('GP_NUMERO') then
     begin
     if (tobgrid.GetInteger('GP_NUMERO') =0 ) or (Pos('GP_INDICEG',colonnes) =0 )
       or (tobgrid.GetString('GP_NATUREPIECEG') ='' ) or (tobgrid.GetString('GP_DATEPIECE') ='' )
       or (tobgrid.GetString('GP_SOUCHE') ='' ) then
       PgiInfo ('Vous devez afficher les éléments pièce pour accéder à la pièce','Pièce')
     else
       AppelPiece([tobgrid.GetString('GP_NATUREPIECEG') + ';' + tobgrid.GetString('GP_DATEPIECE') + ';' +tobgrid.GetString('GP_SOUCHE') + ';' +
             tobgrid.GetString('GP_NUMERO') + ';' +tobgrid.GetString('GP_INDICEG'),stArg],1);
     end
  else if tobgrid.FieldExists('WPN_IDENTIFIANT') then
          begin
          if (tobgrid.GetInteger('WPN_IDENTIFIANT') = 0 ) or ( tobgrid.GetInteger('WPC_IDENTIFIANT') = 0 ) then
             PgiInfo ('Vous devez afficher les colonnes identifiant pour accéder à la fiche','Parc')
          else
            begin
{$IFDEF SAV}
            GetCleWPC.Identifiant := tobgrid.GetInteger ('WPC_IDENTIFIANT');
            stArg:=Action+';WPN_IDENT='+tobgrid.GetString ('WPN_IDENTIFIANT');
            if ficheTiers then stArg:=stArg+';FICHETIERS';
            wCallTreeViewWPN(GetCleWPC, stArg);
{$ENDIF}
            end;
          end
  else if tobgrid.FieldExists('WIV_IDENTIFIANT') then
          begin
          if (tobgrid.GetInteger('WIV_IDENTIFIANT') = 0 ) then
             PgiInfo ('Vous devez afficher la colonne identifiant pour accéder à la fiche','Interventions')
          else
            begin
{$IFDEF SAV}
            AGLLanceFiche('W','WINTERVENTION','', tobgrid.GetString('WIV_IDENTIFIANT'),stArg+'MONOFICHE;PRODUITPGI=GRC') ;
{$ENDIF}
            end;
          end

{$IFDEF AFFAIRE}
  else if tobgrid.FieldExists('AFF_AFFAIRE') then
    begin
    if tobgrid.GetString('AFF_AFFAIRE') = '' then
       PgiInfo ('Vous devez afficher la colonne code affaire pour accéder à la fiche','Affaire')
    else
      begin
      //mcd 02/05/2007 pas la même fiche en GI   , on passe par fct std
      /// AGLLanceFiche('AFF','AFFAIRE','',tobgrid.GetString('AFF_AFFAIRE'),stArg);
        if Action = 'CONSULTATION' then
         V_PGI.DispatchTT(5, taConsult, tobgrid.GetString('AFF_AFFAIRE'), '', 'MONOFICHE')
        else V_PGI.DispatchTT(5, taModif, tobgrid.GetString('AFF_AFFAIRE'), '', 'MONOFICHE')
      end;
    end
{$ENDIF AFFAIRE}
  else
  for i := 1 to nb_vignettes Do
    if (CodeCBS[i] <> '') then
      begin
      if tobgrid.FieldExists(CBSTestChamp(CodeCBS[i])) then
        if tobgrid.GetString(CBSTestChamp(CodeCBS[i])) ='' then
           begin
           stmess:='Vous devez afficher le champ '+CBSTestChamp(CodeCBS[i])+' pour accéder à la fiche';
           PgiInfo (stmess,titrecbs[i]);
           end
        else
          begin
          CBSClickGrid(CodeCBS[i],tobgrid.GetString(CBSTestChamp(CodeCBS[i])),stArg);
          break;
          end;
      end
  ;

  FreeAndNil (tobgrid);
end;

Function TOF_TIERS360.RecupCol(Grid : THGrid) : String;
var i : integer;
begin
  Result:='';
  for i:=0 to Grid.ColCount-1 Do
    if Grid.FColNames[i] <> '' then
      Result := Result + Grid.FColNames[i] +';'
    else
      break;
end;

Procedure TOF_TIERS360.DoClickParam(numGrid : integer);
begin
  try
    {$IFDEF EAGLCLIENT}
      ParamListe (ListeAlim[numGrid], nil, '') ;
    {$ELSE}
      ParamListe (ListeAlim[numGrid], nil, nil, '');
    {$ENDIF}
  finally
    THGrid(GetControl('GRID'+IntToStr(numGrid))).DisplayListe := '';
    TobMere.Detail[numGrid-1].PutGridDetailOnListe(THGrid(GetControl('GRID'+IntToStr(numGrid))), ListeAlim[numGrid]);
    ResizeGrid(numGrid);
  end;
end;

procedure TOF_TIERS360.ResizeGrid ( numGrid : integer) ;
var HTrad : THSystemMenu;
    Grid : THGrid;
begin
    HTrad:=nil;
    Case numGrid of
      1 : HTrad:=HTrad1;
      2 : HTrad:=HTrad2;
      3 : HTrad:=HTrad3;
      4 : HTrad:=HTrad4;
      5 : HTrad:=HTrad5;
      6 : HTrad:=HTrad6;
    end;
  Grid:=THGrid(GetControl('GRID'+IntToStr(numGrid)));
{$ifdef AFFAIRE} //mcd 02/05/2007
  if ListeAlim[numGrid] =ListeAlimAffaires then AffecteGridAffaire (Grid);
{$endif}
  if not Assigned(HTrad) then
    HTrad := THSystemMenu.Create (Grid.Parent);
  if (HTrad.ActiveResize) then
    HTrad.ResizeGridColumns(Grid);
end;

procedure TOF_TIERS360.OnClose;
var i : integer;
begin
  if (Tob(GetArgumentInteger(SArgument, 'TOBTIERS')) = nil) and Assigned(TobTiers) then
      FreeAndNil(TobTiers);
  if Assigned(TobMere) then
    FreeAndNil(TobMere);
  if Assigned(TobActions) then
      FreeAndNil(TobActions);
  if Assigned(TobChainages) then
      FreeAndNil(TobChainages);
  if Assigned(TobPropositions) then
    FreeAndNil(TobPropositions);
  if Assigned(TobPieces) then
    FreeAndNil(TobPieces);
  if Assigned(TobContacts) then
      FreeAndNil(TobContacts);
  if Assigned(TobAdresses) then
    FreeAndNil(TobAdresses);
  if Assigned(TobParc) then
    FreeAndNil(TobParc);
  if Assigned(TobInterventions) then
    FreeAndNil(TobInterventions);
  for i:=1 to nb_vignettes Do
    if Assigned(TobCBS[i]) then
      FreeAndNil(TobCBS[i]);
  SauveTailleEcran;
end;

function TOF_TIERS360.PosVign(Code : String; Liste : array of string) : integer;
var i :  integer;
begin
  result:=0;
  for i :=0 to Nb_Vignettes-1 Do
    if Liste[i]=Code then
      begin
      result:=i+1;
      exit;
      end;
end;

procedure TOF_TIERS360.RecupLesPosVignettes;
Var Liste : array[1..6] of string;
    i,j : integer;
    trouve : boolean;
    stMess : string;
begin
  for i :=1 to Nb_Vignettes Do
    Liste[i]:=GetSynRegKey('RTTiersVign'+IntToStr(i),PositionOriVign[i],TRUE);

  numContacts:=PosVign('C',Liste);
  numAdresses:=PosVign('A',Liste);
  numActions:=0; numPropositions:=0; numChainages:=0;
  if ctxGRC in V_PGI.PGIContexte then
    begin
    numActions:=PosVign('R',Liste);
    RecupDates(numActions,'R');
    numPropositions:=PosVign('P',Liste);
    RecupDates(numPropositions,'P');
    numChainages:=PosVign('H',Liste);
    RecupDates(numChainages,'H');
    end;
  numAffaires:=0;
{$IFDEF AFFAIRE}
  if (ctxAffaire in V_PGI.PGIContexte) or (ctxGcAff in V_PGI.PGIContexte) then
    begin
    numAffaires:=PosVign('F',Liste);
    RecupDates(numAffaires,'F');
    end;
{$ENDIF AFFAIRE}

  numPieces:=PosVign('G',Liste);
  RecupDates(numPieces,'G');
{$IFDEF SAV}
  numParc:=PosVign('W',Liste);
  numInterventions:=PosVign('I',Liste);
  if not VH_GC.SAVSeria then
{$ENDIF}
    begin
    numParc:=0;
    numInterventions:=0;
    end;
  { mng cbs }
  for i:=1 to nb_vignettes Do
    begin
    CodeCBS[i]:=CBSRecupCode(i);
    if CodeCBS[i] <> '' then
      begin
      numcbs[i]:=PosVign(CodeCBS[i],Liste);
      ListeAlimCBS[i]:=CBSRecupListe(CodeCBS[i]);
      tablecbs[i]:=CBSRecupTable(CodeCBS[i]);
      titrecbs[i]:=rechdom('RTTIERS360',CodeCBS[i],false)+' (CBS)';
      end;
    end;
  { s'il existe des demandes de vignette (liste) sans code associé par script, on le signale }
  for i:=1 to nb_vignettes Do
    if copy(Liste[i],1,1) = 'Z' then
      begin
      trouve:=false;
      for j:=1 to nb_vignettes Do
         if Liste[i] = CodeCBS[j] then
           begin
           trouve:=true;
           break;
           end;
      if trouve = false then
        begin
        stMess:='Le code vignette '+Liste[i]+' (Business Studio) est mémorisé en registrerie mais n''est pas présent dans le script de cette base';
        PgiInfo(stMess,'Code vignette CBS');
        end;
      end;

  for i:=1 to nb_vignettes Do
    ListeAlim[i]:='XXXXX';

  if numContacts <> 0 then
    begin
    ListeAlim[numContacts] :=ListeAlimContacts;
    AlimTobContacts;
    CacheLesDates(numContacts,false);
    end;
  if numAdresses <> 0 then
    begin
    ListeAlim[numAdresses] :=ListeAlimAdresses;
    AlimTobAdresses;
    CacheLesDates(numAdresses,false);
    end;
  if numActions <> 0 then
    begin
    ListeAlim[numActions] :=ListeAlimActions;
    AlimTobActions;
    CacheLesDates(numActions,true);
    end;
  if numChainages <> 0 then
    begin
    ListeAlim[numChainages] :=ListeAlimChainages;
    AlimTobChainages;
    CacheLesDates(numChainages,true);
    end;

  if numPropositions <> 0 then
    begin
    ListeAlim[numPropositions] :=ListeAlimPropositions;
    AlimTobPropositions;
    CacheLesDates(numPropositions,true);
    end;
  if numPieces <> 0 then
    begin
    ListeAlim[numPieces] :=ListeAlimPieces;
    AlimTobPieces;
    CacheLesDates(numPieces,true);
    end;
  if numParc <> 0 then
    begin
    ListeAlim[numParc] :=ListeAlimParc;
{$IFDEF SAV}
    AlimTobParc;
{$ENDIF}
    CacheLesDates(numParc,false);
    end;
  if numInterventions <> 0 then
    begin
    ListeAlim[numInterventions] :=ListeAlimInterventions;
{$IFDEF SAV}
    AlimTobInterventions;
{$ENDIF}
    CacheLesDates(numInterventions,true);
    end;

  if numAffaires <> 0 then
    begin
    ListeAlim[numAffaires] :=ListeAlimAffaires;
{$IFDEF AFFAIRE}
    AlimTobAffaires;
{$ENDIF}
    CacheLesDates(numAffaires,true);
    end;
  for i := 1 to nb_vignettes Do
    begin
    if numcbs[i] <> 0 then
      begin
      ListeAlim[numcbs[i]] :=ListeAlimCbs[i];
      AlimTobCBS(i);
      CacheLesDates(numcbs[i],CBSCacherDate(CodeCBS[i]));
      end;
    end;
  CreatEtAfficheMere;
end;

procedure TOF_TIERS360.RecupDates (numGrid : integer; CodeT : String) ;
var stDate : String;
begin
  if numGrid = 0 then exit;
  stDate:=GetSynRegKey('RTTiersDateV'+CodeT,DateToStr(DEBUTDEMOIS(Date))+';'+DateToStr(FINDEMOIS(Date)),TRUE);
  SetControlText('DATE'+IntToStr(numGrid),ReadToKenSt(stDate));
  SetControlText('DATE'+IntToStr(numGrid)+'_',ReadToKenSt(stDate));
end;

procedure TOF_TIERS360.CacheLesDates ( numGrid : integer; Montre : boolean) ;
var Liste, controle : String;
begin
  Liste:=ListeChampsDate;
  repeat
    controle:=ReadToKenSt(Liste);
    if controle <> '' then
      begin
      SetControlVisible (controle+IntToStr(numGrid),Montre);
      SetControlVisible (controle+IntToStr(numGrid)+'_',Montre);
      end;
  until controle = '';
  SetControlVisible ('BVALIDE'+IntToStr(numGrid),Montre);
end;

procedure TOF_TIERS360.CreatEtAfficheMere;
var i : integer;
begin
  FreeAndNil(TobMere);
  TobMere:=TOB.Create ('Mere des tob', Nil, -1);
  for i:= 1 to nb_vignettes Do
    AlimTobs(i);

  for i := 1 to nb_vignettes Do
    if ListeAlim[i]='XXXXX' then
      THPanel(GetControl('P'+IntToStr(i))).Visible:=false
    else
      if (i <= TobMere.Detail.Count) and (Assigned (TobMere.Detail[i-1])) then
        begin
        THPanel(GetControl('P'+IntToStr(i))).Visible:=true;
        TobMere.Detail[i-1].PutGridDetailOnListe(THGrid(GetControl('GRID'+IntToStr(i))), ListeAlim[i]);
        ResizeGrid(i);
        end;
end;

Procedure TOF_TIERS360.AlimTobs ( numOrdre : integer);
var TobFille : tob;
    i : integer;
begin
  if ListeAlim[numOrdre] ='XXXXX' then
     TOB.Create ('videmaisutile', TobMere, numOrdre-1)
  else
  if ListeAlim[numOrdre] =ListeAlimContacts then
    begin
    stTable[numOrdre]:=TableC;
    SetControlCaption ('TITRE'+IntToStr(numOrdre),TitreContacts);
    TobFille:=TOB.Create (TableC, TobMere, numOrdre-1);
    TobFille.Dupliquer(TobContacts, true, True);
    end
  else
  if ListeAlim[numOrdre] =ListeAlimAdresses then
    begin
    stTable[numOrdre]:=TableADR;
    SetControlCaption ('TITRE'+IntToStr(numOrdre),TitreAdresses);
    TobFille:=TOB.Create (TableADR, TobMere, numOrdre-1);
    TobFille.Dupliquer(TobAdresses, true, True);
    end
  else
  if ListeAlim[numOrdre] =ListeAlimActions then
    begin
    stTable[numOrdre]:=TableRAC;
    SetControlCaption ('TITRE'+IntToStr(numOrdre),TitreActions);
    TobFille:=TOB.Create (TableRAC, TobMere, numOrdre-1);
    TobFille.Dupliquer(TobActions, true, True);
    end
  else
  if ListeAlim[numOrdre] =ListeAlimChainages then
    begin
    stTable[numOrdre]:=TableRCH;
    SetControlCaption ('TITRE'+IntToStr(numOrdre),TitreChainages);
    TobFille:=TOB.Create (TableRCH, TobMere, numOrdre-1);
    TobFille.Dupliquer(TobChainages, true, True);
    end
  else
  if ListeAlim[numOrdre] =ListeAlimPropositions then
    begin
    stTable[numOrdre]:=TableRPE;
    SetControlCaption ('TITRE'+IntToStr(numOrdre),TitrePropositions);
    TobFille:=TOB.Create (TableRPE, TobMere, numOrdre-1);
    TobFille.Dupliquer(TobPropositions, true, True);
    end
  else
  if ListeAlim[numOrdre] =ListeAlimPieces  then
    begin
    stTable[numOrdre]:=TableGP;
    SetControlCaption ('TITRE'+IntToStr(numOrdre),TitrePieces);
    TobFille:=TOB.Create (TableGP, TobMere, numOrdre-1);
    TobFille.Dupliquer(TobPieces, true, True);
    end
  else
  if ListeAlim[numOrdre] =ListeAlimParc then
    begin
    SetControlCaption ('TITRE'+IntToStr(numOrdre),TitreParc);
    stTable[numOrdre]:=TableWPC;
    TobFille:=TOB.Create (TableWPC, TobMere, numOrdre-1);
    TobFille.Dupliquer(TobParc, true, True);
    end
  else
  if ListeAlim[numOrdre] =ListeAlimInterventions then
    begin
    SetControlCaption ('TITRE'+IntToStr(numOrdre),TitreInterventions);
    stTable[numOrdre]:=TableWIV;
    TobFille:=TOB.Create (TableWIV, TobMere, numOrdre-1);
    TobFille.Dupliquer(TobInterventions, true, True);
    end
  else
  if ListeAlim[numOrdre] =ListeAlimAffaires then
    begin
{$ifdef GIGI}
    SetControlCaption ('TITRE'+IntToStr(numOrdre),traduitGa(TitreAffaires));
{$else}
    SetControlCaption ('TITRE'+IntToStr(numOrdre),TitreAffaires);
{$endif}
    stTable[numOrdre]:=TableAFF;
    TobFille:=TOB.Create (TableAFF, TobMere, numOrdre-1);
    TobFille.Dupliquer(TobAffaires, true, True);
    end
  else  
  for i:=1 to nb_vignettes Do
    if ListeAlim[numOrdre] =ListeAlimcbs[i] then
      begin
      SetControlCaption ('TITRE'+IntToStr(numcbs[i]),Titrecbs[i]);
      stTable[numcbs[i]]:=Tablecbs[i];
      TobFille:=TOB.Create (Tablecbs[i], TobMere, numOrdre-1);
      TobFille.Dupliquer(Tobcbs[i], true, True);
      end
    ;
end;

procedure TOF_TIERS360.BPARAMVIGN_OnClick(Sender : tObject);
var ListeAvant,ListeApres : String;
    i : integer;
begin
  ListeAvant:=GetSynRegKey('RTTiersVign1','C',TRUE)+GetSynRegKey('RTTiersVign2','A',TRUE)+
  GetSynRegKey('RTTiersVign3','R',TRUE)+GetSynRegKey('RTTiersVign4','P',TRUE)+
  GetSynRegKey('RTTiersVign5','G',TRUE)+GetSynRegKey('RTTiersVign6','W',TRUE);

  AglLanceFiche('RT','RTPARAMTIERS360','','','');

  ListeApres:=GetSynRegKey('RTTiersVign1','C',TRUE)+GetSynRegKey('RTTiersVign2','A',TRUE)+
  GetSynRegKey('RTTiersVign3','R',TRUE)+GetSynRegKey('RTTiersVign4','P',TRUE)+
  GetSynRegKey('RTTiersVign5','G',TRUE)+GetSynRegKey('RTTiersVign6','W',TRUE);
  if ListeAvant <> ListeApres then
    begin
    { on remet visible les non visibles pour qu'ils reprennent leur place }
    for i:= 1 to nb_vignettes Do
      if THPanel(GetControl('P'+IntToStr(i))).Visible = False then
        THPanel(GetControl('P'+IntToStr(i))).Visible:=true;
    { traitement particulier pour 3 et 4 }
    THPanel(GetControl('P3')).Align:=alTop;
    THPanel(GetControl('P3')).Align:=alBottom;
    THPanel(GetControl('P4')).Align:=alTop;
    THPanel(GetControl('P4')).Align:=alBottom;

    RecupLesPosVignettes;
    end;
end;

procedure TOF_TIERS360.BTAILLE_OnClick(Sender : tObject);
begin
 TFVierge(ecran).width:=widthOri ;
 TFVierge(ecran).height:=heightOri;
 THPanel(GetControl('PBASE1')).width:=widthBase1Ori;
end;

procedure TOF_TIERS360.BTIERS_OnClick(Sender : tObject);
var stArg : string;
begin
  stArg:='ACTION=MODIFICATION;';
  if (RTDroitModiftiers(stTiers)=False) then stArg:= 'ACTION=CONSULTATION;';
  if stAuxi <> '' then
     AglLanceFiche('GC','GCTIERS','', stAuxi,stArg+'MONOFICHE;T_NATUREAUXI=' + GetControlText('T_NATUREAUXI')+';FICHE360');
end;

procedure TOF_TIERS360.AffecteClick;
begin
  if Assigned(GetControl('BPARAM1')) then
    TToolBarButton97(GetControl('BPARAM1')).OnClick := BParam1OnClick;
  if Assigned(GetControl('BPARAM2')) then
    TToolBarButton97(GetControl('BPARAM2')).OnClick := BParam2OnClick;
  if Assigned(GetControl('BPARAM3')) then
    TToolBarButton97(GetControl('BPARAM3')).OnClick := BParam3OnClick;
  if Assigned(GetControl('BPARAM4')) then
    TToolBarButton97(GetControl('BPARAM4')).OnClick := BParam4OnClick;
  if Assigned(GetControl('BPARAM5')) then
    TToolBarButton97(GetControl('BPARAM5')).OnClick := BParam5OnClick;
  if Assigned(GetControl('BPARAM6')) then
    TToolBarButton97(GetControl('BPARAM6')).OnClick := BParam6OnClick;
  if Assigned(GetControl('GRID1')) then
    THGRID(GetControl('GRID1')).OnDblClick := Grid1OnDblClick;
  if Assigned(GetControl('GRID2')) then
    THGRID(GetControl('GRID2')).OnDblClick := Grid2OnDblClick;
  if Assigned(GetControl('GRID3')) then
    THGRID(GetControl('GRID3')).OnDblClick := Grid3OnDblClick;
  if Assigned(GetControl('GRID4')) then
    THGRID(GetControl('GRID4')).OnDblClick := Grid4OnDblClick;
  if Assigned(GetControl('GRID5')) then
    THGRID(GetControl('GRID5')).OnDblClick := Grid5OnDblClick;
  if Assigned(GetControl('GRID6')) then
    THGRID(GetControl('GRID6')).OnDblClick := Grid6OnDblClick;
  if Assigned(GetControl('BFERME1')) then
    TToolBarButton97(GetControl('BFERME1')).OnClick := BFerme1OnClick;
  if Assigned(GetControl('BFERME2')) then
    TToolBarButton97(GetControl('BFERME2')).OnClick := BFerme2OnClick;
  if Assigned(GetControl('BFERME3')) then
    TToolBarButton97(GetControl('BFERME3')).OnClick := BFerme3OnClick;
  if Assigned(GetControl('BFERME4')) then
    TToolBarButton97(GetControl('BFERME4')).OnClick := BFerme4OnClick;
  if Assigned(GetControl('BFERME5')) then
    TToolBarButton97(GetControl('BFERME5')).OnClick := BFerme5OnClick;
  if Assigned(GetControl('BFERME6')) then
    TToolBarButton97(GetControl('BFERME6')).OnClick := BFerme6OnClick;
  if Assigned(GetControl('BPARAMVIGN')) then
    TToolBarButton97(GetControl('BPARAMVIGN')).OnClick := BPARAMVIGN_OnClick;
  if Assigned(GetControl('BTAILLE')) then
    TToolBarButton97(GetControl('BTAILLE')).OnClick := BTAILLE_OnClick;
  if Assigned(GetControl('BTIERS')) then
    TToolBarButton97(GetControl('BTIERS')).OnClick := BTIERS_OnClick;
  if Assigned(GetControl('BLISTE1')) then
    TToolBarButton97(GetControl('BLISTE1')).OnClick := BListe1OnClick;
  if Assigned(GetControl('BLISTE2')) then
    TToolBarButton97(GetControl('BLISTE2')).OnClick := BListe2OnClick;
  if Assigned(GetControl('BLISTE3')) then
    TToolBarButton97(GetControl('BLISTE3')).OnClick := BListe3OnClick;
  if Assigned(GetControl('BLISTE4')) then
    TToolBarButton97(GetControl('BLISTE4')).OnClick := BListe4OnClick;
  if Assigned(GetControl('BLISTE5')) then
    TToolBarButton97(GetControl('BLISTE5')).OnClick := BListe5OnClick;
  if Assigned(GetControl('BLISTE6')) then
    TToolBarButton97(GetControl('BLISTE6')).OnClick := BListe6OnClick;
  if Assigned(GetControl('BVALIDE1')) then
    TToolBarButton97(GetControl('BVALIDE1')).OnClick := BValide1OnClick;
  if Assigned(GetControl('BVALIDE2')) then
    TToolBarButton97(GetControl('BVALIDE2')).OnClick := BValide2OnClick;
  if Assigned(GetControl('BVALIDE3')) then
    TToolBarButton97(GetControl('BVALIDE3')).OnClick := BValide3OnClick;
  if Assigned(GetControl('BVALIDE4')) then
    TToolBarButton97(GetControl('BVALIDE4')).OnClick := BValide4OnClick;
  if Assigned(GetControl('BVALIDE5')) then
    TToolBarButton97(GetControl('BVALIDE5')).OnClick := BValide5OnClick;
  if Assigned(GetControl('BVALIDE6')) then
    TToolBarButton97(GetControl('BVALIDE6')).OnClick := BValide6OnClick;
end;

procedure TOF_TIERS360.RecupTailleEcran;
var val : integer;
begin
  widthOri:=TFVierge(ecran).width;
  heightOri:=TFVierge(ecran).height;
  widthBase1Ori:=THPanel(GetControl('PBASE1')).width;

  val:=GetSynRegKey('RTTiersWidth',widthOri,TRUE);
  if val <> 0 then
    TFVierge(ecran).width:=val ;

  val:=GetSynRegKey('RTTiersHeight',heightOri,TRUE);
  if val <> 0 then
    TFVierge(ecran).height:=val;
  val:=GetSynRegKey('RTTiersWidthBase1',widthBase1Ori,TRUE);
  if val <> 0 then
    THPanel(GetControl('PBASE1')).Width:=val;
end;

Function TOF_TIERS360.CBSRecupSql(CodeCBS, CodeTiers, CodeAuxiliaire,DateDebut,DateFin : String) : string;
var stMess : string;
begin
  result:='';
  result:=OnEvenementTiers360('GC_RecupSQL360',ecran,CodeCBS, CodeTiers, CodeAuxiliaire,'',DateDebut+';'+DateFin);
  if result='' then
    begin
    stMess:='Pas de requête SQL paramétrée en script (CBS) pour le code vignette '+CodeCBS;
    PgiInfo(stMess,'Requête de la vignette');
    end;
end;

Function TOF_TIERS360.CBSRecupTable(CodeCBS : String) : string;
var stMess : string;
begin
  result:='';
  result:=OnEvenementTiers360('GC_RecupTable360',ecran,CodeCBS, '','','','');
  if result='' then
    begin
    stMess:='Pas de table paramétrée en script (CBS) pour le code vignette '+CodeCBS;
    PgiInfo(stMess,'Table de la vignette');
    end;
end;

Function TOF_TIERS360.CBSClickListe(CodeCBS, CodeTiers, CodeAuxiliaire,stArg : String)  : string;
begin
  result:='';
  result:=OnEvenementTiers360('GC_ClickListe360',ecran,CodeCBS, CodeTiers, CodeAuxiliaire,stArg,'');
end;

Function TOF_TIERS360.CBSRecupListe(CodeCBS : String)  : string;
var stMess : string;
begin
  result:='';
  result:=OnEvenementTiers360('GC_RecupListe360',ecran,CodeCBS, '','','','');
  if result='' then
    begin
    stMess:='Pas de liste associée à la grille paramétrée en script (CBS) pour le code vignette '+CodeCBS;
    PgiInfo(stMess,'Liste de la vignette');
    end;
end;

Function TOF_TIERS360.CBSClickGrid(CodeCBS,valeur,stArg : String)  : string;
begin
  result:='';
  result:=OnEvenementTiers360('GC_ClickGrid360',ecran,CodeCBS, '','',stArg,Valeur);
end;

Function TOF_TIERS360.CBSTestChamp(CodeCBS : String)  : string;
var stMess : string;
begin
  result:='';
  result:=OnEvenementTiers360('GC_TestChamp360',ecran,CodeCBS, '','','','');
  if result='' then
    begin
    stMess:='Pas de champ à tester paramétrée en script (CBS) pour le code vignette '+CodeCBS;
    PgiInfo(stMess,'Champ de la vignette');
    end;
end;

Function TOF_TIERS360.CBSRecupCode(number : integer) : string;
begin
  result:='';
  result:=OnEvenementTiers360('GC_RecupCode360',ecran,'', '','','',IntToStr(number));
end;

Function TOF_TIERS360.CBSCacherDate(CodeCBS : String) : boolean;
begin
  result:=StrToBool(OnEvenementTiers360('GC_CacherDate360',ecran,CodeCBS, '','','',''));
end;

procedure TOF_TIERS360.RefreshTiers ;
var sql : string;
    Q : TQuery;
begin
  if TobTiers.GetString ('T_TIERS') = '' then
    begin
    sql:= 'SELECT * FROM TIERS WHERE T_AUXILIAIRE = "'+stAuxi+'"';
    Q := OpenSql(sql,False);
    TobTiers.SelectDB('',Q,True);
    Ferme(Q);
    end;
  if Assigned (TobTiers) then
    TobTiers.PutEcran(TFVierge(ecran));
  SetControlCaption('TLIB_APE',RechDom('YYCODENAF', GetControlText('T_APE'), False));
  if not TobTiers.GetBoolean('T_PARTICULIER') then
    SetControlCaption('TT_JURIDIQUE',TraduireMemoire('Abréviat.'));
  EtatRisque:='V' ;
  EtatRisque:=GetEtatRisqueClient(TOBTiers) ;
  SetControlText('RISQUE',EtatRisque);
end;

procedure TOF_TIERS360.SauveTailleEcran;
begin
  SaveSynRegKey('RTTiersHeight',TFVierge(ecran).Height,TRUE);
  SaveSynRegKey('RTTiersWidth',TFVierge(ecran).Width,TRUE);
  SaveSynRegKey('RTTiersWidthBase1',THPanel(GetControl('PBASE1')).Width,TRUE);
end;

procedure AGLRTInitDatesVign(parms:array of variant; nb: integer ) ;
begin
  SaveSynRegKey('RTTiersDateVR',DateToStr(DebutDeMois(Date))+';'+DateToStr(FinDeMois(Date)),TRUE);
  SaveSynRegKey('RTTiersDateVP',DateToStr(DebutDeMois(Date))+';'+DateToStr(FinDeMois(Date)),TRUE);
  SaveSynRegKey('RTTiersDateVH',DateToStr(DebutDeMois(Date))+';'+DateToStr(FinDeMois(Date)),TRUE);
  SaveSynRegKey('RTTiersDateVF',DateToStr(DebutDeMois(Date))+';'+DateToStr(FinDeMois(Date)),TRUE);
  SaveSynRegKey('RTTiersDateVG',DateToStr(DebutDeMois(Date))+';'+DateToStr(FinDeMois(Date)),TRUE);
  SaveSynRegKey('RTTiersDateVI',DateToStr(DebutDeMois(Date))+';'+DateToStr(FinDeMois(Date)),TRUE);
end;

Initialization
registerclasses([TOF_TIERS360]);
RegisterAglProc( 'RTInitDatesVign', true , 0, AGLRTInitDatesVign);
end.

