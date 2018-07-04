{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... : 24/08/2005
Description .. : CCMX-CEGID : Gestion Transport
Mots clefs ... :
*****************************************************************}
unit UtofRTSuppressionTiers;

interface

uses  Controls,Classes,forms,
      HCtrls,HMsgBox,UTOF,HQry,HEnt1, ParamSoc,
      utilselection,          //MulCreerPagesCL
{$ifdef GIGI}
      EntGC,
{$ENDIF}
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
      HDB,mul,Fe_Main,
{$ENDIF}
{$IFDEF GRC}
      UtilRT,
{$ENDIF}
{$ifdef AFFAIRE}
      utofAFANCETRE_GIGA_GRC,
{$ENDIF}
      Utob,Hstatus,M3FP,UtilPGI,UtilSuprTiers,RapSuppr;

Type
{$ifdef AFFAIRE}
                // PL le 18/10/07 : FQ CRM 10737 : on gère tout ce qu'il y a de commun entre
                // GIGA et GRC dans une TOF spécialisée
                //mcd 24/11/2005 pour faire affectation depuis ressource si paramétré
    Tof_RTSupprTiers = Class (TOF_ANCETRE_GIGA_GRC)
{$else}
    Tof_RTSupprTiers = Class (TOF)
{$endif}
    private
        TDelAux  : TList ;
        TNotDel   : TList ;
        Effacer   : Boolean ;
        NotEffacer : Boolean ;
        AuxiCode,GCCode  : String ;
        Function  RTDetruit ( St,RefGC : String): Byte ;
        Procedure RTDegage ;
{$ifdef GIGI}
        procedure ChangeNature (Sender: TObject) ;
{$endif}
    Public
        procedure OnLoad ; override ;
        procedure OnArgument (Arguments : String ); override ;
        procedure OnClose ; override ;
        procedure RTSuppressionTiers;
        procedure RTSuppression (Auxiliaire,Tiers,Lib : String) ;
        procedure RTSuppLesInfosComplRPE(AuxiCode : string);
     END;


Function RTLanceFiche_RTSupprTiers(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

implementation

uses
  wCommuns, MessagesErreur;

procedure tof_RTSupprTiers.OnArgument (Arguments : String );
var stArg : string;
begin
fMulDeTraitement  := true;
{$ifndef GIGI}   // il est fait plus loin en contexte GIGI
  inherited ;
{$endif}
  //if Not _BlocageMonoPoste(False,'',TRUE) then Exit ;
	 fTableName := 'TIERS';

  stArg := Arguments;
  if stArg = '' then stArg := 'PRO';
{$IFDEF GRC}
//  SetControlVisible('PRESSOURCE', False);   // PL le 18/10/07 : FQ CRM 10737
  if stArg = 'PRO' then
  begin
    SetControlText ('NATUREAUXI','PRO');
    SetControlEnabled ('NATUREAUXI',False);
//TJA 02/07/2007
    SetControlVisible ('PCOMPLEMENT', True);
    SetControlProperty('PCOMPLEMENT', 'CAPTION', 'Tables libres');
    SetControlVisible ('PCOMPLEMENTFOUR', False);
    SetControlVisible('ZONESLIBRES', True);
    {$IFDEF AFFAIRE}                      //FQ 10632
//    SetControlVisible('PRESSOURCE', True);      // PL le 18/10/07 : FQ CRM 10737
    {$ENDIF}
    MulCreerPagesCL(TFMul(Ecran), 'NOMFIC=GCTIERS');
  end else
{$ENDIF} // GRC
  begin
    SetControlVisible ('RTSUPPRTIERSACT', False);
    SetControlVisible ('RTSUPPRTIERSPRO', False);
    SetControlVisible ('NATUREAUXI', False);
    SetControlVisible ('TT_NATUREAUXI', False);
    {$IFDEF AFFAIRE}                    // FQ 10632
//    SetControlVisible('PRESSOURCE', True);    // PL le 18/10/07 : FQ CRM 10737
    {$ENDIF}
    if stArg = 'CLI' then
    begin
      SetControlVisible ('PCOMPLEMENTFOUR', False);
      MulCreerPagesCL(TFMul(Ecran), 'NOMFIC=GCTIERS');
    end
    else
    begin
      SetControlVisible ('PCOMPLEMENT', False);
      SetControlVisible('ZONESLIBRES', False);
//      SetControlVisible('PRESSOURCE', False);   // PL le 18/10/07 : FQ CRM 10737
    end;

  end;

  SetControlText ('T_NATUREAUXI', stArg);
  if stArg = 'CLI' then
  begin
    Ecran.Caption := TraduireMemoire('Suppression des clients');
    SetControlProperty ('TT_TIERS', 'Caption', TraduireMemoire('Client'));
  end else if stArg = 'FOU' then
  begin
    Ecran.Caption := Traduirememoire('Suppression des Fournisseurs');
    SetControlProperty ('TT_TIERS', 'Caption', TraduireMemoire('Fournisseur'));
  end else
  begin
    Ecran.Caption := Traduirememoire('Suppression des prospects');
    SetControlProperty ('TT_TIERS', 'Caption', TraduireMemoire('Prospect'));
  end;
{$ifdef GIGI}
  inherited ; //mcd 12/07/06 il faut le faire après affectation nature AUxi pour Ok affectation table libre fou
if GetControlText('NATUREAUXI') ='PRO' then
  begin   //si appel depuis GRC, il faut plus de nature
    SetControlText ('NATUREAUXI',''); //on efface tout car NCP en plus
    SetControlText('T_NatureAuxi','PRO');  //on froce à PRO pour OK rech cleint (13971)
    SetControlProperty ('T_NATUREAUXI', 'Complete', true);
    SetControlProperty ('NATUREAUXI','Complete', true);
    SetControlProperty ('T_NATUREAUXI', 'Datatype', 'TTNATTIERS');
    SetControlProperty ('NATUREAUXI', 'Datatype', 'TTNATTIERS');
    SetControlProperty ('NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
    SetControlProperty ('T_NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
    SetControlText ('TT_TIERS','Tiers');
    SetControlProperty ('RTSUPPRTIERSACT','caption',TraduireMemoire('Suppression tiers avec actions'));
    SetControlProperty  ('RTSUPPRTIERSPRO','caption',TraduireMemoire('Suppression tiers avec propositions'));
    SetControlVisible ('NATUREAUXI', True);
    SetControlEnabled ('NATUREAUXI', True);
    SetControlVisible ('TT_NATUREAUXI', True);
    Ecran.Caption := Traduirememoire('Suppression des tiers');
    SetControlProperty ('T_NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
    THValCOmboBox (GetControl ('NATUREAUXI') ) .OnExit := ChangeNature;
  End;
 If (Not GetParamSocSecur ('SO_AFRTPROPOS',False)) then
   begin
   SetcontrolVisible ('RTSUPPRTIERSPRO',False);
   end;
{$endif GIGI}
  UpdateCaption (Ecran);
end;

procedure tof_RTSupprTiers.OnLoad;
{$IFDEF GRC}
//var xx_where,Confid : string;
var Confid : string;
{$ENDIF}
begin
  inherited ;
  if ( GetControlText ('T_NATUREAUXI') = 'PRO' ) or ( GetControlText ('T_NATUREAUXI') = 'CLI' )
      or ( GetControlText ('T_NATUREAUXI') = 'FOU' ) then // DBR - Fiche 10898
  begin
    {$IFDEF GRC}
    if ctxGRC in V_PGI.PgiContexte then
    begin
      if GetControlText ('T_NATUREAUXI') = 'FOU'  then Confid:='MODF' else Confid:='MOD';
      if (GetControlText('XX_WHERE') = '') then
          SetControlText('XX_WHERE',RTXXWhereConfident(Confid,true))
      else
          begin
//FQ 10632          xx_where := GetControlText('XX_WHERE');
//FQ 10632          xx_where := xx_where + RTXXWhereConfident(Confid,true);
//FQ 10632          SetControlText('XX_WHERE',xx_where) ;
            SetControlText('XX_WHERE', RTXXWhereConfident(Confid, True)); //FQ 10632
          end;
    end;
    {$ENDIF} // GRC
  end;
end;

procedure tof_RTSupprTiers.OnClose;
begin
{$ifdef GIGI}
//mcd 21/11/07 on ne fait plus de blocage monoposte  Bloqueur('AffToutSeul',False);  //mcd 18/09/07 14491
{$else}
_DeblocageMonoPoste(False,'',TRUE);
{$ENDIF}
end;

procedure tof_RTSupprTiers.RTSuppressionTiers;
var  F : TFMul;
{$IFDEF EAGLCLIENT}
       L : THGrid;
{$ELSE}
       L : THDBGrid;
{$ENDIF}
     Q : THQuery;
     i : integer;
     CodeAuxiliaire,CodeTiers,Libelle : string;
     X : DelInfo;
begin
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
   begin
   PGIInfo('Aucun élément sélectionné','');
   exit;
   end;
if PGIAsk('Confirmez-vous le traitement ?','')<>mrYes then
   begin
   if F.FListe.AllSelected then
      begin
      F.bSelectAll.Down := False;
      F.FListe.AllSelected:=False;
      end;
   exit ;
   end;
if F.FListe.AllSelected then
   begin
   if (PGIAsk('Attention : Vous avez sélectionné la totalité des tiers de la liste - Confirmez-vous votre choix?','')<> mrYes) then
      begin
      F.bSelectAll.Down := False;
      F.FListe.AllSelected:=False;
      Exit;
      end;
   end;

{$IFDEF EAGLCLIENT}
if F.bSelectAll.Down then
   if not F.Fetchlestous then
     begin
     F.bSelectAllClick(Nil);
     F.bSelectAll.Down := False;
     exit;
     end else
     F.Fliste.AllSelected := true;
{$ENDIF}

TDelAux:=TList.Create ; TNotDel:=TList.Create ;
Effacer:=False ; NotEffacer:=False ;

L:= F.FListe;
Q:= F.Q;

if L.AllSelected then
   begin
   InitMove(Q.RecordCount,'');
   Q.First;
   while Not Q.EOF do
      begin
      MoveCur(False);
      CodeAuxiliaire:=TFmul(Ecran).Q.FindField('T_AUXILIAIRE').asstring ;
      CodeTiers:=TFmul(Ecran).Q.FindField('T_TIERS').asstring ;
      Libelle:=TFmul(Ecran).Q.FindField('T_LIBELLE').asstring ;
      RTSuppression (CodeAuxiliaire,CodeTiers,Libelle);
      Q.Next;
      end;
   L.AllSelected:=False;
   end else
   begin
   InitMove(L.NbSelected,'');
   for i:=0 to L.NbSelected-1 do
      begin
      MoveCur(False);
      L.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
      Q.TQ.Seek(L.Row-1) ;
{$ENDIF}
      CodeTiers:=F.Q.FindField('T_TIERS').asstring ;
      CodeAuxiliaire:=F.Q.FindField('T_AUXILIAIRE').asstring ;
      Libelle:=TFmul(Ecran).Q.FindField('T_LIBELLE').asstring ;
      RTSuppression (CodeAuxiliaire,CodeTiers,Libelle);
      end;
   L.ClearSelected;
   end;
if Effacer    then if HShowMessage(MSGERR_SupprTiers[15],'','')=mrYes then RapportDeSuppression(TDelAux,1) ;
if NotEffacer then if HShowMessage(MSGERR_SupprTiers[16],'','')=mrYes then RapportDeSuppression(TNotDel,1) ;

if F.bSelectAll.Down then
    F.bSelectAll.Down := False;
FiniMove;

for i:=0 to TDelAux.Count-1 do
    BEGIN
    X:=TDelAux.Items[i] ;
    X.free;
    END ;
for i:=0 to TNotDel.Count-1 do
    BEGIN
    X:=TNotDel.Items[i] ;
    X.free;
    END ;
TDelAux.Clear ; TDelAux.Free ; TNotDel.Clear ; TNotDel.Free ;
end ;

procedure tof_RTSupprTiers.RTSuppression (Auxiliaire,Tiers,Lib : String);
Var j : Byte ;
    X,Y : DelInfo ;
begin
j:=RTDetruit(Auxiliaire,Tiers) ;
if j<=0 then
   begin
   X:=DelInfo.Create ; X.LeCod:=Tiers ; X.LeLib:=Lib ; X.LeMess:=MSGERR_SupprTiers[17] ;
   TDelAux.Add(X) ; Effacer:=True ;
   end else
   begin
   Y:=DelInfo.Create ; Y.LeCod:=Tiers ; Y.LeLib:=Lib ;
   Y.LeMess:=MSGERR_SupprTiers[j] ;
   TNotDel.Add(Y) ;  NotEffacer:=True ;
   end
end;

Function tof_RTSupprTiers.RTDetruit ( St,RefGC : String):Byte ;
var tSupControl :T_SupTiersControle;
BEGIN
  Result:=0 ;
  tSupControl := T_SupTiersControle.Create;
  try
    with tSupControl do
    begin
      if EstBaseMultiSoc then Societes := GetBasesMS
      else Societes := '';
      Encours := false;
      if SupAEstMouvemente(St)  then BEGIN Result:=1 ;  Exit ; END ;
      if SupAEstEcrGuide(St)    then BEGIN Result:=2 ;  Exit ; END ;
      if SupAEstDansSociete(St) then BEGIN Result:=10 ; Exit ; END ;
      if SupAEstDansSection(St) then BEGIN Result:=11 ; Exit ; END ;
      if SupAEstDansUtilisat(St)then BEGIN Result:=13 ; Exit ; END ;
      if SupAEstCpteCorresp(St) then BEGIN Result:=9 ;  Exit ; END ;
      if SupAEstUnPayeur(St)    then BEGIN Result:=14 ; Exit ; END ;
      {Gescom}
      if SupAEstDansPiece(RefGC) then BEGIN Result:=26 ; Exit ; END ;
      if SupAEstDansActivite(RefGC) then BEGIN Result:=28 ; Exit ; END ;
      if SupAEstDansRessource(St) then BEGIN Result:=29 ; Exit ; END ;
      if SupAEstDansAffaire(RefGC) then BEGIN Result:=30 ; Exit ; END ;
      // BDU - 11/04/07 - FQ : 13920
      if SupAEstDansIntervenant(RefGC) then begin Result := 39; Exit; end;
      if GetControlText('RTSUPPRTIERSACT') <> 'X' then if SupAEstDansActions(St) then BEGIN Result:=31 ; Exit ; END ;
      if GetControlText('RTSUPPRTIERSPRO') <> 'X' then if SupAEstDansPersp(St) then BEGIN Result:=32 ; Exit ; END ;
      if SupAEstDansCata(RefGC) then BEGIN Result:=34 ; Exit ; END ;
      if SupAEstDansProjets(RefGC) then BEGIN Result:=36 ; Exit ; END ;
      if SupAEstDansGED(RefGC) then BEGIN Result:=37 ; Exit ; END ;
      {Paie}
      if SupAEstDansPaie(St) then BEGIN Result:=33 ; Exit ; END ;
      if SupAEstDansMvtPaie(St) then BEGIN Result:=35 ; Exit ; END ;
      if SupAEstDansParc(RefGC) then BEGIN Result:=38 ; Exit ; END ;
      AuxiCode:=St ; GCCode:=RefGC ;
      if Transactions(RTDegage,5)<>oeOK then BEGIN MessageAlerte(MSGERR_SupprTiers[18]) ; Result:=17 ; Exit ; END ;
    end;
  finally
    tSupControl.Free;
  end;
END ;

procedure tof_RTSupprTiers.RTDegage ;
Var DelTout : Boolean ;
BEGIN
DelTout:=True ;
if ExecuteSQL('DELETE FROM TIERS WHERE T_AUXILIAIRE="'+AuxiCode+'"')<>1 then
   BEGIN
   V_PGI.IoError:=oeUnknown ; Deltout:=False ;
   END ;
if DelTout then
   BEGIN
   if (GetParamsocSecur('SO_AFLIENDP',False) =true)  then
     begin  //mcd  20/03/2006 GIGA 12758
     ExecuteSql ('update CONTACT set C_AUXILIAIRE=C_GUIDPER ,C_TYPECONTACT="ANN",C_TIERS="" from contact Where C_TYPECONTACT="T" AND C_AUXILIAIRE="'+AuxiCode+'"' );
     end
   else ExecuteSQL('Delete from CONTACT Where C_TYPECONTACT="T" AND C_AUXILIAIRE="'+AuxiCode+'"') ;
   ExecuteSQL('Delete from RIB Where R_AUXILIAIRE="'+AuxiCode+'"') ;
   {Gescom}
   ExecuteSQL('DELETE FROM TIERSPIECE WHERE GTP_TIERS="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM TIERSCOMPL WHERE YTC_TIERS="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM ARTICLETIERS WHERE GAT_REFTIERS="'+GCCode+'"');
   wDeleteTable('YTARIFS', 'YTS_TIERS = "'+GCCode+'"');
{$IFDEF GRC}
  if ctxGRC in V_PGI.PgiContexte then
    ExecuteSQL('DELETE FROM PROSPECTS WHERE RPR_AUXILIAIRE="'+AuxiCode+'"') ;
{$ENDIF} // GRC
   ExecuteSQL('DELETE FROM ADRESSES WHERE ADR_TYPEADRESSE="TIE" AND ADR_REFCODE="'+GCCode+'"') ;
   // CCMX CEGID TRANSPORT : NR le 03/10/2005
   ExecuteSQL('DELETE FROM TIERSEXPE WHERE GXT_TYPEADRESSE="TIE" AND GXT_TIERS="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM GCPRODUITTRA WHERE GTC_TRANSPORTEUR="'+GCCode+'"') ;
   // FIN CCMX CEGID TRANSPORT
//   ExecuteSQL('DELETE FROM LIENSOLE WHERE LO_TABLEBLOB="T" AND LO_IDENTIFIANT="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM LIENSOLE WHERE LO_TABLEBLOB="T" AND LO_IDENTIFIANT="'+AuxiCode+'"') ;
      //mcd 17/05/05 champ auxi supprimé ExecuteSQL('UPDATE ANNUAIRE SET ANN_TIERS="", ANN_AUXILIAIRE="" WHERE ANN_TIERS="'+GCCode+'"') ;
   ExecuteSQL('UPDATE ANNUAIRE SET ANN_TIERS="" WHERE ANN_TIERS="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM TIERSFRAIS WHERE GTF_TIERS="'+GCCode+'"');
{$IFDEF GRC}
  if ctxGRC in V_PGI.PgiContexte then
  begin
   ExecuteSQL('DELETE FROM ACTIONS WHERE RAC_AUXILIAIRE="'+AuxiCode+'"') ;
   ExecuteSQL('DELETE FROM RTINFOS001 WHERE RD1_CLEDATA like "'+AuxiCode+';%"') ;
   ExecuteSQL('DELETE FROM ACTIONSCHAINEES WHERE RCH_TIERS="'+GCCode+'"') ;
   if GetParamSocSecur ('SO_RTGESTINFOS00V',false) then
     RTSuppLesInfosComplRPE(AuxiCode);
   ExecuteSQL('DELETE FROM PERSPECTIVES WHERE RPE_AUXILIAIRE="'+AuxiCode+'"') ;
   ExecuteSQL('DELETE FROM PERSPHISTO WHERE RPH_AUXILIAIRE="'+AuxiCode+'"') ;
   ExecuteSQL('DELETE FROM RTINFOS006 WHERE RD6_CLEDATA like "T;'+AuxiCode+';%"') ;
  end;
{$ENDIF} // GRC
   ExecuteSQL('DELETE FROM RTINFOS003 WHERE RD3_CLEDATA = "'+AuxiCode+'"') ;
{CRM_CRM10820_CD}
   ExecuteSQL('DELETE FROM ACTIONINTERVENANT WHERE RAI_AUXILIAIRE = "'+AuxiCode+'"') ;
   {Paie}
   END ;
END ;

procedure tof_RTSupprTiers.RTSuppLesInfosComplRPE(AuxiCode : string);
var tobPro : tob;
    stSQL : string;
    i : integer;
begin
  stSQL := 'SELECT RPE_PERSPECTIVE FROM PERSPECTIVES WHERE RPE_AUXILIAIRE="'+AuxiCode+'"';
  tobPro := TOB.Create('NumPropal', nil, -1);
  tobPro.LoadDetailDBFromSQL('PERSPECTIVES', stSQL);
  if tobPro.Detail.Count > 0 then
    begin
    for i := 0 to Pred(tobPro.Detail.Count) Do
      ExecuteSQL ('DELETE FROM RTINFOS00V WHERE RDV_CLEDATA = '+tobPro.detail[i].GetString('RPE_PERSPECTIVE'));
    end;
  tobPro.free;
end;

procedure AGLRTSuppressionTiers(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is tof_RTsupprTiers) then tof_RTsupprTiers(TOTOF).RTSuppressionTiers else exit;
end;

Function RTLanceFiche_RTSupprTiers(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
{$IFDEF GESCOM}
  Result := '';
  if Nat = '' then exit;
  if Cod = '' then exit;
{$ENDIF}
  AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

{$ifdef GIGI}
procedure tof_RTSupprTiers.ChangeNature (Sender: TObject) ;
begin
 SetControlText('T_NATUREAUXI',GetControlText ('NATUREAUXI'));
 if getControlText('NatureAuxi')='' then  SetControlText('T_NATUREAUXI','PRO'); //mcd pour OK si rech client  13971
end;
{$endif}


Initialization
registerclasses([tof_RTsupprTiers]);
RegisterAglProc('RTSuppressionTiers',True,0,AGLRTSuppressionTiers);

end.
