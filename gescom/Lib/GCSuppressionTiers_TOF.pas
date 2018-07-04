unit GCSuppressionTiers_Tof;

interface

uses  Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HMsgBox,UTOF,HQry,HEnt1,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
      mul,Fe_Main,dbTables,
{$ENDIF}
      Hstatus,HDB,M3FP,UTOB,UtilPGI, TiersUtil,UtilSuprAuxi,
      RapSuppr;

function GCLanceFiche_SuppressionTiers (Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
    Tof_GCSupprTiers = Class (TOF)
    private
        TDelAux  : TList ;
        TNotDel   : TList ;
        Effacer   : Boolean ;
        NotEffacer : Boolean ;
        AuxiCode,GCCode  : String ;
        Function  GCDetruit ( St,RefGC : String): Byte ;
        Procedure GCDegage ;
    Public
        procedure OnArgument (Arguments : String ); override ;
        procedure OnClose ; override ;
        procedure GCSuppressionTiers;
        procedure GCSuppression (Auxiliaire,Tiers,Lib : String) ;
     END;

const
	// libellés des messages
	HM: array[0..36] of string 	= (
        '',
        'Ce compte auxiliaire est mouvementé',
        'Ce compte auxiliaire est référençé dans un guide d''écriture comptable',
        'Ce compte auxiliaire sert de modèle pour la création',
        'Ce compte auxiliaire est référencé dans un catalogue',
        'Ce compte auxiliaire est référencé dans une ligne d''écriture',
        'Ce compte auxiliaire est référencé dans une ligne de pièce interne',
        'Ce compte auxiliaire est référencé dans une pièce',
        'Ce compte auxiliaire est référencé dans une pièce interne',
        'Ce compte auxiliaire est un compte de correspondance',
        'Ce compte auxiliaire est un client ou un fournisseur d''attente',
        'Ce compte auxiliaire est référencé dans une section analytique',
        'Ce compte auxiliaire est référencé pour un autre compte auxiliaire',
        'Ce compte auxiliaire est un utilisateur',
        '14;Suppression des comptes auxiliaires;Vous n''avez rien sélectionné;E;O;O;O;',
        '15;Suppression des comptes auxiliaires;Désirez-vous un compte rendu des comptes détruits?;Q;YNC;N;C;',
        '16;Suppression des comptes auxiliaires;Désirez-vous un compte rendu des comptes non détruits?;Q;YNC;N;C;',
        'Compte supprimé',
        'Compte en cours d''utilisation !',
        'Aucun',
        'élément sélectionné',
        'éléments sélectionnés',
        '22;','23;','24;','25;',
        'Ce compte est mouvementé en pièces de Gestion Commerciale',
        'Ce compte est mouvementé en règlements de Gestion Commerciale',
        'Ce compte est référencé dans les activités',
        'Ce compte est référencé dans les ressources',
        'Ce compte est référencé en affaire',
        'Ce compte est référencé dans les actions',
        'Ce compte est référencé dans les propositions',
        'Ce compte est référencé comme salarié',
        'Ce compte est référencé dans le catalogue',
        'Ce compte est mouvementé dans la paie comme salarié',
        'Ce compte est référencé dans les projets'
          );
implementation

function GCLanceFiche_SuppressionTiers (Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
    Result := '';
    if Nat = '' then exit;
    if Cod = '' then exit;
    Result := AGLLanceFiche (Nat, Cod, Range, Lequel, Argument);
end;

procedure tof_GCSupprTiers.OnArgument (Arguments : String );
begin
  inherited ;
  SetControlText ('T_NATUREAUXI', Arguments);
  if Arguments = 'CLI' then
  begin
    Ecran.Caption := 'Suppression des clients';
    SetControlProperty ('TT_TIERS', 'Caption', 'Client');
    SetControlProperty ('T_TIERS', 'DataType', 'GCTIERSCLI')
  end else
  begin
    Ecran.Caption := 'Suppression des Fournisseurs';
    SetControlProperty ('TT_TIERS', 'Caption', 'Fournisseur');
    SetControlProperty ('T_TIERS', 'DataType', 'GCTIERSFOURN');
  end;
  UpdateCaption (Ecran);
end;

procedure tof_GCSupprTiers.OnClose;
begin
_DeblocageMonoPoste(False,'',TRUE);
end;

procedure tof_GCSupprTiers.GCSuppressionTiers;
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
if PGIAsk('Confirmez vous le traitement ?','')<>mrYes then exit ;

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
      GCSuppression (CodeAuxiliaire,CodeTiers,Libelle);
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
      Q.TQ.Seek(F.FListe.Row-1) ;
{$ENDIF}
      CodeTiers:=F.Q.FindField('T_TIERS').asstring ;
      CodeAuxiliaire:=F.Q.FindField('T_AUXILIAIRE').asstring ;
      Libelle:=TFmul(Ecran).Q.FindField('T_LIBELLE').asstring ;
      GCSuppression (CodeAuxiliaire,CodeTiers,Libelle);
      end;
   L.ClearSelected;
   end;
if Effacer    then if HShowMessage(HM[15],'','')=mrYes then RapportDeSuppression(TDelAux,1) ;
if NotEffacer then if HShowMessage(HM[16],'','')=mrYes then RapportDeSuppression(TNotDel,1) ;

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

procedure tof_GCSupprTiers.GCSuppression (Auxiliaire,Tiers,Lib : String);
Var j : Byte ;
    X,Y : DelInfo ;
begin
j:=GCDetruit(Auxiliaire,Tiers) ;
if j<=0 then
   begin
   X:=DelInfo.Create ; X.LeCod:=Tiers ; X.LeLib:=Lib ; X.LeMess:=HM[17] ;
   TDelAux.Add(X) ; Effacer:=True ;
   end else
   begin
   Y:=DelInfo.Create ; Y.LeCod:=Tiers ; Y.LeLib:=Lib ;
   Y.LeMess:=HM[j] ;
   TNotDel.Add(Y) ;  NotEffacer:=True ;
   end
end;

Function tof_GCSupprTiers.GCDetruit ( St,RefGC : String):Byte ;
BEGIN
Result:=0 ;
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
if SupAEstDansCata(RefGC) then BEGIN Result:=34 ; Exit ; END ;
if SupAEstDansProjets(RefGC) then BEGIN Result:=36 ; Exit ; END ;
{Paie}
if SupAEstDansPaie(St) then BEGIN Result:=33 ; Exit ; END ;
if SupAEstDansMvtPaie(St) then BEGIN Result:=35 ; Exit ; END ;
AuxiCode:=St ; GCCode:=RefGC ;
if Transactions(GCDegage,5)<>oeOK then BEGIN MessageAlerte(HM[18]) ; Result:=17 ; Exit ; END ;
END ;

procedure tof_GCSupprTiers.GCDegage ;
Var DelTout : Boolean ;
BEGIN
DelTout:=True ;
if ExecuteSQL('DELETE FROM TIERS WHERE T_AUXILIAIRE="'+AuxiCode+'"')<>1 then
   BEGIN
   V_PGI.IoError:=oeUnknown ; Deltout:=False ;
   END ;
if DelTout then
   BEGIN
   ExecuteSQL('Delete from CONTACT Where C_AUXILIAIRE="'+AuxiCode+'"') ;
   ExecuteSQL('Delete from RIB Where R_AUXILIAIRE="'+AuxiCode+'"') ;
   {Gescom}
   ExecuteSQL('DELETE FROM TIERSPIECE WHERE GTP_TIERS="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM TIERSCOMPL WHERE YTC_TIERS="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM ARTICLETIERS WHERE GAT_REFTIERS="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM TARIF WHERE GF_TIERS="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM ADRESSES WHERE ADR_TYPEADRESSE="TIE" AND ADR_REFCODE="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM LIENSOLE WHERE LO_TABLEBLOB="T" AND LO_IDENTIFIANT="'+GCCode+'"') ;
   ExecuteSQL('UPDATE ANNUAIRE SET ANN_TIERS="", ANN_AUXILIAIRE="" WHERE ANN_TIERS="'+GCCode+'"') ;
   END ;
END ;


procedure AGLGCSuppressionTiers(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is tof_GCSupprTiers) then tof_GCSupprTiers(TOTOF).GCSuppressionTiers else exit;
end;

Initialization
registerclasses([tof_GCSupprTiers]);
RegisterAglProc('GCSuppressionTiers',True,0,AGLGCSuppressionTiers);

end.
