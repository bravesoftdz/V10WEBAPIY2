unit AGlFctSuppr;     // Programme de suppression générale
interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, hmsgbox, Mask, Hctrls, StdCtrls, Menus, DB, DBTables, Hqry, Grids,
  DBGrids, ExtCtrls, ComCtrls, Buttons, HEnt1, Ent1, General, HStatus, RapSuppr,
  HCompte, HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HPanel,UiUtil, HRichOLE,
  M3FP,Suprsect,Section,Journal;


Type
  TEvtSup = Class
    Code    : string;
    Codax   : string;
    GCCode  : string;
    procedure DegageGene;
    Function  EstMouvemente(St : String; code : string):Boolean ;
    procedure Degagesection ;
    procedure DegageJournal ;
    procedure DegageTiers ;
    Function  EstDansSection(St : String) : Boolean ;
    Function  EstUnPayeur ( St : String ) : Boolean ;
    Function  EstDansPiece ( St : String ) : Boolean ;
    Function  EstDansActivite ( St : String ) : Boolean ;
    Function  EstDansRessource ( St : String ) : Boolean ;
    Function  EstDansAffaire ( St : String ) : Boolean ;
    Function  EstDansActions ( St : String ) : Boolean ;
    Function  EstDansPersp ( St : String ) : Boolean ;
    Function  EstDansCata ( St : String ) : Boolean ;
    Function  EstDansPaie ( St : String ) : Boolean ;
    Function  EstDansMvtPaie ( St : String ) : Boolean ;

end;
procedure SupprimeListeEnreg(L : THDBGrid; Q :TQuery; St:string);
Function DetruitGene(Code,Libelle : String):Byte;
Function DetruitJournal(Stc,Lib : String):Byte ;
Function EstDansAnalytiqsection(Stc,Sta : String) : Boolean ;
Function DetruitTiers ( St,RefGC,Lib : String):Byte ;

implementation

var
    TNotDel   : TList ;
    TDelGene  : TList ;
    Effacer   : Boolean ;
    NotEffacer : Boolean ;
    HM : array [0..17] of string = ('','Ce compte est mouvementé','Ce compte est un modèle',
    'Ce compte possède des écritures analytiques',
    'Ce compte est un compte d''attente pour un axe analytique',
    'Ce compte est un compte de banque',
    'Ce compte est un compte de correspondance',
    'Ce compte est un compte d''écart de change pour les devises',
    'Ce compte possède des mouvements d''immobilisation',
    'Ce compte est un compte de contrepartie d''un journal de banque',
    'Ce compte est associé à un mode de paiement',
    'Ce compte est un compte d''attente, de clôture d''ouverture ou d''écart Euro',
    'Ce compte est associé à un paramétrage de TVA',
    '13;Suppression des comptes généraux;Vous n''avez rien sélectionné;E;O;O;O;',
    '14;Suppression des comptes généraux;Désirez-vous un compte-rendu des comptes détruits?;Q;YNC;N;C;',
    '15;Suppression des comptes généraux;Désirez-vous un compte-rendu des comptes non détruits?;Q;YNC;N;C;',
    'Compte supprimé','Compte en cours d''utilisation !');

    HMAUX : array [0..35] of string =('','Ce compte auxiliaire est mouvementé',
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
    'Aucun','élément sélectionné','éléments sélectionnés',
    '22;','23;','24;','25;',
    'Ce compte est mouvementé en pièces de Gestion Commerciale',
    'Ce compte est mouvementé en règlements de Gestion Commerciale',
    'Ce compte est référencé dans les activités',
    'Ce compte est référencé dans les ressources',
    'Ce compte est référencé en affaire',
    'Ce compte est référencé dans les actions',
    'Ce compte est référencé dans les perspectives',
    'Ce compte est référencé comme salarié',
    'Ce compte est référencé dans le catalogue',
    'Ce compte est mouvementé dans la paie comme salarié');
                                    
     HMJAL : array [0..22] of string =
     ('Journal supprimé ',
     'Ce journal à des écritures analytiques',
     'Ce journal à des écritures comptables',
     'Ce journal à des écritures d''immobilisations',
     'Ce journal est un un modèle',
     'Ce journal est un journal d''ouverture ou de fermeture pour la société',
     'Ce journal à des écritures comptables pour la gestion commerciale',
     'Compte en cours d''utilisation !', 'Aucun',
     'élément sélectionné',
     'éléments sélectionnés',
     '1;Suppression des journaux;Vous ne pouvez pas supprimer ce journal : il est mouvementé par des écritures analytiques.;W;O;O;O;',
     '2;Suppression des journaux;Vous ne pouvez pas supprimer ce journal : il est mouvementé par des écritures comptables.;W;O;O;O;',
     '3;Suppression des journaux;Vous ne pouvez pas supprimer ce journal : il est mouvementé par des écritures d''immobilisations.;W;O;O;O;',
     '4;Suppression des journaux;Vous ne pouvez pas supprimer ce journal : il est un modèle.;W;O;O;O;',
     '5;Suppression des journaux;Vous ne pouvez pas supprimer ce journal : il est journal d''ouverture ou de fermeture pour la société.;W;O;O;O;',
     '6;Suppression des journaux;Vous ne pouvez pas supprimer ce journal : il est référencé comme journal de relevé.;W;O;O;O;',
     '7;Suppression des journaux;Confirmez-vous la suppression des enregistrements sélectionnés?;Q;YNC;N;C;',
     '8;Suppression des journaux;Désirez-vous un compte-rendu des comptes détruits?;Q;YNC;N;C;',
     '9;Suppression des journaux;Désirez-vous un compte-rendu des comptes non détruits?;Q;YNC;N;C;',
     '10;Suppression des journaux;Vous ne pouvez pas supprimer ce journal : il est en cours d''utilisation.;W;O;O;O;',
     '11;Suppression des journaux;Vous ne pouvez pas supprimer ce journal : il est référencé comme journal de tiers payeur.;W;O;O;O;',
     '12;Suppression des journaux;Vous ne pouvez pas supprimer ce journal : il est référencé comme journal de reprise de balance.;W;O;O;O;');


procedure TEvtSup.DegageGene;
BEGIN
if ExecuteSQL('DELETE FROM GENERAUX WHERE G_GENERAL="'+Code+'"')<>1 then V_PGI.IoError:=oeUnknown ;
END;

procedure TEvtSup.DegageTiers ;
Var DelTout : Boolean ;
BEGIN
DelTout:=True ;
if ExecuteSQL('DELETE FROM TIERS WHERE T_AUXILIAIRE="'+Code+'"')<>1 then
   BEGIN
   V_PGI.IoError:=oeUnknown ; Deltout:=False ;
   END ;
if DelTout then
   BEGIN
   ExecuteSQL('Delete from CONTACT Where C_AUXILIAIRE="'+Code+'"') ;
   ExecuteSQL('Delete from RIB Where R_AUXILIAIRE="'+Code+'"') ;
   {Gescom}
   ExecuteSQL('DELETE FROM TIERSPIECE WHERE GTP_TIERS="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM TIERSCOMPL WHERE YTC_TIERS="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM ARTICLETIERS WHERE GAT_REFTIERS="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM TARIF WHERE GF_TIERS="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM PROSPECTS WHERE RPR_AUXILIAIRE="'+Code+'"') ;
   ExecuteSQL('DELETE FROM ADRESSES WHERE ADR_TYPEADRESSE="TIE" AND ADR_REFCODE="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM LIENSOLE WHERE LO_TABLEBLOB="T" AND LO_IDENTIFIANT="'+GCCode+'"') ;
      // ExecuteSQL('UPDATE ANNUAIRE SET ANN_TIERS="", ANN_AUXILIAIRE="" WHERE ANN_TIERS="'+GCCode+'"') ;
   ExecuteSQL('UPDATE ANNUAIRE SET ANN_TIERS="" WHERE ANN_TIERS="'+GCCode+'"') ;
   {Paie}
   END ;
END ;

procedure TEvtSup.Degagesection ;
BEGIN
if ExecuteSQL('DELETE FROM SECTION WHERE S_SECTION="'+Code+'" AND S_AXE="'+Codax+'"')<>1 then V_PGI.IoError:=oeUnknown ;
END ;

Function TEvtSup.EstMouvemente(St : String; code : string):Boolean ;
BEGIN
if code = 'G' then
Result:=ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="'+St+'" '+
                  ' AND ((EXISTS(SELECT E_GENERAL FROM ECRITURE WHERE E_GENERAL="'+St+'"))'+
                  ' Or (EXISTS(SELECT Y_GENERAL FROM ANALYTIQ WHERE Y_GENERAL="'+St+'")))') ;
if code = 'T' then
Result:=ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="'+St+'" '+
                  'AND EXISTS(SELECT E_AUXILIAIRE FROM ECRITURE WHERE E_AUXILIAIRE="'+St+'" )') ;

END ;

procedure TEvtSup.DegageJournal ;
BEGIN
if ExecuteSQL('DELETE FROM JOURNAL WHERE J_JOURNAL="'+Code+'"')<>1 then V_PGI.IoError:=oeUnknown ;
END ;

Function TEvtSup.EstDansSection(St : String) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select S_MAITREOEUVRE,S_CHANTIER From SECTION Where S_MAITREOEUVRE="'+St+'" OR S_CHANTIER="'+St+'"',True) ;
Result:=(Not Q.EOF) ;
Ferme(Q) ;
END ;

Function  TEvtSup.EstUnPayeur ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('SELECT T_PAYEUR FROM TIERS WHERE T_PAYEUR="'+St+'"',True) ;
Result:=(Not Q.EOF) ;
Ferme(Q) ;
END ;

Function TEvtSup.EstDansPiece ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT GP_TIERS FROM PIECE WHERE GP_TIERS="'+St+'" OR GP_TIERSLIVRE="'+St+'" OR GP_TIERSFACTURE="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;

Function TEvtSup.EstDansActivite ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT ACT_TIERS FROM ACTIVITE WHERE ACT_TIERS="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;

Function TEvtSup.EstDansRessource ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT ARS_AUXILIAIRE FROM RESSOURCE WHERE ARS_AUXILIAIRE="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;

Function TEvtSup.EstDansActions ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT RAC_AUXILIAIRE FROM ACTIONS WHERE RAC_AUXILIAIRE="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;

Function TEvtSup.EstDansPersp ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT RPE_AUXILIAIRE FROM PERSPECTIVES WHERE RPE_AUXILIAIRE="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;

Function TEvtSup.EstDansAffaire ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT AFF_TIERS FROM AFFAIRE WHERE AFF_TIERS="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;

Function TEvtSup.EstDansCata ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT GCA_TIERS FROM CATALOGU WHERE GCA_TIERS="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;

Function TEvtSup.EstDansPaie ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT PSA_AUXILIAIRE FROM SALARIES WHERE PSA_AUXILIAIRE="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;

Function TEvtSup.EstDansMvtPaie ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT PPU_SALARIE FROM PAIEENCOURS LEFT JOIN SALARIES ON PPU_SALARIE=PSA_SALARIE WHERE PSA_AUXILIAIRE="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;


Function DetruitGene(Code,Libelle : String):Byte;
var
EvtSup : TEvtSup;
j      : integer;
X,Y    : DelInfo ;
BEGIN
  EvtSup := TEvtSup.Create;
  EvtSup.Code := Code;
  Result:= 0;

  if EvtSup.EstMouvemente(Code, 'G') then Result:=1
  else if EstAnalytiqPure(Code) then Result:=3
  else if EstCpteAxe(Code)      then Result:=4
  else if EstCpteCorresp(Code)  then Result:=6
  else if EstCpteDevise(Code)   then Result:=7
  else if EstCpteJournal(Code)  then Result:=9
  else if EstCpteModepaie(Code) then Result:=10
  else if EstCpteSociete(Code)  then Result:=11
  else if EstCpteTva(Code)      then Result:=12;

   if Result = 0 then
   begin
       if Transactions(EvtSup.DegageGene,5)<>oeOK then
       BEGIN
       MessageAlerte('Suppression impossible') ; Result:=17 ;
       END else
       BEGIN
       ExecuteSQL('DELETE FROM VENTIL WHERE V_COMPTE="'+Code+'" And V_NATURE Like"GE%"') ;
       ExecuteSql('DELETE From BANQUECP Where BQ_GENERAL="'+Code+'"') ;
       END ;
   end;
   if  Result <=0 then
   BEGIN
            X:=DelInfo.Create ;
            X.LeCod:=Code ;
            X.LeLib:=Libelle ;
            X.LeMess:='Compte supprimé' ;
            TDelGene.Add(X) ; Effacer:=True ;
   END else
   BEGIN
            Y:=DelInfo.Create ;
            Y.LeCod:= Code ;
            Y.LeLib:= Libelle ;
            Y.LeMess:= HM[Result] ;
            TNotDel.Add(Y) ;  NotEffacer:=True ;
   END;

EvtSup.Free;
END ;

Function DetruitTiers ( St,RefGC,Lib : String):Byte ;
var
EvtSup : TEvtSup;
X,Y    : DelInfo ;
BEGIN
  EvtSup := TEvtSup.Create;
  EvtSup.Code := St;
  EvtSup.GCCode := RefGc;

  Result:= 0;
  if EvtSup.EstMouvemente(St, 'T')  then Result:=1
  else if EstEcrGuide(St) then Result:=2
  else if EstDansSociete(St) then Result:=10
  else if EvtSup.EstDansSection(St) then Result:=11
  else if Presence('UTILISAT','US_AUXILIAIRE',St) then Result:=13
  else if EstCpteCorresp(St) then Result:=9
  else if EvtSup.EstUnPayeur(St) then Result:=14
  else if EvtSup.EstDansPiece(RefGC) then Result:=26
  else if EvtSup.EstDansActivite(RefGC) then Result:=28
  else if EvtSup.EstDansRessource(St) then Result:=29
  else if EvtSup.EstDansAffaire(RefGC) then Result:=30
  else if EvtSup.EstDansActions(St) then Result:=31
  else if EvtSup.EstDansPersp(St) then Result:=32
  else if EvtSup.EstDansCata(RefGC) then Result:=34
  {Paie}
  else if EvtSup.EstDansPaie(St) then Result:=33
  else if EvtSup.EstDansMvtPaie(St) then Result:=35;

  if Result = 0 then
  begin
     if Transactions(EvtSup.DegageTiers,5)<>oeOK then
     BEGIN
     MessageAlerte('Suppression impossible') ; Result:=17 ;
     END;
  end;
  if Result =0 then
  BEGIN
          X:=DelInfo.Create ; X.LeCod:=St ; X.LeLib:=Lib ; X.LeMess:=HMAUX[17] ;
          TDelGene.Add(X) ; Effacer:=True ;
  END else
  BEGIN
          Y:=DelInfo.Create ; Y.LeCod:=St ; Y.LeLib:=Lib ;
          Y.LeMess:=HMAUX[Result] ;
          TNotDel.Add(Y) ;  NotEffacer:=True ;
  END ;

EvtSup.Free;
END ;


Function DetruitSection(Stc,Sta : String):Byte ;
var
EvtSup : TEvtSup;
j      : integer;
BEGIN
  EvtSup := TEvtSup.Create;
  EvtSup.Code := Stc;
  EvtSup.Codax := Sta;

  Result:= 0;
if EstDansAnalytiqsection(Stc,Sta)then BEGIN Result:=1 ; Exit ; END ;
if EstDansAxe(Stc,Sta)then BEGIN Result:=2 ; Exit ; END ;
if EstCorresp(Stc,Sta)then BEGIN Result:=3 ; Exit ; END ;
if EstDansVentil(Stc,Sta)then BEGIN Result:=4 ; Exit ; END ;
if Transactions(EvtSup.Degagesection,5)<>oeOK then
   BEGIN
   MessageAlerte('Suppression impossible') ; Result:=17 ; Exit ;
   END;
EvtSup.Free;
END ;

Function DetruitJournal(Stc,Lib : String):Byte ;
var
EvtSup : TEvtSup;
X,Y    : DelInfo ;
BEGIN
  EvtSup := TEvtSup.Create;
  EvtSup.Code := Stc;

  Result:= 0;
  if EstDansEcriture(Stc)then Result:=2
  else if EstDansAnalytiq(Stc)then Result:=1
  else if EstDansSociete(Stc) then Result:=5
  else if EstDansSouche(Stc)  then Result:=6
  else if ((Stc=VH^.JalATP) or (StC=VH^.JalVTP)) then Result:=11
  else if Stc=VH^.JalRepBalAN then Result:=12;
  if Result = 0 then
  begin
    if Transactions(EvtSup.DegageJournal,5)<>oeOK then
     BEGIN
     MessageAlerte('Suppression impossible') ; Result:=17 ;
     END;
  end;
  if Result = 0 then
  BEGIN
          X:=DelInfo.Create ; X.LeCod:=Stc ; X.LeLib:=Lib ; X.LeMess:=HMJAL[0] ;
          TDelGene.Add(X) ; Effacer:=True ;
  END else
  BEGIN
          Y:=DelInfo.Create ; Y.LeCod:=Stc ; Y.LeLib:=Lib ;
          Y.LeMess:=HMJAL[Result] ;
          TNotDel.Add(Y) ;  NotEffacer:=True ;
  END;

EvtSup.Free;
END ;


procedure SupprimeListeEnreg(L : THDBGrid; Q :TQuery; St:string);
var i,j : integer;
Texte : string;
Q1    : TQuery;
begin
  TDelGene:=TList.Create ; TNotDel:=TList.Create ;
  Effacer:=False ; NotEffacer:=False ;

  if (L.NbSelected=0) and (not L.AllSelected) then
  begin
    MessageAlerte('Aucun élément sélectionné');
    exit;
  end;
  Texte:='Vous allez supprimer définitivement les informations.#13#10Confirmez vous l''opération ?';
  if HShowMessage('0;Suppression;'+Texte+';Q;YN;N;N;','','')<>mrYes then exit ;
  if L.AllSelected then
  BEGIN
    Q.First;
    while Not Q.EOF do
    BEGIN
      MoveCur(False);
      if St ='GENERAUX' then
      begin
           DetruitGene(Q.FindField('G_GENERAL').AsString,Q.FindField('G_LIBELLE').AsString);
      end;
      if St ='SECTION' then
      begin
           if (DetruitSection(Q.FindField('S_SECTION').AsString,Q.FindField('S_AXE').AsString) >0) then break;
      end;
      if St ='JOURNAL' then
      begin
           if (DetruitJournal(Q.FindField('J_JOURNAL').AsString,Q.FindField('J_LIBELLE').AsString) >0) then break;
      end;
      if St = 'TIERS' then
      begin
           Q1 := OpenSql ('SELECT T_TIERS FROM TIERS Where T_AUXILIAIRE="'+ Q.FindField('T_AUXILIAIRE').AsString+'"',TRUE);
           if not Q1.EOF then
           begin
           if (DetruitTiers(Q.FindField('T_AUXILIAIRE').AsString,Q1.FindField('T_TIERS').AsString,Q.FindField('T_LIBELLE').AsString) >0) then break;
           Ferme (Q1);
           end;
      end;
      Q.Next;
    END;
    L.AllSelected:=False;
  END
  ELSE
  BEGIN
    InitMove(L.NbSelected,'');
    for i:=0 to L.NbSelected-1 do
    BEGIN
      MoveCur(False);
      L.GotoLeBookmark(i);
      if St ='GENERAUX' then
      begin
           DetruitGene(Q.FindField('G_GENERAL').AsString,Q.FindField('G_LIBELLE').AsString);
      end;
      if St ='SECTION' then
      begin
           if (DetruitSection(Q.FindField('S_SECTION').AsString,Q.FindField('S_AXE').AsString) >0) then break;
      end;
      if St ='JOURNAL' then
      begin
           if (DetruitJournal(Q.FindField('J_JOURNAL').AsString,Q.FindField('J_LIBELLE').AsString) >0) then break;
      end;
      if St = 'TIERS' then
      begin
           Q1 := OpenSql ('SELECT T_TIERS FROM TIERS Where T_AUXILIAIRE="'+ Q.FindField('T_AUXILIAIRE').AsString+'"',TRUE);
           if not Q1.EOF then
           begin
           if (DetruitTiers(Q.FindField('T_AUXILIAIRE').AsString,Q1.FindField('T_TIERS').AsString,Q.FindField('T_LIBELLE').AsString) >0) then break;
           Ferme (Q1);
           end;
      end;
    END;
    L.ClearSelected;
  END;
  FiniMove;
  if St = 'GENERAUX' then
  begin
       if Effacer    then if HShowMessage(HM[14],'','')=mrYes then RapportDeSuppression(TDelGene,1) ;
       if NotEffacer then if HShowMessage(HM[15],'','')=mrYes then RapportDeSuppression(TNotDel,1) ;
  end;
  if St = 'TIERS' then
  begin
       if Effacer    then if HShowMessage(HMAUX[15],'','')=mrYes then RapportDeSuppression(TDelGene,1) ;
       if NotEffacer then if HShowMessage(HMAUX[16],'','')=mrYes then RapportDeSuppression(TNotDel,1) ;
  end;
  if St = 'JOURNAL' then
  begin
       if Effacer    then if HShowMessage(HMJAL[18],'','')=mrYes then RapportDeSuppression(TDelGene,1) ;
       if NotEffacer then if HShowMessage(HMJAL[19],'','')=mrYes then RapportDeSuppression(TNotDel,1) ;
  end;
  TDelGene.Clear ; TDelGene.Free ; TNotDel.Clear ; TNotDel.Free ;
end;

procedure AGLSupprimeListCpte(parms: array of variant; nb: integer );
var
  F : TForm;
  Liste : THDBGrid;
  Query : TQuery;
begin
  F:=TForm(Longint(Parms[0])) ;
  if F=Nil then exit ;
  Liste:=THDBGrid(F.FindComponent('FListe') );
  if Liste=Nil then exit;
  Query:=TQuery(F.FindComponent('Q')) ;
  if (Query=Nil) then exit;
  SupprimeListeEnreg(Liste,Query, parms[1]);
end;

Function EstDansAnalytiqsection(Stc,Sta : String) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select Y_SECTION From ANALYTIQ Where Y_SECTION="'+Stc+'" AND Y_AXE="'+Sta+'"',True) ;
Result:=(Not Q.EOF) ; Ferme(Q) ;
END ;


Initialization
  RegisterAglProc( 'SupprimeListeCpte', TRUE , 1, AGLSupprimeListCpte);
finalization

end.
