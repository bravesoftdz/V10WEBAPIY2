{***********UNITE*************************************************
Auteur  ...... : Philippe Dumet
Créé le ...... : 30/09/2002
Modifié le ... : 30/09/2002
Description .. :
Suite ........ :
Suite ........ :
Suite ........ :
Mots clefs ... : PAIE;BUDGET
*****************************************************************}
{ Module de gestion des budgets et de la masse salariale

}
unit UTofPG_PrepBudg;

interface
uses  StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls,HTB97,Hqry,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Fe_Main,
{$ELSE}
       MaineAgl,
{$ENDIF}
      Grids,HCtrls,HEnt1,EntPaie,HMsgBox,UTOF,UTOB,Vierge,
      AGLInit,Ed_Tools;
Type
     TOF_PGPrepBudg = Class (TOF)
       private
       TSal,TBudg, : TOB; // Tob utilisees
       TT                          : TQRProgressForm ;
       procedure LanceMaj (Sender: TObject);
       public
       procedure OnArgument(Arguments : String ) ; override ;
     END ;

implementation


procedure TOF_PGPrepBudg.LanceMaj (Sender: TObject);
var st,LeWhere     : String;
    BtnCherche     : TToolbarButton97;
    T1             : TOB; // des tob filles
    DB,DR          : TDateTime;
    QQ             : TQuery;
    RBTC,RBTB,RBTS : TRadioButton;
    rep            : Integer;
begin
DB := StrToDate (GetControlText ('DATEBUDGET'));
if DB > 0 then
   begin
   st := 'SELECT COUNT (*) FROM BUDGETPAIE WHERE PBG_TYPEBUDG="'+
       GetControlText ('TYPEBUDG')+'" AND PBG_DATEBUDG ="+UsDateTime(GetControlText ('DATEBUDGET'))+'"';
   if ExistSql (St) then
      begin
      rep := PGIAsk ('Attention, il existe déjà un budget sur cette période. Voulez-vous continuer ?',Ecran.Caption);
      if rep <> mrYes then exit;
      end;
   end;
RBTC := TRadioButton(GetControl ('RDBTCREAT'));
RBTB := TRadioButton(GetControl ('RDBTCOPIEBUDG'));
RBTS := TRadioButton(GetControl ('RDBTCOPIESAL'));
if (RBTC = NIL) OR (RBTB = NIL) OR (RBTS = NIL) then exit;
if RBTC.Checked then DR := DB;
if (RBTB.Checked) AND NOT IsValidDate(GetControlText ('DATERECOPIE'))  then
   begin
   PgiBox ('La date du budget à recopier est invalide', Ecran.Caption);
   exit;
   end;
if (RBTS.Checked) AND NOT IsValidDate(GetControlText ('DATERECOPIE'))  then
   begin
   PgiBox ('La date du budget concernant les salariés budgétés est invalide', Ecran.Caption);
   exit;
   end;
if RBTB.Checked then DR := StrToDate (GetControlText ('DATERECOPIE'));
if RBTS.Checked then DR := StrToDate (GetControlText ('DATESALARIES'));
LeWhere := ' WHERE (PSA_DATEENTREE <="'+UsDateTime (DR)+'") AND (((PSA_DATESORTIE >="'+UsDateTime (DR)+
    '")) OR (PSA_DATESORTIE IS NULL) OR (PSA_DATESORTIE <= "'+UsDateTime(iDate1900)+'"))';

st := 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_DATEENTREE,PSA_DATESORTIE,PSA_SUSPENSIONPAIE,PSA_LIBELLEEMPLOI'+
   'PSA_COEFFICIENT,PSA_QUALIFICATION,PSA_CODEEMPLOI,PSA_DATEANCIENNETE,PSA_HORAIREMOIS,PSA_SALAIRETHEO,PSA_TAUXHORAIRE,'+
   'PSA_SALAIREMOIS1,PSA_SALAIREMOIS2,PSA_SALAIREMOIS3,PSA_SALAIREMOIS4,PSA_SALAIREMOIS5,PSA_SALAIRANN1,PSA_SALAIRANN2,PSA_SALAIRANN3,'+
   'PSA_SALAIRANN4,PSA_SALAIRANN5,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_INDICE,PSA_NIVEAU,PSA_PRISEFFECTIF,'+
   'PSA_UNITEPRISEFF,PSA_LIBREPCMB1,PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4,PSA_DADSPROF FROM '+
   'SALARIES '+LeWhere+ ' ORDER BY PSA_SALARIE';
Tsal := TOB.create ('Mes Salaries', NIL, -1);
Q:=OpenSql(st, TRUE);
Tsal.LoadDetailDB('SALARIES','','',Q,FALSE,False) ;
Ferme (Q);
j := TSal.Detail.count-1;
TT :=  DebutProgressForm(NIL,'Calcul des informations pour le budget', 'Veuillez patienter SVP ...',j+25,TRUE,TRUE);

TT.Value:=TT.Value + 1 ; TT.SubText:='Chargement histo Ok' ;
try
BeginTrans;
TT.Value:=TT.value+1 ; TT.SubText:='Sauvegarde en cours ...'; ;
TTMP.SetAllModifie (TRUE);
TTMP.InsertOrUpdateDB (TRUE);
TT.Value:=TT.value+1 ; TT.SubText:='Traitement terminé ...'; ;
CommitTrans;
Except
Rollback;
PGIBox ('Une erreur est survenue lors de l''export',Ecran.Caption);
END;
// RAZTOB ();
end;


procedure TOF_PGPrepBudg.OnArgument(Arguments: String);
var BtnValid : TToolbarButton97;
begin
inherited ;
BtnValid:=TToolbarButton97 (GetControl ('BValider'));
if BtnValid <> NIL then BtnValidMul.OnClick := LanceMaj;
SetControlText ('TYPEBUDG','BUD');  // valeur par defaut à un budget
end;



Initialization
registerclasses([TOF_PGPrepBudg]);
end.
