{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Unit de gestion du multi critère de la reprise de saisie
Suite ........ : des cumuls et des bases de cotisation
Mots clefs ... : PAIE
*****************************************************************}
{
PT1   : 07/01/2002 VG V571 Les salariés ayant une date de sortie, même
                           postèrieure à la date de reprise, n'étaient pas dans
                           la liste des salariés - FICHE de BUG N°410
PT2   : 01/10/2002 VG V585 Si le code salarié est numérique, on remplit à gauche
                           par des 0 - Fiche qualité N° 10253
}
unit UTofPG_MulReprise;

interface
uses  StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Mul,Fe_Main,DbGrids,
{$ELSE}
       MaineAgl,eMul,
{$ENDIF}
      Grids,HCtrls,HEnt1,vierge,EntPaie,HMsgBox,Hqry,UTOF,UTOB,UTOM,
      AGLInit,PgOutils,PGoutils2,P5Def;

Type
     TOF_PGMULREPRISE = Class (TOF)
       private
       WW : THEdit;
       AuDebut   : String;
       DDebut,DFin : TDateTime;
       DateDebut, DateFin : THEdit;
       procedure ActiveWhere (Quoi : Boolean);
       procedure GrilleDblClick (Sender: TObject);
       public
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnLoad ; override;
       procedure ExitSal(Sender: TObject);
       end;

          TOF_PGMULCONSULTREPRISE = Class (TOF)
       procedure OnArgument(Arguments : String ) ; override ;
       private
       procedure ExitEdit(Sender : Tobject);

     END ;

implementation

procedure TOF_PGMULREPRISE.ActiveWhere(Quoi : Boolean);
begin
if Quoi = FALSE then
   begin // la requete ne peut qu'échouer donc liste vide donc pas de saisie possible
   if WW <> NIL then WW.Text:='(PSA_SUSPENSIONPAIE <> "X") AND (PSA_SUSPENSIONPAIE = "X")';
   exit;
   end;
// RendBornesExerSocial (DDebut,DFin,vcbxAnnee.Value);
DateDebut := THEdit (GetControl ('DATEDEBUT'));
if DateDebut = NIL then exit;
DateFin := THEdit (GetControl ('DATEFIN'));
if DateFin = NIL then exit;

DDebut  := StrToDate (DateDebut.Text);
DFin := StrToDate (DateFin.Text);
if (WW <>NIL) then
//PT1
{
 WW.Text:='((PSA_SUSPENSIONPAIE <> "X") AND (PSA_DATEENTREE <="'+UsDateTime (DFin)+'") AND (((PSA_DATESORTIE >="'+UsDateTime (DDebut)+
    '") AND (PSA_DATESORTIE <="'+UsDateTime (DFin)+'")) OR (PSA_DATESORTIE IS NULL) OR (PSA_DATESORTIE <= "'+UsDateTime(iDate1900)+'")))';
}

 WW.Text:= '((PSA_SUSPENSIONPAIE <> "X") AND'+
           ' (PSA_DATEENTREE <="'+UsDateTime (DFin)+'") AND'+
           ' ((PSA_DATESORTIE >="'+UsDateTime (DDebut)+'") OR'+
           ' (PSA_DATESORTIE IS NULL) OR'+
           ' (PSA_DATESORTIE <= "'+UsDateTime(iDate1900)+'")))';
//FIN PT1
end;


procedure TOF_PGMULREPRISE.OnArgument(Arguments: String);
var Num     : Integer;
{$IFNDEF EAGLCLIENT}
    Grille  : THDBGrid;
{$ELSE}
    Grille  : THGrid;
{$ENDIF}
    LeDebut : TCheckBox;
    CodeSal: THEdit;
begin
inherited ;
{$IFNDEF EAGLCLIENT}
Grille:=THDBGrid (GetControl ('Fliste'));
{$ELSE}
Grille:=THGrid (GetControl ('Fliste'));
{$ENDIF}
if Grille <> NIL then Grille.OnDblClick := GrilleDblClick;
WW:=THEdit (GetControl ('XX_WHERE'));
//vcbxAnnee := THValComboBox (GetControl ('CBXANNEE'));
LeDebut := TCheckBox (GetControl ('CHBXDEBUTEXER'));
if LeDebut <> NIL then
 begin
 if LeDebut.Checked = TRUE then AuDebut := 'X' else AuDebut :='-';
 end;
{
if RendExerSocialEnCours (MoisE, AnneeE, ComboExer, DebExer, FinExer) = TRUE then
 begin
 if vcbxAnnee <> NIL then vcbxAnnee.value:=ComboExer;
 end;
if vcbxAnnee <> NIL then vcbxAnnee.OnChange := AnneeChange;
}
For Num := 1 to 4 do
 begin
 VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
 end;
VisibiliteStat(GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT'));

//PT2
CodeSal:=ThEdit(getcontrol('PSA_SALARIE'));
If CodeSal<>nil then CodeSal.OnExit:=ExitSal;
//FIN PT2
end;


procedure TOF_PGMULREPRISE.GrilleDblClick(Sender: TObject);
var CodeSal,St: String;
    LeDebut : TCheckBox;
    DateDebut, DateFin : THEdit;
begin
LeDebut := TCheckBox (GetControl ('CBXDEBUTEXER'));
if LeDebut <> NIL then
 begin
 if LeDebut.Checked = TRUE then AuDebut := 'X' else AuDebut :='-';
 end;
DateDebut := THEdit (GetControl ('DATEDEBUT'));
if DateDebut = NIL then exit;
DateFin := THEdit (GetControl ('DATEFIN'));
if DateFin = NIL then exit;
{$IFDEF EAGLCLIENT}
   TheMulQ:=TOB (TFMul(Ecran).Q.TQ);
   if TheMulQ.RecordCount=0 then exit; //grille vide
   TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row-1) ;
   CodeSal:=TheMulQ.detail[TFMul(Ecran).FListe.Row-1].GetValue('PSA_SALARIE');
 {$ELSE}
   TheMulQ := THQuery(Ecran.FindComponent('Q'));
   if TheMulQ.RecordCount=0 then exit; //grille vide
   CodeSal:=TheMulQ.findfield('PSA_SALARIE').AsString;
{$ENDIF}
St := CodeSal+';'+AuDebut+';'+DateDebut.Text+';'+DateFin.Text;
AGLLanceFiche ('PAY','REPCUMULS',  '', '', St);
TheMulQ := NIL;
end;

procedure TOF_PGMULREPRISE.OnLoad;
var DateDebut,DateFin : THEdit;
    DD, DF : TDateTime;
    FinMois, DebutMois : TDateTime;
begin
inherited;
DateDebut := THEdit (GetControl ('DATEDEBUT'));
if DateDebut = NIL then begin ActiveWhere (FALSE); exit; end;
DateFin := THEdit (GetControl ('DATEFIN'));
if DateFin = NIL then begin ActiveWhere (FALSE); exit; end;
DD := StrToDate (DateDebut.text);
DF := StrToDate (DateFin.text);
if DD > DF then
   begin
   PGIBox ('La date de début ne pas être supérieure à la date de fin', Ecran.Caption);
   ActiveWhere (FALSE);
   end;
DebutMois := DEBUTDEMOIS (DD);
if DebutMois <> DD then
   begin
   PGIBox ('Attention, la date de début ne correspond pas au début du mois', Ecran.Caption);
   ActiveWhere (FALSE);
   end;

FinMois := FINDEMOIS (DF);
if FinMois <> DF then
   begin
   PGIBox ('Attention, la date de fin ne correspond pas à la fin du mois', Ecran.Caption);
   ActiveWhere (FALSE);
   end;
ActiveWhere (TRUE);
end;


//PT2
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 01/10/2002
Modifié le ... :
Description .. : Si le code salarié est numérique, on remplit à gauche par
Suite ........ : des 0
Mots clefs ... : PAIE
*****************************************************************}
procedure TOF_PGMULREPRISE.ExitSal(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;
//FIN PT2
{TOF_PGMULCONSULTREPRISE}
procedure TOF_PGMULCONSULTREPRISE.OnArgument(Arguments : String ) ;
var Edit : THEdit;
begin
        Edit := THEdit(GetControl('PHC_SALARIE'));
        If Edit <> Nil then Edit.OnExit := ExitEdit;
        Edit := THEdit(GetControl('PHC_SALARIE_'));
        If Edit <> Nil then Edit.OnExit := ExitEdit;
end;

procedure TOF_PGMULCONSULTREPRISE.ExitEdit(Sender : TObject);
var edit : thedit;
begin
        edit:=THEdit(Sender);
        if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
        if (VH_Paie.PgTypeNumSal='NUM') and
       (length(Edit.text)<11) and
       (isnumeric(edit.text)) then
      edit.text:=AffectDefautCode(edit,10);
end;

Initialization
registerclasses([TOF_PGMULREPRISE,TOF_PGMULCONSULTREPRISE]);
end.
