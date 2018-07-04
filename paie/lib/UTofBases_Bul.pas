{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 25/06/2001
Modifié le ... : 25/06/2001
Description .. : Affichage pendant la saisie du bulletin sur click droit
Suite ........ : des lignes des basesbcotisations pour un salarié pour une
Suite ........ : seule base de cotisation
Suite ........ : Ceci permet de voir les infos ayant servi aux calculs des
Suite ........ : réguls de paie
Mots clefs ... : PAIE;PGBULLETIN
*****************************************************************}
{
PT1  : 29/08/2005 PH V_60 FQ 12521 erreur req SQL
PT2 : 31/03/2006 PH V_65 FQ 12568 Affichage Code+libelle dans la tablette
PT3 : 31/03/2006 PH V_65 FQ 12225 Prise en compte paie validée du mois
}
unit UTofBases_Bul;

interface
uses  StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DBGrids,
{$ELSE}
      UtileAGL,
{$ENDIF}
      uPaieBases,
      Grids,HCtrls,HEnt1,HMsgBox,UTOF, UTOB,  Vierge, P5Util, P5Def, AGLInit;
Type
     TOF_PGBASE_BUL = Class (TOF)
       private
       CodeSal, Etab, DateD, DateF : String;
       DD, DF : TDateTime;
       Grille : THGrid;
       VCbxLRub : THValComboBox;
       procedure RechHisto(Sender: TObject);
       procedure PasDeSaisieRow(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
       procedure PasDeSaisieCell(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
//       procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
       public
       procedure OnLoad ; override ;
       procedure OnArgument(Arguments : String ) ; override ;       
     END ;

implementation
{
procedure TOF_PGBASE_BUL.GetCellCanvas(Acol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
if ARow = Grille.RowCount - 2 then
 begin
 Grille.RowHeights[ARow] := 2;
 exit;
 end;
if ARow <> Grille.RowCount - 1 then exit; // On ne fait rien car on n'est pas positionné sur la dernière ligne affichant les cumuls
Grille.canvas.Font.Height:=13;
if ACol=1 then Grille.canvas.Font.Style:=Grille.canvas.Font.Style+[fsItalic]+[fsBold]
 else Grille.canvas.Font.Style:=Grille.canvas.Font.Style+[fsItalic];
Grille.Canvas.Brush.Color:=Grille.FixedColor ;
Grille.Canvas.Font.Color:=Grille.Font.Color ;
end;
}
procedure TOF_PGBASE_BUL.OnArgument(Arguments: String);
var i : Integer;
    T1 : TOB;
    Btn : TToolbarButton97;
    st : String;
begin
inherited ;
st:=Trim (Arguments);
CodeSal:=ReadTokenSt(st);// Recup code Salarie
Etab:=ReadTokenSt(st);   // Recup Code Etablissement
DateD:=ReadTokenSt(st);  // Recup Date debut et date de fin de la session de paie
DateF:=ReadTokenSt(st);
RendDateExerSocial (StrToDate(DateD),StrToDate(DateF), DD,DF, TRUE); //PT1 PT3

Btn := TToolbarButton97(GetControl('RECHHISTO'));
if Btn <> NIL then Btn.OnClick := RechHisto;

Grille :=THGrid (GetControl('GRILLEBASE')) ;
if Grille <> NIL then
 begin
 Grille.OnRowEnter := PasDeSaisieRow;
 Grille.OnCellEnter := PasDeSaisieCell;
 Grille.CacheEdit ;
 Grille.Options:=Grille.Options-[GoEditing,GoTabs,GoAlwaysShowEditor] ;
 Grille.Options:=Grille.Options+[GoRowSelect] ;
 end
 else Grille.Enabled := FALSE;
VCbxLRub:=THValComboBox(GetControl ('VALCBXRUB'));     // Controle sur la combo contenant les codes des rubriques de bases
if VCbxLRub<>NIL then
 begin // il convient de récupèrer LaTob pour alimenter les combos Au maxi 5 à 6 Rubriques
 VCbxLRub.clear;
 for I:=0 to LaTOB.Detail.Count-1 do
   begin
   T1 := LaTOB.Detail[I] ;
   VCbxLRub.Items.Add(T1.GetValue('CODEBASE')+' '+T1.GetValue ('LIBELBASE')) ; // PT2
   VCbxLRub.Values.Add(T1.GetValue ('CODEBASE')) ;
   end ;
   VCbxLRub.ItemIndex:=0;
 end;
// Design canvas de cellule ==> Grille.GetCellCanvas:=GetCellCanvas;
end;

procedure TOF_PGBASE_BUL.OnLoad;
begin

end;

procedure TOF_PGBASE_BUL.PasDeSaisieCell(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
exit;
end;

procedure TOF_PGBASE_BUL.PasDeSaisieRow(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
exit;
end;

procedure TOF_PGBASE_BUL.RechHisto(Sender: TObject);
var Q : TQuery;
    St,Rub : String;
    i, NbDec : Integer;
    T2 : TOB;
    CC : Array [1..8] of Double;
begin
Rub:=VCbxLRub.Value;
for i:=1 to 8 do
 begin
 CC[i]:=0;
 end;
St:='SELECT PHB_DATEDEBUT,PHB_DATEFIN,PHB_BASECOT,PHB_PLAFOND,PHB_TRANCHE1,PHB_TRANCHE2,PHB_TRANCHE3,PHB_PLAFOND1,PHB_PLAFOND2,PHB_PLAFOND3'+
    ' FROM HISTOBULLETIN WHERE PHB_SALARIE="'+CodeSal+'" AND PHB_NATURERUB="BAS" AND PHB_RUBRIQUE="'+Rub+
    '" AND PHB_DATEDEBUT>="'+UsDateTime(DD)+'" AND PHB_DATEFIN<="'+UsDateTime(DF)+'" ORDER BY PHB_DATEFIN';
Q:=OpenSQL(St,TRUE) ;
T2:=TOB_Bases.FindFirst (['PCT_RUBRIQUE'],[Rub], TRUE);
NbDec:=2;
Grille.RowCount:=2;
if T2 <> NIL then NbDec:=T2.GetValue('PCT_DECBASECOT');
i:=1;
// Boucle sur les 12 mois maxi de la paie dans un exercice soit 12 lignes remplies dans la grille
While Not Q.EOF do
  begin
  Grille.Cells[0,i]:= DateToStr(Q.FindField('PHB_DATEDEBUT').AsFloat);  // DateDebut
  Grille.Cells[1,i]:= DateToStr(Q.FindField('PHB_DATEFIN').AsFloat);  // DateFin
  CC[1]:=CC[1]+Q.FindField('PHB_BASECOT').AsFloat;
  CC[2]:=CC[2]+Q.FindField('PHB_PLAFOND').AsFloat;
  CC[3]:=CC[3]+Q.FindField('PHB_TRANCHE1').AsFloat;
  CC[4]:=CC[4]+Q.FindField('PHB_TRANCHE2').AsFloat;
  CC[5]:=CC[5]+Q.FindField('PHB_TRANCHE3').AsFloat;
  CC[6]:=CC[6]+Q.FindField('PHB_PLAFOND1').AsFloat;
  CC[7]:=CC[7]+Q.FindField('PHB_PLAFOND2').AsFloat;
  CC[8]:=CC[8]+Q.FindField('PHB_PLAFOND3').AsFloat;
  Grille.Cells[2,i]:= DoubleToCell (Q.FindField('PHB_BASECOT').AsFloat, NbDec); // Base de Cotisation
  Grille.Cells[3,i]:= DoubleToCell (Q.FindField('PHB_PLAFOND').AsFloat, NbDec); // Plafond
  Grille.Cells[4,i]:= DoubleToCell (Q.FindField('PHB_TRANCHE1').AsFloat, NbDec); // Tranche 1
  Grille.Cells[5,i]:= DoubleToCell (Q.FindField('PHB_TRANCHE2').AsFloat, NbDec); // Tranche 2
  Grille.Cells[6,i]:= DoubleToCell (Q.FindField('PHB_TRANCHE3').AsFloat, NbDec); // Tranche 3
  Grille.Cells[7,i]:= DoubleToCell (Q.FindField('PHB_PLAFOND1').AsFloat, NbDec); // Plafond Tranche 1
  Grille.Cells[8,i]:= DoubleToCell (Q.FindField('PHB_PLAFOND2').AsFloat, NbDec); // Plafond Tranche 2
  Grille.Cells[9,i]:= DoubleToCell (Q.FindField('PHB_PLAFOND3').AsFloat, NbDec); // Plafond Tranche 3
  Q.Next ;
  i:=i+1;
  Grille.RowCount:=Grille.RowCount+1;
  end ;
Ferme (Q);

SetControlText ('Lbl0', DoubleToCell(CC[1], NbDec));
SetControlText ('Lbl1', DoubleToCell(CC[2], NbDec));
SetControlText ('Lbl2', DoubleToCell(CC[3], NbDec));
SetControlText ('Lbl3', DoubleToCell(CC[4], NbDec));
SetControlText ('Lbl4', DoubleToCell(CC[5], NbDec));
SetControlText ('Lbl5', DoubleToCell(CC[6], NbDec));
SetControlText ('Lbl6', DoubleToCell(CC[7], NbDec));
SetControlText ('Lbl7', DoubleToCell(CC[8], NbDec));
end;

Initialization
registerclasses([TOF_PGBASE_BUL]);
end.
