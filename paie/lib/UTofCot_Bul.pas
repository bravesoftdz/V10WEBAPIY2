{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 25/06/2001
Modifié le ... : 25/06/2001
Description .. : Affichage pendant la saisie du bulletin sur click droit
Suite ........ : des lignes cotisations pour un salarié pour une cotisation
Mots clefs ... : PAIE;PGBULLETIN
*****************************************************************}
{
PT1 : 31/03/2006 PH V_65 FQ 12568 Affichage Code+libelle dans la tablette
PT2 : 31/03/2006 PH V_65 FQ 12225 Prise en compte paie validée du mois

}
unit UTofCot_Bul;

interface
uses  StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DBGrids,
{$ELSE}

{$ENDIF}
      uPaieCotisations,
      Grids,HCtrls,HEnt1,HMsgBox,UTOF,UTOB,Vierge,P5Util,P5Def,AGLInit;
Type
     TOF_PGCot_Bul = Class (TOF)
       private
       CodeSal, Etab, DateD, DateF: String;
       DD, DF : TDateTime;
       Grille : THGrid;
       VCbxLRub : THValComboBox;
       procedure RechHisto(Sender: TObject);
       procedure PasDeSaisieRow(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
       procedure PasDeSaisieCell(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
//       procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
       public
       procedure OnArgument(Arguments : String ) ; override ;
     END ;

implementation

(*
procedure TOF_PGCot_BUL.GetCellCanvas(Acol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
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
*)

procedure TOF_PGCot_BUL.OnArgument(Arguments: String);
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
RendDateExerSocial (StrToDate(DateD),StrToDate(DateF), DD,DF, TRUE); // PT2
Btn := TToolbarButton97(GetControl('RECHHISTO'));
if Btn <> NIL then Btn.OnClick := RechHisto;

Grille :=THGrid (GetControl('GRILLECOT')) ;
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
   VCbxLRub.Items.Add(T1.GetValue('CODEBASE')+' '+T1.GetValue ('LIBELBASE')) ;  // PT1
   VCbxLRub.Values.Add(T1.GetValue ('CODEBASE')) ;
   end ;
   VCbxLRub.ItemIndex:=0;
 end;
// Design canvas de cellule ==> Grille.GetCellCanvas:=GetCellCanvas;
end;

procedure TOF_PGCot_BUL.PasDeSaisieCell(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
exit;
end;

procedure TOF_PGCot_BUL.PasDeSaisieRow(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
exit;
end;

procedure TOF_PGCot_BUL.RechHisto(Sender: TObject);
var Q : TQuery;
    St,Rub : String;
    i, NbDec : Integer;
    T2 : TOB;
    CC : Array [1..2] of Double;
begin
Rub:=VCbxLRub.Value;
for i:=1 to 2 do
 begin
 CC[i]:=0;
 end;
St:='SELECT PHB_DATEDEBUT,PHB_DATEFIN,PHB_BASECOT,PHB_TAUXSALARIAL,PHB_MTSALARIAL,PHB_TAUXPATRONAL,PHB_MTPATRONAL'+
    ' FROM HISTOBULLETIN WHERE PHB_SALARIE="'+CodeSal+'" AND PHB_NATURERUB="COT" AND PHB_RUBRIQUE="'+Rub+
    '" AND PHB_DATEDEBUT>="'+UsDateTime(DD)+'" AND PHB_DATEFIN<="'+UsDateTime(DF)+'" ORDER BY PHB_DATEFIN';
Q:=OpenSQL(St,TRUE) ;
T2:=TOB_Cotisations.FindFirst (['PCT_RUBRIQUE'],[Rub], TRUE);
NbDec:=2;
Grille.RowCount:=2;
if T2 <> NIL then NbDec:=T2.GetValue('PCT_DECBASECOT');
i:=1;
// Boucle sur les 12 mois maxi de la paie dans un exercice soit 12 lignes remplies dans la grille
While Not Q.EOF do
  begin
  Grille.Cells[0,i]:= DateToStr(Q.FindField('PHB_DATEDEBUT').AsFloat);  // DateDebut
  Grille.Cells[1,i]:= DateToStr(Q.FindField('PHB_DATEFIN').AsFloat);  // DateFin
  CC[1]:=CC[1]+Q.FindField('PHB_MTSALARIAL').AsFloat;
  CC[2]:=CC[2]+Q.FindField('PHB_MTPATRONAL').AsFloat;
  Grille.Cells[2,i]:= DoubleToCell (Q.FindField('PHB_BASECOT').AsFloat, NbDec); // Base de Cotisation
  Grille.Cells[3,i]:= DoubleToCell (Q.FindField('PHB_TAUXSALARIAL').AsFloat, NbDec); // taux SAl
  Grille.Cells[4,i]:= DoubleToCell (Q.FindField('PHB_MTSALARIAL').AsFloat, NbDec); // Mt Sal
  Grille.Cells[5,i]:= DoubleToCell (Q.FindField('PHB_TAUXPATRONAL').AsFloat, NbDec); // taux Pat
  Grille.Cells[6,i]:= DoubleToCell (Q.FindField('PHB_MTPATRONAL').AsFloat, NbDec); // Mt Pat
  Q.Next ;
  i:=i+1;
  Grille.RowCount:=Grille.RowCount+1;
  end ;
Ferme (Q);

SetControlText ('Lbl2', DoubleToCell(CC[1], NbDec));
SetControlText ('Lbl4', DoubleToCell(CC[2], NbDec));
end;

Initialization
registerclasses([TOF_PGCOT_BUL]);
end.
