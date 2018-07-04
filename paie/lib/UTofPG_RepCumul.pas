{***********UNITE*************************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 21/08/2001
Modifié le ... : 21/08/2001
Description .. : Gestion de la reprise des cumuls
Mots clefs ... : PAIE;PGIMPORT;CUMUL
*****************************************************************}
{
PT1-1 PH-VG 21/08/2001 V547 Modification du calcul des tranches et des plafonds
      des bases de cotisation
   -2 VG 31/08/2001 V547 La tranche 3 doit être limitée au plafond 3
PT2-1 VG 04/09/2001 V547 Formatage à droite des éléments de la grille
   -2 VG 04/09/2001 V547 Duplication automatique de la base et du plafond saisi
      pour la base située sur la 2ème ligne
   -3 VG 26/09/2001 V547 La duplication est désormais activée sur click d'un bouton
   -4 VG 28/09/2001 V547 Il faut que le bouton soit rendu visible !!!
   -5 VG 09/10/2001 V562 Duplcation => Duplication
PT3   PH 21/11/2002 V591 Prise en compte du ##PREDEFINI

}
unit UTofPG_RepCumul;

interface
uses  Windows,StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls,HTB97,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Fe_Main,DBGrids,
{$ELSE}
      MaineAgl,UtileAgl,
{$ENDIF}
      Grids,HCtrls,HEnt1,EntPaie,HMsgBox,UTOF,UTOB,Vierge,P5Def,P5Util,
      AGLInit ;
Type
     TOF_PG_RepCumul = Class (TOF)
       private
       QMul : TQUERY;     // Query recuperee du mul
       CodeSal,LeTitre : String;  // Code Salarie en traitement et Titre de la forme
       Modifier,GrilleAccessible : Boolean;
       AuDebut : Boolean; // Reprise en debut exercice
       DateDebut, DateFin : TDateTime; // Date de debut exercice social
       BtnPrev, BtnFirst, BtnLast, BtnNext, Dupliquer : TToolbarButton97; // Boutons geres par la TOF
       PageCtrl : TPageControl; //PT2-3
       TOB_LesCumuls, Tob_LesCumulSal, Tob_BaseSal, TOB_ProfilR, TOB_ProfilP,Tob_LesBases : TOB;
       T_Sal, T_Etab : TOB; // Tob du salarie et de son etablissement
       CGrille,BGrille  : THGrid; // Grilles de saisie des cumuls et des bases
       RowMax : integer; //PT2-2
       procedure PostDrawCellBase(ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState);
       procedure GrilleCellExit(Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
       procedure CalculRubrique(BGrille: THGrid; Base, Plafond: Double; ARow: Integer);
       procedure GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
       procedure FirstOnClick(Sender: TObject);
       procedure PrevOnClick(Sender: TObject);
       procedure NextOnClick(Sender: TObject);
       procedure LastOnClick(Sender: TObject);
       procedure DupliquerClick(Sender: TObject); //PT2-3
       procedure PageCtrlChange(Sender: TObject); //PT2-3
       Function  Bouge(Button: TNavigateBtn) : boolean ;
       procedure GereQuery;  // procedure de taitement du salarie à partir de la query du multicritere
       Function  EnregOK : boolean ;
       function  OnSauve: boolean;
       procedure ZoneSuivanteOuOk ( Var ACol,ARow : Longint ; Var Cancel : boolean ) ;
       Function  ZoneAccessible ( Var ACol,ARow : Longint) : Boolean;
       procedure ValiderClick (Sender: TObject);
       procedure BGrilleAcces (Sender: TObject);
       procedure FermeClick(Sender: TObject);
       public
       procedure OnArgument(Arguments : String ) ; override ;
     END ;

implementation


procedure TOF_PG_RepCumul.OnArgument(Arguments: String);
var F : TFVierge ;
    st, LeDebut, LaDated, LaDateF, Profil, Nat, Rub : STring;
    Q : TQuery ;
    T1,T : TOB;
    BtnValid,BtnFerme : TToolbarButton97;
    i : integer;
begin
inherited ;
TOB_LesCumuls := NIL;
Tob_LesCumulSal := NIL;
Tob_BaseSal := NIL;
TOB_ProfilP := NIL;
TOB_ProfilR := NIL;
Tob_LesBases := NIL;
T_Sal := NIL;
T_Etab := NIL;
LeTitre := Ecran.Caption ;
st:= Trim (Arguments);
CodeSal:=ReadTokenSt(st);   // Recup Code Salarie
LeDebut:=ReadTokenSt(st);  // Recup si reprise faite au debut de l'exercice pour filtrer les cumuls
if LeDebut = 'X' then AuDebut := TRUE else AuDebut := FALSE;
LaDateD:=ReadTokenSt(st);  // Recup Date debut
DateDebut := StrToDate (LaDateD);
LaDateF:=ReadTokenSt(st);  // Recup Date debut
DateFin := StrToDate (LaDateF);
RowMax := 1; //PT2-2

// recuperation de la query du multicritere
{$IFNDEF EAGLCLIENT}
if not (Ecran is TFVierge) then exit;
F:=TFVierge(Ecran) ;
if F <> NIL then  QMUL:=F.FMULQ ;
{$ELSE}
QMUL:=NIL;
{$ENDIF}
BtnValid:=TToolbarButton97 (GetControl ('BVALIDER'));
if BtnValid<>NIL then BtnValid.OnClick := ValiderClick ;
BtnFerme:=TToolbarButton97 (GetControl ('BFERME'));
if BtnFerme<>NIL then BtnFerme.OnClick := FermeClick ;
// design des grille et definition des tailles des colonnes
CGrille := THGrid(Getcontrol('GRILLECUM'));
if CGrille <> nil then
  begin
  CGrille.ColLengths [0]:=6;
  CGrille.ColLengths [1]:=37;
  CGrille.ColLengths [2]:=12;
// PT2-1
// valeurs des colonnes numériques cadrées à droite de la cellule
  CGrille.ColAligns[2] := taRightJustify;
// FIN PT2-1
  CGrille.OnCellExit := GrilleCellexit;
  end;

BGrille := THGrid(Getcontrol('GRILLEBAS'));
if BGrille <> nil then
  begin
  BGrille.ColLengths [0]:=6;
  BGrille.ColLengths [1]:=37;
{PT2-1
  BGrille.ColLengths [2]:=12;
  BGrille.ColLengths [3]:=12;
  BGrille.ColLengths [4]:=12;
  BGrille.ColLengths [5]:=12;
  BGrille.ColLengths [6]:=12;
  BGrille.ColLengths [7]:=12;
  BGrille.ColLengths [8]:=12;
}
// valeurs des colonnes numériques cadrées à droite de la cellule
  for i := 2 to 9 do
      begin
      BGrille.ColLengths [i]:=12;
      BGrille.ColAligns[i] := taRightJustify;
      end;
// FIN PT2-1

  BGrille.PostDrawCell  := PostDrawCellBase;
  BGrille.OnCellExit  := GrilleCellexit;
  BGrille.OnCellEnter := GrilleCellenter;
  BGrille.OnEnter := BGrilleAcces ;
  end;

// Gestion des boutons
{$IFDEF EAGLCLIENT}
SetControlVisible('BPREV',False);
SetControlVisible('BFIRST',False);
SetControlVisible('BNEXT',False);
SetControlVisible('BLAST',False);
{$ELSE}
BtnPrev:=TToolbarButton97 (GetControl ('BPREV'));
if BtnPrev <> NIL then BtnPrev.OnClick := PrevOnClick;
BtnFirst:=TToolbarButton97 (GetControl ('BFIRST'));
if BtnFirst <> NIL then BtnFirst.OnClick := FirstOnClick;
BtnNext:=TToolbarButton97 (GetControl ('BNEXT'));
if BtnNext <> NIL then BtnNext.OnClick := NextOnClick;
BtnLast:=TToolbarButton97 (GetControl ('BLAST'));
if BtnLast <> NIL then BtnLast.OnClick := LastOnClick;
{$ENDIF}
//PT2-3
Dupliquer:=TToolbarButton97 (GetControl('BDUPLIQUER'));
if Dupliquer <> NIL then
   begin
   Dupliquer.Hint := 'Duplication des bases et des plafonds'; // PT2-5
   Dupliquer.Visible := True;    //PT2-4
   Dupliquer.Enabled := False;
   Dupliquer.OnClick := DupliquerClick;
   end;
PageCtrl := TPageControl (GetControl('PGCTRL'));
if PageCtrl <> NIL then
   begin
   PageCtrl.ActivePageIndex := 0;
   PageCtrl.OnChange := PageCtrlChange;
   end;
//FIN PT2-3

// Gestion de la TOb des cumuls en fonction du type de reprise debut exercice
TOB_LesCumuls := TOB.Create ('Les Cumuls gérés', NIL, -1);
if AuDebut = TRUE then st:='SELECT * FROM CUMULPAIE WHERE ##PCL_PREDEFINI## PCL_RAZCUMUL<>"00" ORDER BY PCL_CUMULPAIE' //**// selection des cumuls tq periode raz <> debut exercice
 else st:='SELECT * FROM CUMULPAIE WHERE ##PCL_PREDEFINI## ORDER BY PCL_CUMULPAIE';   //**//
Q:=OpenSql(st, TRUE);
TOB_LesCumuls.LoadDetailDB('CUMULPAIE','','',Q,False) ;
Ferme (Q);
CGrille.RowCount := TOB_LesCumuls.Detail.count; // Dimension de la grille à 1 ligne Supp
// Gestion de la TOB des Bases de Cotisation
Tob_LesBases:=TOB.Create('Les Bases de Cotisation',Nil,-1) ;
st:='SELECT * FROM COTISATION WHERE ##PCT_PREDEFINI## PCT_NATURERUB="BAS" ORDER BY PCT_RUBRIQUE';//**//
Q:=OpenSql(st, TRUE);
Tob_LesBases.LoadDetailDB('COTISATIONS','','',Q,FALSE,False) ;
Ferme (Q);
// Gestion de la tob des profils gérés dans la paie
TOB_ProfilP:=TOB.Create('Les Profils',Nil,-1) ;
TOB_ProfilR:=TOB.Create('Les Rubriques des Profils',Nil,-1) ;

TOB_ProfilP.LoadDetailDB('PROFILPAIE','','',Nil,FALSE,False) ;
TOB_ProfilP.detail.Sort ('PPI_PROFIL');
// Gestion des rubriques associées  à chaque profil
st:='SELECT * FROM PROFILRUB WHERE ##PPM_PREDEFINI## PPM_NATURERUB="BAS" ORDER BY PPM_PROFIL,PPM_NATURERUB,PPM_RUBRIQUE';
Q:=OpenSql(st, TRUE);
TOB_ProfilR.LoadDetailDB('PROFILRUB','','',Q,FALSE,False) ;
Ferme (Q);
T1 := TOB_ProfilR.FindFirst ([''],[''], FALSE);
While T1 <> NIL do
 begin
 Profil:=T1.GetValue ('PPM_PROFIL');
 Nat:=T1.GetValue ('PPM_NATURERUB');
 Rub:=T1.GetValue ('PPM_RUBRIQUE');
 T := TOB_ProfilP.FindFirst (['PPI_PROFIl'],[profil], TRUE);
 if T <> NIL then  T1.ChangeParent (T, -1);
 T1 := TOB_ProfilR.FindNext ([''],[''], FALSE);
 end;
GereQuery; // traitement du 1er Salarie ou du salarie selectionne à partir du multicritere
end;

procedure TOF_PG_RepCumul.PostDrawCellBase(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var
T   : TOB;
Grise : Boolean;
begin
Grise := FALSE;
if ARow = 0 then exit;
if (ACol=0) OR (ACol=1) then exit;
T := Tob_LesBases.FindFirst (['PCT_NATURERUB','PCT_RUBRIQUE'],['BAS', BGrille.Cells[0, ARow]], FALSE);
if T <> NIL then
 begin
 if ((ACol = 4) OR (ACol = 7)) AND (T.GetValue ('PCT_TYPETRANCHE1') = '') then Grise := TRUE;
 if ((ACol = 5) OR (ACol = 8)) AND (T.GetValue ('PCT_TYPETRANCHE2') = '') then Grise := TRUE;
 if ((ACol = 6) OR (ACol = 9)) AND (T.GetValue ('PCT_TYPETRANCHE3') = '') then Grise := TRUE;
 end
 else GridGriseCell(BGrille,Acol,Arow,Canvas);
if GRISE then GridGriseCell(BGrille,Acol,Arow,Canvas);
// PT2-2
if (Arow>RowMax) then
   RowMax := Arow;
// FIN PT2-2
end;

procedure TOF_PG_RepCumul.GrilleCellExit(Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
var
st : String;
Mt, Mt2 : Double;
begin
if Acol = 0 then exit;
Modifier := TRUE;
if Sender = CGrille then St:=CGrille.Cells[ACol,ARow] else
 St:=BGrille.Cells[ACol,ARow];
Mt := Valeur (St);
Mt:=ARRONDI (Mt, 2);
if Sender = CGrille then CGrille.Cells[ACol, ARow]:=DoubleToCell(Mt,2)
 else BGrille.Cells[ACol, ARow]:=DoubleToCell(Mt,2);
if Mt = 0 then
 begin
 if Sender = CGrille then CGrille.Cells[ACol, ARow]:=''
  else BGrille.Cells[ACol, ARow]:='';
 end;

{ VG 09/03/01 Rajout des fonctions de calcul des tranches et des plafonds pour
la gestion des bases de cotisation}
if Sender = BGrille then
   BEGIN
   if Acol = 2 then
      BEGIN
      St:=BGrille.Cells[3,ARow];
      Mt2 := Valeur (St);
      Mt2:=ARRONDI (Mt2, 2);
      CalculRubrique(BGrille, Mt, Mt2, ARow);
// PT2-2
{PT2-3
      if (ARow = 2) then
         begin
         for i := 3 to RowMax do
             begin
             BGrille.Cells[2, i]:=BGrille.Cells[2,ARow];
             CalculRubrique(BGrille, Mt, Mt2, i);
             end;
         end;
}
// FIN PT2-2
      END;
   if Acol = 3 then
      BEGIN
      St:=BGrille.Cells[2,ARow];
      Mt2 := Valeur (St);
      Mt2:=ARRONDI (Mt2, 2);
      CalculRubrique(BGrille, Mt2, Mt, ARow);
// PT2-2
{
      if (ARow = 2) then
         begin
         for i := 3 to RowMax do
             begin
             BGrille.Cells[3, i]:=BGrille.Cells[3,ARow];
             CalculRubrique(BGrille, Mt2, Mt, i);
             end;
         end;
FIN PT2-3}         
// FIN PT2-2
      END;
   END;
{ VG 09/03/01 FIN}
end;

{ VG 09/03/01 Rajout des fonctions de calcul des tranches et des plafonds pour
             la gestion des bases de cotisation
  PT1-1 PH-VG 08/01 V547 Modification du calcul des tranches en fonction du paramètrage. Calcul tranche 3
}
procedure TOF_PG_RepCumul.CalculRubrique(BGrille: THGrid; Base, Plafond: Double; ARow: Integer);
var
Numero, StCot : String;
QRechCot : TQuery;
i : integer;
Typ1, Typ2, Typ3 : String;
Tranche1, Tranche2, Tranche3,Mt, Mt1,Mt2,Mt3 : double;
begin
Numero:=BGrille.Cells[0,ARow];
for i := 4 to 9 do
    BGrille.Cells[i, ARow] := DoubleToCell(0, 2);
// PT3   PH 21/11/2002 V591 Prise en compte du ##PREDEFINI
StCot:= 'SELECT PCT_TRANCHE1,PCT_TRANCHE2,PCT_TRANCHE3,PCT_TYPETRANCHE1,PCT_TYPETRANCHE2,PCT_TYPETRANCHE3'+
        ' FROM COTISATION WHERE ##PCT_PREDEFINI## '+
        ' PCT_RUBRIQUE = "'+Numero+'"';
QRechCot:=OpenSQL(StCot,TRUE);
// PORTAGE CWAS
if NOT QRechCot.EOF then
   begin
   Typ1 := QRechCot.FindField('PCT_TYPETRANCHE1').Asstring;
   Typ2 := QRechCot.FindField('PCT_TYPETRANCHE2').Asstring;
   Typ3 := QRechCot.FindField('PCT_TYPETRANCHE3').Asstring;
   end
else
   begin
   Typ1 := '';
   Typ2 := '';
   Typ3 := '';
   end;

Mt1:=0;
Mt2:=0;
Mt3:=0;
Tranche1:=0;
Tranche2:=0;
Tranche3:=0;

if (Typ1 = 'VAR') OR (Typ1 = '') then
   exit; // on ne traite pas car on ne peut pas evaluer une variable
if Typ1 = 'NBR' then
   begin
   Tranche1:=Valeur(QRechCot.FindField('PCT_TRANCHE1').Asstring);
   Mt1 := Plafond*Tranche1;
   end
else
   begin
   if Typ1 = 'ELN' then
      Mt1 := ValEltNat (QRechCot.FindField('PCT_TRANCHE1').Asstring, DateFin);
   end;

if Typ2 = 'NBR' then
   begin
   Tranche2:=Valeur(QRechCot.FindField('PCT_TRANCHE2').Asstring) - Tranche1;
   Mt2 := Plafond*Tranche2;
   end
else
   begin
   if Typ2 = 'ELN' then
      Mt2 := ValEltNat (QRechCot.FindField('PCT_TRANCHE2').Asstring, DateFin);
   end;
if Tranche2 < 0 then
   Tranche2 := 0;

if Typ3 = 'NBR' then
   begin
   Tranche3:=Valeur(QRechCot.FindField('PCT_TRANCHE3').Asstring) - Tranche2 - Tranche1;
   Mt3 := Plafond*Tranche3;
   end
else
   begin
   if Typ3 = 'ELN' then
      Mt3 := ValEltNat (QRechCot.FindField('PCT_TRANCHE3').Asstring, DateFin);
   end;
if Tranche3 < 0 then
   Tranche3 := 0;

BGrille.Cells[7, ARow]:=DoubleToCell(Mt1,2);
BGrille.Cells[8, ARow]:=DoubleToCell(Mt2,2);
BGrille.Cells[9, ARow]:=DoubleToCell(Mt3,2);
Ferme(QRechCot);

if (Mt1 < Base) then
   BEGIN
   BGrille.Cells[4, ARow]:=DoubleToCell(Mt1,2);
   Mt := Base-Plafond;
   if Mt > Mt2 then
      begin
      if Tranche3 <> 0 then
         begin
         if Typ2 <> 'DEP' then
            begin
            BGrille.Cells[5, ARow] := DoubleToCell(Mt2, 2);
{PT1-2
            BGrille.Cells[6, ARow] := DoubleToCell(Mt-Mt2,2);
}
            if (Mt-Mt2<Mt3) then
               BGrille.Cells[6, ARow] := DoubleToCell(Mt-Mt2,2)
            else
               BGrille.Cells[6, ARow] := DoubleToCell(Mt3,2);
{FIN PT1-2}
            end
         else
            begin
            BGrille.Cells[5, ARow] := DoubleToCell(Mt, 2);
            BGrille.Cells[6, ARow] := DoubleToCell(0,2);
            end;
         end
      else
         begin
         BGrille.Cells[5, ARow] := DoubleToCell(Mt2, 2);
         BGrille.Cells[6, ARow] := DoubleToCell(0,2);
         end;
      end
   else
      begin
      if Tranche2 <> 0 then
         begin
         if Typ2 <> 'DEP' then
            begin
            if Mt > Mt2 then
               BGrille.Cells[5, ARow] := DoubleToCell(Mt2,2)
            else
               BGrille.Cells[5, ARow] := DoubleToCell(Mt,2);
            end
         else
            BGrille.Cells[5, ARow] := DoubleToCell(Base - Mt1,2);
         end
      else
         BGrille.Cells[5, ARow] := DoubleToCell(0,2);
      BGrille.Cells[6, ARow] := DoubleToCell(0,2);
      end;

   END
else
   BEGIN
   BGrille.Cells[4, ARow]:=DoubleToCell(Base,2);
   BGrille.Cells[5, ARow]:=DoubleToCell(0,2);
   BGrille.Cells[6, ARow]:=DoubleToCell(0,2);
   END;

For i:=4 to 9 do
    BEGIN
    if (BGrille.Cells[i, ARow]='') then
       BGrille.Cells[i, ARow]:=DoubleToCell(0,2);
    END;
end;
{ FIN VG 09/03/01
  FIN PT1-1 PH-VG 08/01 V547 Modification du calcul des tranches
}


{Fonctions de gestion des boutons du navigateur
et fonctions des gestions des evenements lies à ces boutons
}
procedure TOF_PG_RepCumul.FirstOnClick(Sender: TObject);
begin
Bouge(nbFirst) ;  BtnLast.Enabled := TRUE; BtnNext.Enabled := TRUE;
BtnFirst.Enabled :=FALSE; BtnPrev.Enabled := FALSE;
end;

procedure TOF_PG_RepCumul.PrevOnClick(Sender: TObject);
begin
Bouge(nbPrior) ;
if QMul.BOF then begin BtnFirst.Enabled := FALSE; BtnPrev.Enabled := FALSE; end;
BtnLast.Enabled := TRUE;
BtnNext.Enabled := TRUE;
end;

procedure TOF_PG_RepCumul.NextOnClick(Sender: TObject);
begin
Bouge(nbNext) ;
if QMul.EOF then begin BtnLast.Enabled := FALSE; BtnNext.Enabled := FALSE; end;
BtnFirst.Enabled :=TRUE;
BtnPrev.Enabled := TRUE;
end;

procedure TOF_PG_RepCumul.LastOnClick(Sender: TObject);
begin
Bouge(nbLast) ;
BtnLast.Enabled := FALSE; BtnNext.Enabled := FALSE;
BtnFirst.Enabled :=TRUE; BtnPrev.Enabled := TRUE;
end;

Function TOF_PG_RepCumul.EnregOK : boolean ;
BEGIN
result:=FALSE  ; Modifier:=True ;
END;


Function TOF_PG_RepCumul.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve then Exit ;
   nbPost           : if Not EnregOK then Exit ;
//   nbDelete         : if Not OnDelete then Exit ;
   end ;
result:=TRUE ;
//if Button=NbInsert then begin NewEnreg ; exit; end;
if Button=NbPost then begin Modifier:=False ; exit; end;
{$IFNDEF EAGLCLIENT}
if Button=nbDelete then
  begin
  if QMul.EOF = FALSE then
    begin
    QMul.Next;
    if QMul.EOF = TRUE then
      begin
      QMul.prior ;
      if QMul.BOF then Close;
      end
    end
  else
    begin
    if QMul.BOF = FALSE then
      begin
      QMul.prior;
      if QMul.BOF = TRUE then Close;
      end;
    end;
  end;
if QMul <> NIL then
  begin
  Case Button of
  nblast   : QMul.Last;
  nbfirst  : QMul.First;
  nbnext   : QMul.Next;
  nbprior  : QMul.prior;
  end;
  end;
{$ENDIF}  
GereQuery;
END ;
// Fonction de validation de la saisie
procedure TOF_PG_RepCumul.ValiderClick(Sender: TObject);
Var Old : Boolean ;
    rep : Integer;
begin
inherited ;
Old:=Modifier ; Modifier:=FALSE ;
if Not Bouge(nbPost) then Modifier:=Old;
Rep:=PGIAsk ('Voulez vous sauvegarder votre saisie ?', 'Saisie des cumuls de reprise') ;
if rep=mrNo then exit ;
if rep=mrCancel then exit;
if rep=mryes then OnSauve;
end;
// fermeture de la forme
procedure TOF_PG_RepCumul.FermeClick(Sender: TObject);
begin
ValiderClick(Sender);
Modifier := FALSE;
if TOB_LesCumuls <> NIL then begin TOB_LesCumuls.free; TOB_LesCumuls:=NIL; end;
if Tob_LesCumulSal <> NIL then begin Tob_LesCumulSal.free; Tob_LesCumulSal:=NIL; end;
if Tob_BaseSal <> NIL then begin Tob_BaseSal.free; Tob_BaseSal:=NIL; end;
if TOB_ProfilR <> NIL then begin TOB_ProfilR.free; TOB_ProfilR:=NIL; end;
if TOB_ProfilP <> NIL then begin TOB_ProfilP.free; TOB_ProfilP:=NIL; end;
if Tob_LesBases <> NIL then begin Tob_LesBases.free; Tob_LesBases:=NIL; end;
if T_Sal <> NIL then begin T_Sal.free; T_Sal:=NIL; end;
if T_Etab <> NIL then begin T_Etab.free; T_Etab:=NIL; end;
Close;
end;


//PT2-3
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 26/09/2001
Modifié le ... :   /  /
Description .. : Duplication d'une base
Mots clefs ... : PAIE;PGIMPORT;CUMUL
*****************************************************************}
procedure TOF_PG_RepCumul.DupliquerClick(Sender: TObject);
var
st : String;
Mt, Mt2 : Double;
i : integer; // PT2-2
begin
if GrilleAccessible = FALSE then
   begin
   PGIBox ('Vous ne pouvez pas saisir vos reprises de bases de cotisations#13#10 car vous avez déjà fait des bulletins', 'Reprises des bases de cotisations');
   BGrille.Enabled := FALSE;
   end
else
   begin
   BGrille.Enabled := true;
   St:=BGrille.Cells[2,2];
   Mt := Valeur (St);
   Mt:=ARRONDI (Mt, 2);
   St:=BGrille.Cells[3,2];
   Mt2 := Valeur (St);
   Mt2:=ARRONDI (Mt2, 2);
   for i := 3 to RowMax do
       begin
       BGrille.Cells[2, i]:=BGrille.Cells[2,2];
       BGrille.Cells[3, i]:=BGrille.Cells[3,2];
       CalculRubrique(BGrille, Mt, Mt2, i);
       end;
   end;
end;


procedure TOF_PG_RepCumul.PageCtrlChange(Sender: TObject);
begin
if Dupliquer <> NIL then
   begin
   if PageCtrl.ActivePageIndex = 0 then
      Dupliquer.Enabled := False
   else
      Dupliquer.Enabled := True;
   end;
end;
//FIN PT2-3

{ Fonction de chargement des differentes tob en fonction du salarie
TOB des cumuls de reprise du salarie en fonction de la tob des cumuls selectionnes
TOB historique des bases de cotisations
Pour les bases de cotisation, il faut parcourir les differents profils du salarie pour trouver les
bases de cotisations
Attention : la date de debut et de  fin sera tjrs dans le cas d'une reprise la date de debut de
l'exercice social. Ce qui permet de modifier ou supprimer les saisies déjà effectuées.
}
procedure TOF_PG_RepCumul.GereQuery;
var
st : String;
Q  : TQuery;
TB, TC, TPP, Salarie, Etab : TOB; // definition des tob contenant les éléments dejà saisis
begin
if T_Sal <> NIL then begin T_Sal.Free; T_Sal := NIL; end;
if T_Etab <> NIL then begin T_Etab.Free; T_Etab := NIL end;
Modifier := FALSE;
{$IFNDEF EAGLCLIENT}
if QMul = NIL then exit;
CodeSal:=QMul.FindField ('PSA_SALARIE').AsString;
{$ENDIF}
if Tob_LesCumulSal <> NIL then begin Tob_LesCumulSal.free; Tob_LesCumulSal := NIL; end;
if Tob_BaseSal <>  NIL then begin Tob_BaseSal.free; Tob_BaseSal := NIL; end;
//BGrille.RowCount := TOB_LesCumuls.Detail.count; // Dimension de la grille à 1 ligne Supp
// Recup tob du salarie
T_Sal := TOB.Create ('Le salarie', NIL, -1);
st:='SELECT * FROM SALARIES WHERE PSA_SALARIE="'+CodeSal+'"';
Q:=OpenSql(st, TRUE);
T_Sal.LoadDetailDB('SALARIES','','',Q,False) ;
Ferme (Q);
Salarie := T_Sal.FindFirst (['PSA_SALARIE'],[CodeSal], FALSE);
// Recup tob de l'etablissement du salarie
Ecran.Caption := LeTitre+' '+CodeSal+' '+Salarie.GetValue ('PSA_LIBELLE')+' pour la période du '+DateToStr(DateDebut)+' au '+DateToStr(DateFin);
T_Etab := TOB.Create ('Etablissement du salarie', NIL, -1);
st:='SELECT * FROM ETABCOMPL WHERE ETB_ETABLISSEMENT="'+Salarie.GetValue ('PSA_ETABLISSEMENT')+'"';
Q:=OpenSql(st, TRUE);
T_Etab.LoadDetailDB('ETABCOMPL','','',Q,False) ;
Ferme (Q);
Etab := T_Etab.FindFirst (['ETB_ETABLISSEMENT'],[Salarie.GetValue ('PSA_ETABLISSEMENT')], FALSE);
// recup liste des cumuls salaries déjà saisis
Tob_LesCumulsal := TOB.Create ('Les Cumuls salarie', NIL, -1);
st:='SELECT * FROM HISTOCUMSAL WHERE PHC_REPRISE = "X" AND PHC_SALARIE="'+CodeSal+'" AND PHC_DATEDEBUT="'+USDateTime(DateDebut)+'" AND PHC_DATEFIN="'+USDateTime(DateFin)+'" ORDER BY PHC_CUMULPAIE';
Q:=OpenSql(st, TRUE);
Tob_LesCumulsal.LoadDetailDB('HISTOCUMSAL','','',Q,False) ;
Ferme (Q);
TPP:=TOB_LesCumuls.FindFirst([''],[''],FALSE) ;
while TPP<>Nil  do
 begin
 TC := Tob_LesCumulsal.FindFirst (['PHC_CUMULPAIE'], [TPP.GetValue('PCL_CUMULPAIE')], FALSE);
 if TC = NIL then // Creation d'une ligne cumul de reprise afin de pouvoir eventuellement saisir un montant
  begin
  TC:=TOB.Create('HISTOCUMSAL',Tob_LesCumulsal,-1) ;
  TC.PutValue ('PHC_ETABLISSEMENT', Salarie.GetValue ('PSA_ETABLISSEMENT'));
  TC.PutValue ('PHC_SALARIE', CodeSal);
  TC.PutValue ('PHC_DATEDEBUT',DateDebut);
  TC.PutValue ('PHC_DATEFIN',DateFin);
  TC.PutValue ('PHC_CUMULPAIE',TPP.GetValue('PCL_CUMULPAIE'));
  TC.PutValue ('PHC_REPRISE', 'X');
  TC.PutValue ('PHC_MONTANT', 0);
  end;
 TC.PutValue ('PHC_CONFIDENTIEL', Salarie.GetValue ('PSA_CONFIDENTIEL'));
 TC.PutValue ('PHC_TRAVAILN1', Salarie.GetValue ('PSA_TRAVAILN1'));
 TC.PutValue ('PHC_TRAVAILN2', Salarie.GetValue ('PSA_TRAVAILN2'));
 TC.PutValue ('PHC_TRAVAILN3', Salarie.GetValue ('PSA_TRAVAILN3'));
 TC.PutValue ('PHC_TRAVAILN4', Salarie.GetValue ('PSA_TRAVAILN4'));
 TC.PutValue ('PHC_CODESTAT', Salarie.GetValue ('PSA_CODESTAT'));
 TC.PutValue ('PHC_LIBREPCMB1', Salarie.GetValue ('PSA_LIBREPCMB1'));
 TC.PutValue ('PHC_LIBREPCMB2', Salarie.GetValue ('PSA_LIBREPCMB2'));
 TC.PutValue ('PHC_LIBREPCMB3', Salarie.GetValue ('PSA_LIBREPCMB3'));
 TC.PutValue ('PHC_LIBREPCMB4', Salarie.GetValue ('PSA_LIBREPCMB4'));
 TC.AddChampSup ('LIBELLE', FALSE);
 TC.PutValue ('LIBELLE', TPP.Getvalue ('PCL_LIBELLE'));
 TPP:=TOB_LesCumuls.FindNext([''],[''],FALSE) ;
 end;
Tob_LesCumulsal.Detail.Sort('PHC_CUMULPAIE') ;

// recup liste des bases salaries déjà saisies
Tob_BaseSal := TOB.Create ('Les bases salarie', NIL, -1);
st:='SELECT * FROM HISTOBULLETIN WHERE PHB_SALARIE="'+CodeSal+'" AND PHB_DATEDEBUT="'+USDateTime(DateDebut)+'" AND PHB_DATEFIN="'+USDateTime(DateFin)+'" AND PHB_NATURERUB = "BAS" ORDER BY PHB_RUBRIQUE' ;
Q:=OpenSql(st, TRUE);
Tob_BaseSal.LoadDetailDB('HISTOBULLETIN','','',Q,False) ;
Ferme (Q);
// recup de la TOB liste des bases de cotisation theoriques contenues dans les différents profil duj salarie
TB := TOB.Create ('Les bases theoriques des profils salarie', NIL, -1);
RechercheRubriqueProfil (Salarie,Etab,TB,TOB_ProfilP,Tob_LesBases, TRUE, DateDebut, DateFin);
TPP:=TB.FindFirst([''],[''],FALSE) ;
while TPP<>Nil  do
 begin
 TC := Tob_BaseSal.FindFirst (['PHB_NATURERUB','PHB_RUBRIQUE'], ['BAS',TPP.GetValue('PHB_RUBRIQUE')], FALSE);
 if TC = NIL then  TPP.ChangeParent (Tob_BaseSal, -1);
 TPP:=TB.FindNext([''],[''],FALSE) ;
 end;
Tob_BaseSal.Detail.Sort ('PHB_NATURERUB;PHB_RUBRIQUE');
Tob_LesCumulsal.PutGridDetail (CGrille,FALSE,FALSE,'PHC_CUMULPAIE;LIBELLE;PHC_MONTANT',FALSE);
Tob_BaseSal.PutGridDetail (BGrille,FALSE,FALSE,'PHB_RUBRIQUE;PHB_LIBELLE;PHB_BASECOT;PHB_PLAFOND;PHB_TRANCHE1;PHB_TRANCHE2;PHB_TRANCHE3;PHB_PLAFOND1;PHB_PLAFOND2;PHB_PLAFOND3',FALSE);
BGrille.RowCount := Tob_BaseSal.Detail.count + 1;
GrilleAccessible := TRUE;
st := 'SELECT Count (*) FROM HISTOCUMSAL WHERE PHC_SALARIE="'+CodeSal+'" AND PHC_REPRISE <> "X"';
Q:=OpenSql(st, TRUE);
if NOT Q.EOF then
   begin
   if Q.Fields[0].AsInteger > 0 then GrilleAccessible := FALSE;
   end;
Ferme (Q);

end;
// fonction d'ecriture des elements saisis
function TOF_PG_RepCumul.OnSauve: boolean;
var
st : String;
T  : TOB;
begin
result := TRUE;
if Modifier = FALSE then exit;
try
BeginTrans;
Modifier  := FALSE;
st:='DELETE FROM HISTOCUMSAL WHERE PHC_SALARIE="'+CodeSal+'" AND PHC_REPRISE="X" AND PHC_DATEDEBUT="'+USDateTime(DateDebut)+'" AND PHC_DATEFIN="'+USDateTime(DateFin)+'"';
ExecuteSQL(st) ;
if GrilleAccessible = true then
   begin
   st:='DELETE FROM HISTOBULLETIN WHERE PHB_SALARIE="'+CodeSal+'" AND PHB_DATEDEBUT="'+USDateTime(DateDebut)+'" AND PHB_DATEFIN="'+USDateTime(DateFin)+'"';
   ExecuteSQL(St) ;
   end;
// Recup des montants saisis dans les 2 grilles
Tob_LesCumulsal.GetGridDetail (CGrille,-1,'','PHC_CUMULPAIE;LIBELLE;PHC_MONTANT');
Tob_LesCumulsal.DelChampSup ('LIBELLE',TRUE);
Tob_BaseSal.GetGridDetail (BGrille,-1,'','PHB_RUBRIQUE;PHB_LIBELLE;PHB_BASECOT;PHB_PLAFOND;PHB_TRANCHE1;PHB_TRANCHE2;PHB_TRANCHE3;PHB_PLAFOND1;PHB_PLAFOND2;PHB_PLAFOND3');
// Elimination des lignes dont tous les montants sont nuls cad sans saisie et stockage dans la base
T :=Tob_LesCumulsal.FindFirst ([''],[''],FALSE);
While T <> NIL do
 begin
 if T.GetValue ('PHC_MONTANT') = 0  then
  T.Free;
 T :=Tob_LesCumulsal.FindNext ([''],[''],FALSE);
 end;
Tob_LesCumulsal.SetAllModifie (TRUE);
Tob_LesCumulsal.InsertDB(Nil,FALSE) ;
if GrilleAccessible = true then
   begin
   T :=Tob_BaseSal.FindFirst ([''],[''],FALSE);
   While T <> NIL do
         begin
         if (T.GetValue ('PHB_BASECOT') = 0) AND (T.GetValue ('PHB_TRANCHE1') = 0) AND (T.GetValue ('PHB_TRANCHE2') = 0) AND
            (T.GetValue ('PHB_TRANCHE3') = 0) AND (T.GetValue ('PHB_PLAFOND') = 0) AND (T.GetValue ('PHB_PLAFOND1') = 0) AND (T.GetValue ('PHB_PLAFOND2') = 0) AND (T.GetValue ('PHB_PLAFOND3') = 0)
            then T.Free;
         T :=Tob_BaseSal.FindNext ([''],[''],FALSE);
         end;
   Tob_BaseSal.SetAllModifie (TRUE);
   Tob_BaseSal.InsertDB(Nil,FALSE) ;
  end;
CommitTrans;
Except
result := FALSE;
Rollback;
PGIBox ('Une erreur est survenue lors de la validation de la saisie',LeTitre);
end;
Modifier := FALSE;
end;


procedure TOF_PG_RepCumul.GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
Modifier := TRUE;
ZoneSuivanteouOk (ACol,ARow,cancel);
end;

procedure TOF_PG_RepCumul.ZoneSuivanteOuOk ( Var ACol,ARow : Longint ; Var Cancel : boolean ) ;
Var Sens,ii : integer ;
    OldEna  : boolean ;
BEGIN
OldEna:=BGrille.SynEnabled ; BGrille.SynEnabled:=False ;
Sens:=-1 ; if BGrille.Row>ARow then Sens:=1 else if ((BGrille.Row=ARow) and (ACol<BGrille.Col)) then Sens:=1 ;
if (sens=-1) AND ((ACol=0) OR (BGrille.Col=0))AND (ARow=1)  then begin BGrille.SynEnabled:=OldEna ; Cancel:=TRUE; exit; end;
ACol:=BGrille.Col ; ARow:=BGrille.Row ; ii:=0 ;
While Not ZoneAccessible(ACol,ARow)  do
   BEGIN
   Cancel:=True ; inc(ii) ; if ii>1000 then Break ;
   if Sens=1 then
      BEGIN
      if ((ACol=BGrille.ColCount-1) and (ARow=BGrille.RowCount-1)) then begin ACol:=BGrille.FixedCols; ARow := 1; Break ; end;
      if ACol<BGrille.ColCount-1 then Inc(ACol) else BEGIN Inc(ARow) ; ACol:=BGrille.FixedCols ; END ;
      END else
      BEGIN
      if ((ACol=BGrille.FixedCols) and (ARow=1)) then Break ;
      if ACol>BGrille.FixedCols then Dec(ACol) else BEGIN Dec(ARow) ; ACol:=BGrille.ColCount-1 ; END ;
      END ;
   END ;
BGrille.SynEnabled:=OldEna ;
END ;

Function TOF_PG_RepCumul.ZoneAccessible ( Var ACol,ARow : Longint) : Boolean;
var T : TOB;
begin
result:=TRUE;
if (Acol=0) OR (ACol=1) then
   begin
   result:=FALSE;
   exit;
   end;
T := Tob_LesBases.FindFirst (['PCT_NATURERUB','PCT_RUBRIQUE'],
                             ['BAS', BGrille.Cells[0, ARow]], FALSE);
if T <> NIL then
   begin
   if ((ACol = 4) OR (ACol = 7)) AND (T.GetValue ('PCT_TYPETRANCHE1') = '') then
      result := FALSE;
   if ((ACol = 5) OR (ACol = 8)) AND (T.GetValue ('PCT_TYPETRANCHE2') = '') then
      result := FALSE;
   if ((ACol = 6) OR (ACol = 9)) AND (T.GetValue ('PCT_TYPETRANCHE3') = '') then
      result := FALSE;
   end
else
   result := FALSE;
end;

procedure TOF_PG_RepCumul.BGrilleAcces(Sender: TObject);
begin
if GrilleAccessible = FALSE then
   begin
   PGIBox ('Vous ne pouvez pas saisir vos reprises de bases de cotisations#13#10 car vous avez déjà fait des bulletins', 'Reprises des bases de cotisations');
   BGrille.Enabled := FALSE;
   end
   else BGrille.Enabled := true;
end;

Initialization
registerclasses([TOF_PG_RepCumul]);
end.
