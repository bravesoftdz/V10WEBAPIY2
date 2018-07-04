{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 22/08/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : CONVTYPOJAL ()
Mots clefs ... : TOF;CONVTYPOJAL
*****************************************************************}
Unit UTOFCONVTYPOJAL ;

Interface

Uses StdCtrls, Controls, Classes, db, forms, sysutils, dbTables, ComCtrls,
     HCtrls, HEnt1, Ent1, HMsgBox, UTOF , TOFMeth , Journal,
     HPanel , HTB97, HDB, UTob , Mul, SaisUtil, HQry, FE_Main, Grids ;
const
  Titre = 'Conversion Typologie Journal' ;

var LibCptAtt : String ;


Type
  TOF_CONVTYPOJAL = Class (TOF)
    BOuvrir : TToolbarButton97 ;
    BCherche : TToolbarButton97 ;
    ListeJal : THDbGrid  ;
    QListeJal : TQuery ;
    Pages : TPageControl ;
    TOBJalSelect : TOB ;
    TobDev : Tob ;
    procedure OnArgument (S : String ) ; Override ;
    procedure OnLoad                   ; Override ;
    procedure OnClose                  ; Override ;
    procedure NewModeOnChange ( sender : TObject ) ;
    procedure ModeSaisieOnClick (sender : TObject ) ;
    procedure ModeSaisieOnExit (sender : TObject ) ;
    procedure RechercheDecDev(Var DecDev : integer ; TobDet :Tob) ;
    procedure MarqueOrigine  ;
    procedure PrechargeOrigines  ;
    procedure ToutMarquer  ;
    procedure VireInutiles  ;
    procedure RecupMultiSel ;
    Procedure MajEcrCompta( TobDet : TOB ; NP, NL, NG : Integer ) ;
    Procedure MajEcrAna(var TOBdet : TOB ; NP, NL : Integer) ;
    Procedure MajSoldeEcr (var TobDet : TOB ; SE, SD : Double) ;
    Function LigSuivEcartEuro (TT : TOB ; Cpt : integer ) : Boolean ;
    Procedure ConvVersLibreOuBor (var TobEcr : Tob);
    Procedure ConvBorToPie (var TobEcr : Tob ; CodJal : string );
    procedure MajTotaux(Quoi, CodJal, Exo: String; Debit, Credit: Double);
    Procedure CreerLigCptAttente (TOBDet : TOB ; Solde, SoldeEuro, SoldeDev : Double ; NumLig, NumPiece, NumGroupe : Integer ) ;
    Procedure ConvLibToBor (var TobEcr : Tob ;CodJal : string; LibToBor : Boolean) ;
    procedure RechercherEcriture(var TobEcr: Tob; CodJal: String; DateDeb,DateFin: TDateTime; Libre : boolean);
    procedure ConversionJournal ;
    procedure BOuvrirOnClick (Sender : TObject) ;
    procedure FListeOnDblClick(sender: TObject);
  end ;


Implementation


var MaxFolio : integer ;
    TobCptAttente : Tob ;

procedure TOF_CONVTYPOJAL.OnClose ;
begin
  TobDev.free ;
  Inherited ;
end ;

procedure TOF_CONVTYPOJAL.OnLoad  ;
Begin
// Chargement des devises dans la TOB
TobDev.LoadDetailDB('DEVISE', '', '', Nil, False, True ) ;
Inherited ;
End ;

procedure TOF_CONVTYPOJAL.OnArgument (S : String ) ;
var Fiche : TFMul ;
begin
Inherited ;
Fiche := TFMul(Ecran) ;
ListeJal :=Fiche.FListe ;
QListeJal := Fiche.Q ;

MaxFolio := -1 ;

TobDev := Tob.Create('DEVISE', Nil, -1) ;

BOuvrir := TToolBarButton97(GetControl('BOUVRIR')) ;
BCherche := TToolBarButton97(GetControl('BCHERCHE')) ;
Pages := TPageControl(GetControl('Pages')) ;
THValComboBox(GetControl('J_MODESAISIE')).ItemIndex := 0 ;
THValComboBox(GetControl('J_NATUREJAL')).ItemIndex := 0 ;
SetControlProperty ( 'NEWMODE' , 'Plus' , 'And CO_CODE <> "' + GetControlText('J_MODESAISIE')+ '"' ) ;

THValComboBox(GetControl('NEWMODE')).OnChange := NewModeOnChange ;
THValComboBox(GetControl('J_MODESAISIE')).OnClick := ModeSaisieOnClick ;
THValComboBox(GetControl('J_MODESAISIE')).OnExit := ModeSaisieOnExit ;

THDBGrid(GetControl('FLISTE')).OnDblClick := FListeOnDblClick ;

if BOuvrir <> Nil then
   Begin
   BOuvrir.OnClick := BOuvrirOnClick ;
   end ;
END ;

procedure TOF_CONVTYPOJAL.NewModeOnChange ( Sender : TObject ) ;
var Q : TQuery ;
    ReqSql : string ;
begin
  // Folio accessible seulement pour mode Bordereau ou Libre
  if (GetControlText('NEWMODE')='BOR') Or (GetControlText('NEWMODE')='LIB') then
    BEGIN
    if MaxFolio = -1 then
      begin
      // recherche du premier n° de folio à utiliser

      ExecuteSQL('CREATE NONCLUSTERED INDEX [ms] ON [dbo].[ecriture]([e_modesaisie])') ; // ESSAI RAPIDITE

      ReqSql := 'SELECT MAX(E_NUMEROPIECE) FROM ECRITURE WHERE E_MODESAISIE="BOR" OR E_MODESAISIE="LIB" ' ;
      Q := OpenSql(ReqSql,True) ;
      MaxFolio := Q.Fields[0].AsInteger + 1 ;
      Ferme(Q) ;

      ExecuteSql('DROP INDEX ECRITURE.MS') ; // ESSAI RAPIDITE

      SetControlProperty('FOLIO','MinValue',MaxFolio) ;
      end ;
    setControlText('FOLIO',IntToStr(MaxFolio)) ;
    SetControlProperty('FOLIO','Enabled', True );                              
    SetControlProperty('TFOLIO','Enabled', True );
    END
  else
    BEGIN
    SetControlProperty('FOLIO','Enabled', false );
    SetControlProperty('TFOLIO','Enabled', false );
    END  ;
end ;

procedure TOF_CONVTYPOJAL.ModeSaisieOnClick (sender : TObject ) ;
begin
  // réinitialisation des nouveaux param au changement du mode de saisie d'origine
  SetControlProperty('NEWMODE', 'Plus', 'And Co_Code <> "' + GetControlText('J_MODESAISIE') + '" ') ;
  SetControlText('NEWMODE','') ;
  SetControlProperty('FOLIO','Enabled', false );
  SetControlProperty('TFOLIO','Enabled', false );
end ;

procedure TOF_CONVTYPOJAL.ModeSaisieOnExit (sender : TObject ) ;
begin
  SetControlProperty('NEWMODE', 'Plus', 'And Co_Code <> "' + GetControlText('J_MODESAISIE') + '" ') ;
end ;

procedure TOF_CONVTYPOJAL.RechercheDecDev(Var DecDev : integer ; TobDet : Tob) ;
Var DevTrouve : boolean ;
    CptDev : Integer ;
begin
  DevTrouve := false ;
  CptDev := 0 ;
  // recherche de la devise dans la TOB
  While (CptDev < TobDet.detail.count) and not DevTrouve do
    Begin
    DevTrouve := (TobDet.GetValue('E_DEVISE'))= (TobDev.detail[CptDev].GetValue('D_DEVISE'));
    Inc(CptDev) ;
    end ;
  // récuperation du nombre de décimales de la devise si trouvée
  If DevTrouve then DecDev := TobDev.detail[CptDev-1].GetValue('D_DECIMALE') ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 31/08/2001
Modifié le ... :   /  /
Description .. : Charge dans la TOB la liste des journaux suivant les critères
Suite ........ : du MUL.
Mots clefs ... :
*****************************************************************}
procedure TOF_CONVTYPOJAL.PrechargeOrigines ;
Var QQ : TQuery ;
    st : String ;
BEGIN
TobJalSelect.ClearDetail ;
St:='SELECT J_JOURNAL,J_MODESAISIE, J_TOTALCREDIT,J_TOTALDEBIT,J_TOTCREE,J_TOTCREP,J_TOTCRES,J_TOTDEBE,J_TOTDEBP,J_TOTDEBS from JOURNAL ' ;
QQ:=OpenSQL(St+RecupWhereCritere(Pages),True) ;
TOBJalSelect.LoadDetailDB('JOURNAL','','',QQ,False,True) ;
Ferme(QQ) ;
if TOBJalSelect.Detail.Count>0 then TOBJalSelect.Detail[0].AddChampSup('MARQUE',True) ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 31/08/2001
Modifié le ... :   /  /
Description .. : Marquage dans la TOB du journal sélectionné dans la liste.
Mots clefs ... :
*****************************************************************}
procedure TOF_CONVTYPOJAL.MarqueOrigine ;
Var TOBL : TOB ;
BEGIN
TOBL := TOBJalSelect.FindFirst(['J_JOURNAL'],[QListeJal.FindField('J_JOURNAL').AsString],False) ;
if TOBL<>Nil then TOBL.PutValue('MARQUE','X') ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 31/08/2001
Modifié le ... :   /  /
Description .. : supprime de la TOB les journaux non marqués.
Mots clefs ... :
*****************************************************************}
procedure TOF_CONVTYPOJAL.VireInutiles ;
Var i : integer ;
    TOBL  : TOB ;
BEGIN
for i:=TOBJalSelect.Detail.Count-1 downto 0 do
  BEGIN
  TOBL:=TOBJalSelect.Detail[i] ;
  if TOBL.GetValue('MARQUE')<>'X' then BEGIN TOBL.Free ; TOBL:=Nil ; END ;
  END ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 31/08/2001
Modifié le ... :   /  /
Description .. : Marque tous les journaux de la TOB.
Mots clefs ... :
*****************************************************************}
procedure TOF_CONVTYPOJAL.ToutMarquer ;
Var i : integer ;
    TOBL : TOB ;
BEGIN
for i:=0 to TOBJalSelect.Detail.Count-1 do
    BEGIN
    TOBL:=TOBJalSelect.Detail[i] ;
    TOBL.PutValue('MARQUE','X') ;
    END ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 31/08/2001
Modifié le ... :   /  /
Description .. : Récupération des journaux sélectionnés de la liste.
Mots clefs ... :
*****************************************************************}
procedure TOF_CONVTYPOJAL.RecupMultiSel ;
var i :integer ;
BEGIN
PrechargeOrigines ;
if Not ListeJal.AllSelected then
   BEGIN
   for i:=0 to ListeJal.NbSelected-1 do
     BEGIN
     ListeJal.GotoLeBookmark(i) ;
     MarqueOrigine ;
     END ;
   VireInutiles ;
   END
else ToutMarquer ;
End ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 31/08/2001
Modifié le ... :   /  /
Description .. : Mise à jour des zones écritures.
Mots clefs ... : ECRITURE, MODESAISIE
*****************************************************************}
Procedure TOF_CONVTYPOJAL.MajEcrCompta( TOBdet : TOB ; NP, NL, NG : Integer ) ;
BEGIN
Tobdet.PutValue('E_NUMEROPIECE',NP) ;
Tobdet.PutValue('E_NUMLIGNE',NL) ;
Tobdet.PutValue('E_NUMGROUPEECR',NG) ;
Tobdet.PutValue('E_MODESAISIE',GetControlText('NEWMODE')) ;
TobDet.PutValue('E_IO', 'X') ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 17/09/2001
Modifié le ... :   /  /
Description .. : Mise à jour des écritures analytiques
Mots clefs ... :
*****************************************************************}
Procedure TOF_CONVTYPOJAL.MajEcrAna(var TOBdet : TOB; NP, NL : Integer) ;
var CptAna : integer ;
BEGIN
For CptAna := 0 to TobDet.Detail.Count-1 do
  BEGIN
  TobDet.Detail[CptAna].PutValue('Y_NUMEROPIECE', NP) ;
  TobDet.Detail[CptAna].PutValue('Y_NUMLIGNE', NL) ;
  END ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 31/08/2001
Modifié le ... :   /  /
Description .. : Mise à niveau du solde Euro de la piece sur la dernière
Suite ........ : écriture.
Mots clefs ... : ECRITURE, SOLDEEURO
*****************************************************************}
Procedure TOF_CONVTYPOJAL.MajSoldeEcr (var TobDet : TOB ; SE, SD :Double );
Var NomColEuro, NomColDev : string ;
    TobAna : Tob ;
    CptAna : integer ;
BEGIN
  If Tobdet.getvalue('E_DEBIT') <> 0 then
    BEGIN
    NomColEuro := 'Y_DEBITEURO' ;
    NomColDev := 'Y_DEBITDEV' ;
    SE := -SE ;  SD := -SD ;
    END
  ELSE
    BEGIN
    NomColEuro := 'Y_CREDITEURO' ;
    NomColDev := 'Y_CREDITDEV' ;
    END ;
  Tobdet.PutValue(NomColEuro,Tobdet.getvalue(NomColEuro) + SE ) ;
  Tobdet.PutValue(NomColDev, Tobdet.getvalue(NomColDev) + SD ) ;
  If Tobdet.GetValue('E_ANA') ='X' then
    BEGIN
    If (TobDet.Detail<>NIL) And (TobDet.Detail.Count>0) Then
      For CptAna:=0 To TobDet.Detail.Count-1 Do
        BEGIN
        TobAna:=TobDet.Detail[CptAna] ;
        If TobAna.GetValue('Y_NUMVENTIL')=1 Then
          BEGIN
          TobAna.PutValue(NomColEuro, TobAna.GetValue(NomColEuro) + SE) ;
          TobAna.PutValue(NomColDev, TobAna.GetValue(NomColDev) + SD) ;
          END ;
        END ;
    END ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 27/09/2001
Modifié le ... :   /  /
Description .. : Test si la ligne suivante est une ligne sur le compte écart
Suite ........ : euro.
Mots clefs ... :
*****************************************************************}
Function TOF_CONVTYPOJAL.LigSuivEcartEuro (TT : TOB ; Cpt : integer ) : Boolean ;
BEGIN
  Result := False ;
  If Cpt < TT.Detail.Count-1 then
    BEGIN
    Result := (TT.detail[Cpt+1].GetValue('E_GENERAL')=VH^.EccEuroCredit) or (TT.detail[Cpt+1].GetValue('E_GENERAL')=VH^.EccEuroDebit) ;
    END ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 10/09/2001
Modifié le ... : 10/09/2001
Description .. : Conversion  De (pièce ver BOR ou LIB) Ou (de BOR à LIB)
Mots clefs ... :
*****************************************************************}
Procedure TOF_CONVTYPOJAL.ConvVersLibreOuBor (var TobEcr : Tob);
Var TobDet : TOB ;
    Numpiece, CptEcr : integer ;
    NumGroupeEcr, NumPieceLu : integer ;
BEGIN
Numpiece := StrToInt(GetControlText('Folio')) ;
NumGroupeEcr := 0 ; NumPieceLu := 0 ;
For CptEcr := 0 To TobEcr.Detail.Count -1 do
  BEGIN
  // Pour passage en bordereaux, renumérotation des groupes d'écritures
  If (GetControlText('NEWMODE')= 'BOR') and (NumPieceLu<>TobEcr.Detail[CptEcr].GetValue('E_NumeroPiece')) then
    begin
    Inc(NumGroupeEcr) ;
    NumPieceLu := TobEcr.Detail[CptEcr].GetValue('E_NumeroPiece') ;
    end ;
  TobDet := TobEcr.Detail[CptEcr] ;
  if TobDet.GetValue('E_ANA') = 'X' then MajEcrAna(TOBDet,NumPiece,CptEcr+1) ;
  MajEcrCompta (TOBEcr.Detail[CptEcr], Numpiece, CptEcr+1, NumGroupeEcr) ;
  END ;
TobEcr.UpdateDB ;
Inc(Numpiece) ;
SetControlText('Folio',IntToStr(Numpiece)) ;
MaxFolio := NumPiece ;
SetControlProperty('FOLIO','MinValue',MaxFolio) ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 10/09/2001
Modifié le ... : 11/09/2001
Description .. : Conversion de LIBre à BORdereau ou LIBre à Pièce '-'
Mots clefs ... :
*****************************************************************}
procedure TOF_CONVTYPOJAL.ConvLibToBor (var TobEcr : Tob; CodJal : String ; LibToBor : Boolean) ;
Var TOBDet : TOB ;
    CptEcr, NumPiece, NumLig, DecDev : integer ;
    Solde, SoldeDev, SoldeEuro : Double ;
    NouvellePiece : boolean ;
    NumGroupeEcr : integer ;

  Function LigSuivDateEgal (TT : TOB ; Cpt : integer ) : Boolean ;
  BEGIN
    // Renvoie vrai si la ligne suivante est à la même date
    Result := False ;
    If Cpt < TT.Detail.Count-1 then
      BEGIN
      Result := (TT.detail[Cpt].GetValue('E_DATECOMPTABLE'))=(TT.detail[Cpt+1].GetValue('E_DATECOMPTABLE')) ;
      END ;
  END ;

  Function NatureSuivEgal (TT : TOB ; Cpt :Integer ) :Boolean ;
  BEGIN
    // Renvoie vrai si la ligne suivante est de même nature
    Result := False ;
    If Cpt < TT.Detail.Count-1 then
      BEGIN
      Result := (TT.Detail[Cpt].GetValue('E_NATUREPIECE'))= (TT.Detail[Cpt+1].GetValue('E_NATUREPIECE')) ;
      END ;
  END ;

BEGIN
NouvellePiece := True ;
Solde :=0 ; SoldeDev :=0 ; SoldeEuro :=0 ;
NumGroupeEcr := 0 ;
// Pour chaque ecrit. de la période pour le journal en cours
For CptEcr := 0 to TobEcr.Detail.Count-1 do
  BEGIN
  TobDet := TobEcr.Detail[CptEcr] ;
  DecDev := V_PGI.OkDecV ;
  If TobDet.GetValue('E_DEVISE') <> V_PGI.DevisePivot then RechercheDecDev( DecDev, TobDet ) ;
  If NouvellePiece then // Nouvelle pièce E_numeropiece
    BEGIN
    NumLig := 1 ;
    If LibToBor then Inc(NumGroupeEcr) ;
    // recherche du n° pièce en fonction du nouveau mode de saisie du journal
    if LibToBor then NumPiece := StrToInt(GetControlText('FOLIO'))
                else Numpiece := GetNewNumJal(Codjal,True,Date);
    NouvellePiece := False ;
    END ;
  if TobDet.GetValue('E_ANA') = 'X' then MajEcrAna(TOBDet,NumPiece,NumLig) ;
  MajEcrCompta (TOBDet, Numpiece, Numlig, NumGroupeEcr ) ;
  Solde := Arrondi(Solde + (TobDet.GetValue('E_DEBIT')-TobDet.GetValue('E_CREDIT')),DecDev) ;
  SoldeEuro := Arrondi(SoldeEuro + (TobDet.GetValue('E_DEBITEURO')-TobDet.GetValue('E_CREDITEURO')),DecDev) ;
  SoldeDev := Arrondi(SoldeDev + (TobDet.GetValue('E_DEBITDEV')-TobDet.GetValue('E_CREDITDEV')),DecDev) ;
  // cas ou le solde = 0 et la ligne suivante n'est pas une ligne d'écart
  //     ou la ligne suivante n'est pas à la même date
  //     ou la ligne suivante n'est pas de même nature
  if ((solde=0) and not (LigSuivEcartEuro(TobEcr, CptEcr)))
     or not (LigSuivDateEgal (TobEcr, CptEcr ))
     or not (NatureSuivEgal(TobEcr,Cptecr)) then
    BEGIN
      // mise à niveau des soldes Euro et Dev pour écart d'arrondi si solde =0
      if (solde=0) and ((SoldeEuro <> 0) or (SoldeDev <> 0)) then MajSoldeEcr (TobDet, SoldeEuro, SoldeDev ) ;
      // si solde <> 0 création d'une ligne sur compte d'attente
      If (solde<>0) then
        BEGIN
        Inc(Numlig) ;
        CreerLigCptAttente(Tobdet, Solde, SoldeEuro, Soldedev, Numlig, Numpiece, NumGroupeEcr ) ;
        END ;
      NouvellePiece := Not LibToBor ;
      If LibToBor then Inc(NumGroupeEcr) ;
      Solde :=0 ; SoldeDev :=0 ; SoldeEuro :=0 ;
    END ; // End du if ((Solde =0)
  Inc(NumLig) ;
  END ; // End du  For Cptecr
TobEcr.UpdateDB ;
if LibToBor then
  begin
  SetControlText('FOLIO', IntToStr(NumPiece+1)) ;
  MaxFolio := NumPiece + 1 ;
  SetControlProperty('FOLIO','MinValue',MaxFolio) ;
  end ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 02/10/2001
Modifié le ... :   /  /    
Description .. : Mise à jour des totaux débit, crédit pour journal ou cpte 
Suite ........ : d'attente
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONVTYPOJAL.MajTotaux ( Quoi, CodJal, Exo : String ; Debit, Credit :Double ) ;
var T : Tob ;
    Prefixe : String ;
begin
  If Quoi = 'Jal' then
    begin // Maj du journal
    T := TobJalSelect.FindFirst(['J_JOURNAL'],[CodJal],false ) ;
    Prefixe := 'J_' ;
    end
  else
    begin // Maj du compte d'attente
    T := TobCptAttente.Detail[0] ;
    Prefixe := 'G_' ;
    end ;
  if Exo = VH^.Precedent.Code then // Exercice précedent
    begin
    T.PutValue(Prefixe + 'TOTDEBP', T.GetValue(Prefixe + 'TOTDEBP')+ Debit ) ;
    T.PutValue(Prefixe + 'TOTCREP', T.GetValue(Prefixe + 'TOTCREP')+ Credit ) ;
    end
  else
    begin
    // maj des totaux sur exercices en cours et suivant
    T.PutValue(Prefixe + 'TOTALDEBIT', T.GetValue(Prefixe + 'TOTALDEBIT')+ Debit ) ;
    T.PutValue(Prefixe + 'TOTALCREDIT', T.GetValue(Prefixe + 'TOTALCREDIT')+ Credit ) ;
    if Exo = VH^.Suivant.Code then // exercice suivant
         begin
         T.PutValue(Prefixe + 'TOTDEBS', T.GetValue(Prefixe + 'TOTDEBS')+ Debit ) ;
         T.PutValue(Prefixe + 'TOTCRES', T.GetValue(Prefixe + 'TOTCRES')+ Credit ) ;
         end
       else // exercice en cours
         begin
         T.PutValue(Prefixe + 'TOTDEBE', T.GetValue(Prefixe + 'TOTDEBE')+ Debit ) ;
         T.PutValue(Prefixe + 'TOTCREE', T.GetValue(Prefixe + 'TOTCREE')+ Credit ) ;
         end ;
    end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. : Créaion Nouvelle écriture avec compte d'attente par défaut
Mots clefs ... :
*****************************************************************}
Procedure TOF_CONVTYPOJAL.CreerLigCptAttente (TOBdet : TOB ; Solde, SoldeEuro, SoldeDev : Double ; NumLig, NumPiece, NumGroupe : Integer ) ;
var  TobLigAttente : tob ;
BEGIN
Tobligattente := Tob.Create('ECRITURE',Nil,-1);
TobLigAttente.Dupliquer(TobDet,True,True) ;
TobLigAttente.PutValue('E_GENERAL', VH^.Cpta[fbGene].Attente) ;
TobLigAttente.PutValue('E_TYPEMVT', 'TTC') ;
TobLigAttente.PutValue('E_ETATLETTRAGE', 'RI') ;
TobLigAttente.PutValue('E_ENCAISSEMENT', 'RIE') ;
TobLigAttente.PutValue('E_CONTROLETVA', 'RIE') ;
TobLigAttente.PutValue('E_EXPORTE', '---') ;
TobLigAttente.PutValue('E_CODEACCEPT', 'NON') ;
TobLigAttente.PutValue('E_REGIMETVA', 'FRA') ;
TobLigAttente.PutValue('E_AUXILIAIRE', '') ;
TobLigAttente.PutValue('E_CREERPAR', 'NEW') ;
TobLigAttente.PutValue('E_LIBELLE', LibCptAtt) ;
TobLigAttente.PutValue('E_NUMEROPIECE', NumPiece) ;
TobLigAttente.PutValue('E_NUMLIGNE', NumLig) ;
TobLigAttente.PutValue('E_NUMGROUPEECR', NumGroupe) ;
TobLigAttente.PutValue('E_IO', 'X') ;
if solde <=0 then
  BEGIN
  TobLigAttente.PutValue('E_DEBIT', Abs(Solde)) ;
  TobLigAttente.PutValue('E_DEBITDEV', Abs(SoldeDev)) ;
  TobLigAttente.PutValue('E_DEBITEURO', Abs(SoldeEuro)) ;
  TobLigAttente.PutValue('E_CREDIT', 0) ;
  TobLigAttente.PutValue('E_CREDITDEV', 0) ;
  TobLigAttente.PutValue('E_CREDITEURO', 0) ;
  End Else
  BEGIN
  TobLigAttente.PutValue('E_CREDIT', Abs(Solde)) ;
  TobLigAttente.PutValue('E_CREDITDEV', Abs(SoldeDev)) ;
  TobLigAttente.PutValue('E_CREDITEURO', Abs(SoldeEuro)) ;
  TobLigAttente.PutValue('E_DEBIT', 0) ;
  TobLigAttente.PutValue('E_DEBITDEV', 0) ;
  TobLigAttente.PutValue('E_DEBITEURO', 0) ;
  End ;
TobLigAttente.insertDB(Nil) ;
MajTotaux ('Jal', TobLigAttente.GetValue('E_JOURNAL'), TobLigAttente.GetValue('E_EXERCICE'), TobLigAttente.GetValue('E_DEBIT'), TobLigAttente.GetValue('E_CREDIT')) ;
MajTotaux ('CptAtt', TobLigAttente.GetValue('E_JOURNAL'), TobLigAttente.GetValue('E_EXERCICE'), TobLigAttente.GetValue('E_DEBIT'), TobLigAttente.GetValue('E_CREDIT')) ;
TobLigAttente.Free ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Conversion de BORdereau à Pièce
Mots clefs ... :
*****************************************************************}
Procedure TOF_CONVTYPOJAL.ConvBORToPIE (var TobEcr : Tob ;CodJal : string ) ;
Var TOBDet : TOB ;
    Numlig , Numpiece, CptEcr : integer ;
    Solde, Soldeeuro, soldedev : Double ;
    NouvellePiece : Boolean ;
    DecDev : integer ;
BEGIN
Solde := 0 ; SoldeEuro:= 0 ; SoldeDev := 0 ;
NouvellePiece := True ;
// Pour chaque écriture de la période pour le journal en cours
For CptEcr := 0 to TobEcr.Detail.Count-1 do
  BEGIN
  TobDet := TobEcr.Detail[CptEcr] ;
  DecDev := V_PGI.OkDecV ;
  If TobDet.GetValue('E_DEVISE') <> V_PGI.DevisePivot then RechercheDecDev( DecDev, TobDet ) ;
  if NouvellePiece then // Nouvelle pièce E_numeropiece
    BEGIN
    // recherche du n° pièce en fonction du journal
    NumPiece := GetNewNumJal(CodJal,True,Date) ;
    NumLig := 1 ;
    NouvellePiece := False ;
    END ;
  if TobDet.GetValue('E_ANA') = 'X' then MajEcrAna(TOBDet,NumPiece,NumLig) ;
  MajEcrCompta (TOBDet, Numpiece, Numlig, 0) ;
  Solde := Arrondi(Solde + (TobDet.GetValue('E_DEBIT')-TobDet.GetValue('E_CREDIT')),V_PGI.OkDecV) ;
  SoldeEuro := Arrondi(SoldeEuro + (TobDet.GetValue('E_DEBITEURO')-TobDet.GetValue('E_CREDITEURO')), V_PGI.OkDecE) ;
  SoldeDev := Arrondi(SoldeDev + (TobDet.GetValue('E_DEBITDEV')-TobDet.GetValue('E_CREDITDEV')), DecDev) ;
  // si le solde=0 et la ligne suivante n'est pas une ligne d'écart
  if (Solde=0) and not (LigSuivEcartEuro(TOBEcr,CptEcr)) then
    BEGIN
    // mise à niveau des soldes Euro et Dev
    if (SoldeEuro <> 0) or (SoldeDev <> 0) then MajSoldeEcr (TobDet, SoldeEuro, SoldeDev ) ;
    SoldeEuro:=0 ; SoldeDev:=0 ; Solde :=0 ;
    // changement de pièce
    NouvellePiece := True ;
    END ; // End du if Solde
  Inc(NumLig) ;
  END ; // End du  For Cptecr
TobEcr.UpdateDB ;
END ;

procedure TOF_CONVTYPOJAL.RechercherEcriture ( Var TobEcr : Tob; CodJal : String; DateDeb,DateFin : TDateTime ; Libre : Boolean ) ;
var ReqSql : string ;
    Q : TQuery ;
    TobAna : Tob ;
    IdxEcr : integer ;
begin
// Chargement des écritures comptables du journal sur la période
ReqSql := 'SELECT E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE, E_DEVISE,' ;
ReqSql := ReqSql + 'E_NUMECHE, E_QUALIFPIECE, E_GENERAL,E_AUXILIAIRE,E_ANA,E_DEBIT,E_CREDIT,E_DEBITEURO,' ;
ReqSql := ReqSql + 'E_CREDITEURO,E_DEBITDEV,E_CREDITDEV,E_MODESAISIE, E_NUMGROUPEECR, E_IO' ;
if Libre then
  begin
  // selection de champs supplémentaire pour le mode Libre vers Bor ou pièce pour création ligne attente
  ReqSql := ReqSql + ', E_CODEACCEPT, E_CONFIDENTIEL, E_CONTREPARTIEAUX, E_CONTREPARTIEGEN, E_CONTROLETVA, E_COTATION' ;
  ReqSql := ReqSql + ', E_CREERPAR, E_DATETAUXDEV, E_ECRANOUVEAU, E_ENCAISSEMENT, E_ETABLISSEMENT, E_ETAT, E_ETATLETTRAGE' ;
  ReqSql := ReqSql + ', E_EXPORTE, E_LIBELLE, E_NATUREPIECE, E_PERIODE, E_QUALIFQTE1, E_QUALIFQTE2' ;
  ReqSql := ReqSql + ', E_REFINTERNE, E_REGIMETVA, E_SEMAINE, E_SOCIETE, E_TAUXDEV, E_TYPEMVT'
  end ;
ReqSql := ReqSql + ' FROM ECRITURE WHERE E_JOURNAL="'+CodJal+'" and E_DATECOMPTABLE >="'+USDatetime(DateDeb)+'"' ;
ReqSql := ReqSql + ' and E_DATECOMPTABLE <="'+USDateTime(DateFin)+'" AND E_MODESAISIE<>"'+GetControlText('NEWMODE')+'"'  ;
ReqSql := ReqSql + ' ORDER BY E_JOURNAL,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_QUALIFPIECE' ;
Q := OpenSql(ReqSql,True) ;
TobEcr.LoadDetailDB('ECRITURE', '', '', Q, True, True) ;
Ferme(Q) ;

If TobEcr.Detail.Count > 0 Then
  begin
  // Chargement des écritures analytiques du journal sur la période
  ReqSql := 'SELECT Y_JOURNAL,Y_EXERCICE,Y_DATECOMPTABLE,Y_NUMEROPIECE,Y_NUMLIGNE,Y_AXE,Y_NUMVENTIL,' ;
  ReqSql := ReqSql + 'Y_QUALIFPIECE,Y_GENERAL' ;
  ReqSql := ReqSql + ' FROM ANALYTIQ WHERE Y_JOURNAL="'+CodJal+'" AND Y_DATECOMPTABLE>="'+USDateTime(DateDeb)+'"' ;
  ReqSql := ReqSql + ' AND Y_DATECOMPTABLE<="'+USDateTime(DateFin)+'"' ;
  ReqSql := ReqSql + ' ORDER BY Y_JOURNAL,Y_DATECOMPTABLE,Y_NUMEROPIECE,Y_NUMLIGNE,Y_QUALIFPIECE' ;
  Q:=OpenSql(ReqSql,True) ;
  If not Q.Eof then
    begin
    TobAna:=TOB.Create('_ANALYTIQ', Nil, -1) ;
    TobAna.LoadDetailDB('ANALYTIQ', '', '', Q, True, True) ;
    // mise à jour de la TobEcr avec la TobAna
    For IdxEcr := 0 to TobEcr.Detail.Count -1 do
      begin
      While (TobAna.Detail.Count > 0) AND (TobEcr.Detail[IdxEcr].GetValue('E_NUMEROPIECE')=TobAna.Detail[0].GetValue('Y_NUMEROPIECE'))
         And (TobEcr.Detail[IdxEcr].GetValue('E_NUMLIGNE')=TobAna.Detail[0].GetValue('Y_NUMLIGNE'))
         And (TobEcr.Detail[IdxEcr].GetValue('E_GENERAL')=TobAna.Detail[0].GetValue('Y_GENERAL'))
         And (TobEcr.Detail[IdxEcr].GetValue('E_AUXILIAIRE')=TobAna.Detail[0].GetValue('Y_AUXILIAIRE')) Do
         begin
           TobAna.Detail[0].ChangeParent(TobEcr.Detail[IdxEcr], -1) ;
         end ;
      end ;
    TobAna.free ;
    end ;
  Ferme(Q) ;
  end ;
end ;

procedure TOF_CONVTYPOJAL.ConversionJournal ;
Var CptJal : integer ;
    Q : TQuery ;
    ReqSql : String ;
    Codjal : string ;
    DateDeb, Datefin, Datemax : Tdatetime ;
    TobEcr : Tob ;
BEGIN
EnableControls(Ecran, false ) ;
// récupération des journaux sélectionnés dans une TOB
TOBJalSelect := Tob.create('_JOURNAL',NIL,-1) ;
RecupMultisel ;

TobEcr := Tob.Create ('_ECRITURE', Nil, -1 ) ;

If GetControlText('J_MODESAISIE') = 'LIB' then
  begin
  // récupération du libellé du compte d'attente
  ReqSql := 'SELECT G_GENERAL, G_LIBELLE,G_TOTALCREDIT,G_TOTALDEBIT,G_TOTCREE,G_TOTCREP,' ;
  ReqSql := ReqSql + 'G_TOTCRES,G_TOTDEBE,G_TOTDEBP,G_TOTDEBS FROM GENERAUX' ;
  ReqSql := ReqSql + ' WHERE G_GENERAL="' + VH^.Cpta[fbGene].Attente+'"' ;
  Q := OpenSql(ReqSql,true) ;
  LibCptAtt := Q.fields[1].AsString ;
  TobCptAttente := Tob.Create ('_GENERAUX', Nil, -1 ) ;
  TobCptAttente.LoadDetailDB('GENERAUX', '', '', Q, False, True) ;
  Ferme(Q) ;
  end ;

// pour chaque journal sélectionné
For CptJal := 0 to TobJalSelect.Detail.Count-1 do
  BEGIN
  CodJal := TOBJalSelect.Detail[CptJal].GetValue('J_JOURNAL') ;
  // découpage des écritures à traiter en périodes mensuelles
  ReqSql :='SELECT MIN(E_DATECOMPTABLE),MAX(E_DATECOMPTABLE) FROM ECRITURE WHERE E_JOURNAL="'+CodJal+'"' ;
  Q:=OpenSql(ReqSql,True) ;
  DateMax := Q.Fields[1].AsDateTime ;
  DateDeb := DebutdeMois(Q.Fields[0].AsDateTime) ;
  DateFin := FindeMois(DateDeb) ;
  Ferme(Q) ;
  // BOUCLE pour chaque periode (mois)
  While DateDeb <= DateMax do
    begin
      RechercherEcriture ( TobEcr, CodJal, DateDeb, DateFin, GetControlText('J_MODESAISIE')='LIB' ) ;
      If TobEcr.Detail.Count > 0 Then
        begin
        Try
          BeginTrans ;
          // Passage de pièce (-)  vers (BOR ou LIB) ou de BOR vers LIB
          If (GetControlText('J_MODESAISIE')='-') or
             ((GetControlText('J_MODESAISIE')='BOR') and (GetControlText('NEWMODE')='LIB')) then ConvVersLibreOuBor(TobEcr)
          else
            begin
            //Passage de BORdereau à Pièce '-'
            if (GetControlText('J_MODESAISIE')='BOR') then ConvBorToPie (TobEcr, CodJal)
            else
              // Passage de LIB à BOR
              if (GetControlText('J_MODESAISIE')='LIB') and (GetControlText('NEWMODE') = 'BOR') then ConvLibToBor (TobEcr, CodJal, True)
               // Passage de LIB à Pièce '-'
              else ConvLibToBor (TobEcr, CodJal, False) ;
            end ;
          CommitTrans ;
        Except
          Rollback ;
          PgiBox(TraduireMemoire('ATTENTION : Programme Interrompu. Relancer le traitement.'),TraduireMemoire(Titre)) ;
          Exit ;
          End ;
        end ;
      // Calcul période suivante
      DateDeb := PlusMois(DateDeb,1) ;
      DateFin := FindeMois(DateDeb) ;
      TobEcr.ClearDetail ;
    END ;
  // mise à jour du journal avec le nouveau mode de saisie
  TobJalSelect.Detail[CptJal].PutValue('J_MODESAISIE',GetControlText('NEWMODE')) ;
  END ; // End du for cptjal

// mise à jour du compte d'attente si la tob existe
If TobCptAttente <> Nil Then
  begin
  TobCptAttente.UpdateDB ;
  TobCptAttente.Free ;
  end ;

// mise à jour des journaux dans la base
TobJalSelect.UpdateDB ;

TobEcr.Free ;
TobJalSelect.Free ;
EnableControls(Ecran, true ) ;

END ;

procedure TOF_CONVTYPOJAL.FListeOnDblClick (sender : TObject) ;
begin
  // visu du journal sur double clic
  FicheJournal(QListeJal,'',QListeJal.FindField('J_JOURNAL').asString,taConsult,0) ;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Verifications avant lancement
Mots clefs ... :
*****************************************************************}
procedure TOF_CONVTYPOJAL.BOuvrirOnClick (Sender : TObject ) ;
var Retour : string ;
begin
// lancement du traitement si au moins un journal sélectionné
if (ListeJal.NbSelected=0) and (Not ListeJal.AllSelected) then
  BEGIN
  PgiInfo (TraduireMemoire('Aucune ligne sélectionnée. Traitement impossible.') , TraduireMemoire(Titre) ) ;
  exit ;
  END ;
// Lancement du traitement si le nouveau mode est saisi
if GetControlText('NEWMODE')='' then
  BEGIN
  PgiInfo (TraduireMemoire('Aucun nouveau mode de saisie sélectionné. Traitement impossible.') , TraduireMemoire(Titre) ) ;
  SetFocusControl('NEWMODE') ;
  exit ;
  END ;
If PGIAsk(TraduireMemoire('Confirmez-vous le traitement?'),TraduireMemoire(Titre))=mrYes then
  BEGIN
  If GetControlText('J_MODESAISIE') = 'LIB' then
    begin
      // Affichage avertissement en cas de passage de Libre à Bordereau ou pièce, de la création de ligne Attente
      Retour := AGLLanceFiche('CP','CPAVERT','','','' ) ;
      If Retour <> '1' then exit ;
    end ;
  Conversionjournal ;
  PgiBox (TraduireMemoire('Traitement terminé.'),TraduireMemoire(Titre)) ;
  END ;
if Bcherche <> nil then BCherche.Click ;

END ;


Initialization
  registerclasses ( [ TOF_CONVTYPOJAL ] ) ;
end.
