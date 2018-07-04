 {***********UNITE*************************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 02/03/2001
Modifié le ... : 12/03/2001
Description .. : Source TOF de la TABLE : ICCINTERET ()
Mots clefs ... : TOF;ICCINTERET
*****************************************************************}
Unit uTOFIccInteret ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     MaineAGL,             // AGLLanceFICHE
     UtileAGL,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,              // AGLLanceFICHE
  {$IFDEF V530}
      EdtEtat,
  {$ELSE}
      EdtREtat,
  {$ENDIF}
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTof,
     UTob,
     IccGlobale,
     Graphics,
     Grids,
     Htb97,
     Windows, // TRect
     Ent1    // VH^.
     ;  // V_PGI_ENV

procedure CPLanceFiche_ICCINTERET(psz : String);

Type
  TOF_ICCINTERET = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure DessineCell(ACol, ARow: Integer;Canvas : TCanvas ; AState: TGridDrawState);
    procedure OnClickListBox(Sender : TObject);
    procedure PrepareCalcul(NumCompte : String);
    procedure OnClickBImprimer(Sender: TObject);
    procedure PostDrawCell(ACol, ARow: LongInt; Canvas: TCanvas; AState: TGridDrawState);

  private
    C : TICCCalcul;
    EtatOk : Boolean;
    Onglet : TPageControl;
    ZListBox : TListBox;
    GridEcriture : THGrid;
    TobMere,TobEtat : TOB;

    SWhereEtat : String;
    FGroupe, FLimitation, FDirigeant : Boolean;

    procedure ParametrageGrille ;
  end;

Const LargeurColonne  = 100;  // Largeur des colonnes pour les montants avec DB ou CR
      LargeurColonne2 = 85;

      ColDate       = 0;
      ColSolde      = 1;
      ColCapital    = 2;
      ColTaux       = 3;
      ColMax        = 4;
      ColDuree      = 5;
      ColInteret    = 6;
      ColSurCapital = 7;
      ColSurTaux    = 8;

Implementation

uses
  uLibWindows;

procedure CPLanceFiche_ICCINTERET(psz : String);
begin
  AGLLanceFiche('CP','ICCINTERET','','', psz);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/09/2001
Modifié le ... : 12/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCINTERET.OnArgument (S : String ) ;
var stArg, STemp : String;
begin
  Inherited;
  { Récupération dans ICC_DATA des paramètres de calcul saisis }
  stArg := S;

  ICC_Data.NombreJours := StrToInt(ReadTokenSt(StArg));       // Nombre de jours
  ICC_Data.Methode     := StrToInt(ReadTokenSt(StArg));       // Methode de calcul utilisée
  ICC_Data.DateDu      := StrToDate(ReadTokenSt(StArg));      // Date du des paramètres de calcul
  ICC_Data.DateAu      := StrToDate(ReadTokenSt(StArg));      // Date au des paramètres de calcul
  ICC_Data.TauxLegal   := Arrondi(CalculTauxMaxLegal(Icc_Data.DateDu,Icc_Data.DateAu),V_PGI.OkDecV);

  if Icc_Data.NombreJours = 365 then
  begin
    if IsValidDate('29/02/'+ FormatDateTime('YYYY',ICC_Data.DateDu)) then
    begin
      if (StrToDate('29/02/'+ FormatDateTime('YYYY',ICC_Data.DateDu))>= ICC_Data.DateDu) and
         (StrToDate('29/02/'+ FormatDateTime('YYYY',ICC_Data.DateDu))<= ICC_Data.DateAu) then
        Icc_Data.NombreJours := 366;
    end;

    if IsValidDate('29/02/'+ FormatDateTime('YYYY',ICC_Data.DateAu)) then
    begin
      if (StrToDate('29/02/'+ FormatDateTime('YYYY',ICC_Data.DateAu))>= ICC_Data.DateDu) and
         (StrToDate('29/02/'+ FormatDateTime('YYYY',ICC_Data.DateAu))<= ICC_Data.DateAu) then
        Icc_Data.NombreJours := 366;
    end;
  end;

  Onglet := TPageControl(GetControl('ZPageControl'));
  Onglet.ActivePage := Onglet.Pages[0];

  THLabel(GetControl('ZNBJOURS')).Caption := IntToStr(Icc_Data.NombreJours);
  THLabel(GetControl('ZDATEDUDATEAU')).Caption := TraduireMemoire('Période du ') +
                                                  DateToStr(ICC_Data.DateDu) +
                                                  TraduireMemoire(' au ') +
                                                  DateToStr(ICC_Data.DateAu);

  SWhereEtat := '';
  STemp := ''; // Ré-utilisation de STEMP
  // Chargement la ListBox avec les comptes sélectionnés
  ZListBox := TListBox(GetControl('ZLISTBOX'));
  ZListBox.Clear;
  while StArg <> '' do
  begin
    STemp := ReadTokenSt(StArg);
    ZListBox.Items.Add(STemp);
    if SWhereEtat = '' then
      SWhereEtat := ' (ICZ_GENERAL="'+ STemp + '")'
    else
      SWhereEtat := SWhereEtat + ' OR' + ' (ICZ_GENERAL="'+ STemp + '")';
  end;
  SWhereEtat := SWhereEtat + ' ';
  //SWhereEtat := SWhereEtat + ' ORDER BY ICT_GENERAL,ICT_DATE';

  ZListBox.OnClick := OnClickListBox;
  TToolBarButton97(GetControl('BIMPRIMER')).OnClick := OnClickBImprimer;

  ParametrageGrille;

  TobMere := Tob.Create('',nil,-1);
  TobEtat := Tob.Create('',nil,-1);
  try
    BeginTrans;
    ExecuteSQL('DELETE FROM ICCEDTTEMP');
    CommitTrans;
  except
    Rollback;
  end;

  if CtxPcl in V_Pgi.PGIContexte then
    THEdit(GetControl('ENODOSSIER')).Text := V_PGI.NoDossier
  else
    THEdit(GetControl('ENODOSSIER')).Text := V_PGI.CodeSociete+' '+V_PGI.NomSociete ;

  THEdit(GetControl('ETAUXABSENT')).Text     := IIF(Icc_Data.TauxAbsent, 'TRUE', 'FALSE');
  THEdit(GetControl('ETAUXPROVISOIRE')).Text := IIF(Icc_Data.TauxProvisoire, 'TRUE', 'FALSE');
  THEdit(GetControl('ERECAPITULATIF')).Text  := IIF(Icc_Data.AvecRecapitulatif, 'TRUE', 'FALSE');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/09/2001
Modifié le ... : 12/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCINTERET.OnClose ;
begin
  Inherited ;
  if TobMere <> nil then
  begin
    TobMere.ClearDetail;
    TobMere.Free;
  end;

  if TobEtat <> nil then
  begin
    TobEtat.ClearDetail;
    TobEtat.Free;
  end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/09/2001
Modifié le ... : 12/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCINTERET.OnNew ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/09/2001
Modifié le ... : 12/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCINTERET.OnDelete ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/09/2001
Modifié le ... : 12/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCINTERET.OnUpdate ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/09/2001
Modifié le ... : 12/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCINTERET.OnLoad ;
var i : integer;
begin
  Inherited ;
  Icc_Data.TotInt := 0;
  Icc_Data.TotIntNonDedCapital := 0;
  Icc_Data.TotIntNonDedTaux := 0;
  Icc_Data.TotIntNonDed := 0;
  Icc_Data.TotIntDed := 0;

  EtatOk := False;
  (* ---- On effectue le calcul pour tous les comptes sélectionnés ---- *)
  for i:=0 to ZListBox.Items.count-1 do
  begin
    PrepareCalcul(ZListBox.Items[i]);
  end;
  EtatOk := True;

  //TobEtat.SaveToFile('C:\Export.txt',false,true,true,'');
  TobEtat.InsertDB(nil,False);

  THLabel(GetControl('ZTOTINTERET')).Caption := StrfMontant(Icc_Data.TotInt,15,V_PGI.OkDecV,'',True);
  THLabel(GetControl('ZTAUX')).Caption := StrfMontant(Icc_Data.TotIntNonDedTaux,15,V_PGI.OkDecV,'',True);
  THLabel(GetControl('ZCAPITAL')).Caption := StrfMontant(Icc_Data.TotIntNonDedCapital,15,V_PGI.OkDecV,'',True);

  Icc_Data.TotIntNonDed := Icc_Data.TotIntNonDedTaux + Icc_Data.TotIntNonDedCapital;
  Icc_Data.TotIntDed    := Icc_Data.TotInt - Icc_Data.TotIntNonDed;

  THLabel(GetControl('ZTOTNONDED')).Caption  := StrfMontant(Icc_Data.TotIntNonDed,15,V_PGI.OkDecV,'',True);
  THLabel(GetControl('ZTOTNONDED2')).Caption := StrfMontant(Icc_Data.TotIntNonDed,15,V_PGI.OkDecV,'',True);
  THLabel(GetControl('ZTOTDED')).Caption := StrfMontant(Icc_Data.TotIntDed,15,V_PGI.OkDecV,'',True);

  // Calcul des intérêts du premier compte de la ListBox
  PrepareCalcul(ZListBox.Items[0]);
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 06/09/2001
Modifié le ... : 06/09/2001
Description .. : Gestion du grisage des colonnes de la grille
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCINTERET.PostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var TheRect : TRect;
begin
  if Arow = 0 then Exit;

  if (((ACol = ColCapital) or (ACol = ColSurCapital)) and (not FDirigeant)) or
     ((ACol = ColSurTaux) and (not FLimitation)) then
  begin
    if Valeur(GridEcriture.Cells[ACol, ARow]) = 0 then
    begin
      TheRect := GridEcriture.CellRect(ACol,ARow);
      Canvas.TextRect(TheRect,TheRect.Left,TheRect.Top,'');
      Canvas.Brush.Color := GridEcriture.FixedColor ;
      Canvas.Brush.Style := bsBDiagonal ;
      Canvas.Pen.Color   := GridEcriture.FixedColor ;
      Canvas.Pen.Mode    := pmCopy ;
      Canvas.Pen.Style   := psClear ;
      Canvas.Pen.Width   := 1 ;
      Canvas.Rectangle(TheRect); // Afficher les rayures
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/09/2001
Modifié le ... : 12/09/2001
Description .. : Prépare et effectue le calcul des intérets du compte
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCINTERET.PrepareCalcul(NumCompte : String);
var Q : TQuery;
    i : Integer;
    Interet,InteretTaux,InteretCapital : Double;
    TobTemp : Tob;
    SLibelle : String;
begin
  TobMere.ClearDetail;

  Interet := 0;
  InteretTaux := 0;
  InteretCapital := 0;

  (* ---- Recherche du Libelle dans la table des GENERAUX ---- *)
  Q := nil;
  try
    Q := OpenSQL('SELECT G_GENERAL, G_LIBELLE FROM GENERAUX WHERE G_GENERAL = "'+ NumCompte +'"',True);

    SLibelle := IIF(Q.EOF,'', Q.FindField('G_LIBELLE').AsString);
    THLabel(GetControl('ZLIBELLE')).Caption := SLibelle;
  finally
    if Assigned(Q) then Ferme(Q);
  end;

  // Recherche des ecritures du compte
  Q := nil;
  try
    Q := OpenSQL('SELECT ICE_GENERAL, ICE_LIGNE, ICE_LIBELLE, ICE_DATEVALEUR, ICE_DEBIT, ICE_CREDIT FROM ICCECRITURE' +
                 ' WHERE (ICE_GENERAL = "' + NumCompte+ '") and (ICE_DATEVALEUR <= "'+USDateTime(Icc_Data.DateAu)+'") ORDER BY ICE_DATEVALEUR',True);

    TobMere.LoadDetailDB('ICCECRITURE', '', '', Q, True, False);
  finally
    if Assigned(Q) then Ferme(Q);
  end;

  (* ---- Recherche du compte pour récupération des infos ---- *)
  Q := nil;
  try
    Q := OpenSQL('SELECT ICG_GENERAL,ICG_TXANNUEL,ICG_GROUPE,ICG_LIMITATION,ICG_DIRIGEANT,ICG_DATEEFFET,ICG_SOLDEDEBEX FROM ICCGENERAUX' +
                 ' WHERE ICG_GENERAL ="' + NumCompte+  '"',True);

    THLabel(GetControl('ZCOMPTE')).Caption := NumCompte;
    FGroupe := (Q.FindField('ICG_GROUPE').AsString = 'X');         // Compte de Groupe
    FLimitation := (Q.FindField('ICG_LIMITATION').AsString = 'X'); // Limitation légale du taux
    FDirigeant := (Q.FindField('ICG_DIRIGEANT').AsString = 'X');   // Dirigeant ou plus de 50 % du capital

    // Envoi des paramètres pour calcul
    C := TICCCalcul.Create;
    C.CalculTOB(TobMere,Q);
    C.Free;
  finally
    if Assigned(Q) then Ferme(Q);
  end;

  // Les calculs sont terminées, on affiche les resultats dans la grille
  GridEcriture.RowCount := TobMere.Detail.Count + 1;
  TobMere.PutGridDetail(GridEcriture,False,False,'ICE_DATEVALEUR; ICE_SOLDE2; ICE_CAPITAL2 ;ICE_TAUX; ICE_TAUXMAX; ICE_NBJOURS; ICE_INTERET; ICE_SURCAPITAL; ICE_SURTAUX');


  (* ---- Cumul des intérêts déductibles et non déductibles ---- *)
  for i:=0 to TobMere.Detail.Count-1 do
  begin

    if not EtatOk then // On est dans le OnLoad de la Fiche
    begin
      TobTemp := Tob.Create('ICCEDTTEMP', TobEtat,-1);
      TobTemp.PutValue('ICZ_UTILISATEUR', V_PGI.User);
      TobTemp.PutValue('ICZ_GENERAL'    , NumCompte);
      TobTemp.PutValue('ICZ_LIGNE'      , i + 1);
      TobTemp.PutValue('ICZ_LIBELLE'    , SLibelle);
      TobTemp.PutValue('ICZ_DATE'       , TobMere.Detail[i].GetValue('ICE_DATEVALEUR'));
      TobTemp.PutValue('ICZ_SOLDE'      , TobMere.Detail[i].GetValue('ICE_SOLDE'));
      TobTemp.PutValue('ICZ_CAPITAL'    , TobMere.Detail[i].GetValue('ICE_CAPITAL'));
      TobTemp.PutValue('ICZ_TAUX'       , TobMere.Detail[i].GetValue('ICE_TAUX'));
      TobTemp.PutValue('ICZ_TAUXMAX'    , TobMere.Detail[i].GetValue('ICE_TAUXMAX'));
      TobTemp.PutValue('ICZ_NBJOURS'    , TobMere.Detail[i].GetValue('ICE_NBJOURS'));
      TobTemp.PutValue('ICZ_INTERET'    , TobMere.Detail[i].GetValue('ICE_INTERET'));
      TobTemp.PutValue('ICZ_SURCAPITAL' , TobMere.Detail[i].GetValue('ICE_SURCAPITAL'));
      TobTemp.PutValue('ICZ_SURTAUX'    , TobMere.Detail[i].GetValue('ICE_SURTAUX'));
    end;

    Interet        := Interet        + Arrondi(TobMere.Detail[i].GetValue('ICE_INTERET'),V_PGI.OkDecV);
    InteretTaux    := InteretTaux    + Arrondi(TobMere.Detail[i].GetValue('ICE_SURTAUX'),V_PGI.OkDecV);
    InteretCapital := InteretCapital + Arrondi(TobMere.Detail[i].GetValue('ICE_SURCAPITAL'),V_PGI.OkDecV);

  end;

  if not EtatOk then
  begin
    Icc_Data.TotInt := Icc_Data.TotInt + Interet;
    Icc_Data.TotIntNonDedTaux := Icc_Data.TotIntNonDedTaux + InteretTaux;
    Icc_Data.TotIntNonDedCapital := Icc_Data.TotIntNonDedCapital + InteretCapital;
  end;

  if VH^.TenueEuro then
    THLabel(GetControl('ZINTERET')).Caption  := StrfMontant(Interet,15,V_PGI.OkDecV,'',True) + ' Euros'
  else
    THLabel(GetControl('ZINTERET')).Caption  := StrfMontant(Interet,15,V_PGI.OkDecV,'',True) + ' Francs';

  if VH^.TenueEuro then
    THLabel(GetControl('ZINTERET2')).Caption := StrfMontant(InteretTaux + InteretCapital,15,V_PGI.OkDecV,'',True) + ' Euros'
  else
    THLabel(GetControl('ZINTERET2')).Caption := StrfMontant(InteretTaux + InteretCapital,15,V_PGI.OkDecV,'',True) + ' Francs';

  THNumEdit(GetControl('ZTOTAL1')).Text := StrfMontant(Interet,15,V_PGI.OkDecV,'',True);
  THNumEdit(GetControl('ZTOTAL2')).Text := StrfMontant(InteretCapital,15,V_PGI.OkDecV,'',True);
  THNumEdit(GetControl('ZTOTAL3')).Text := strfMontant(InteretTaux,15,V_PGI.OkDecV,'',True);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/09/2001
Modifié le ... : 12/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOF_ICCINTERET.DessineCell(ACol, ARow: Integer; Canvas: TCanvas;AState: TGridDrawState);
Begin
  if ARow = 0 then Exit;

  case ACol of

    ColSolde : Canvas.Font.Color := IIF(Pos('C',GridEcriture.Cells[ACol,Arow]) > 0,clGreen, clRed);

    ColInteret :
      begin
        //Canvas.Font.Style := [fsBold];
        Canvas.Font.Color := clNavy;
      end;

    ColSurCapital,

    ColSurTaux :
      begin
        //Canvas.Font.Style := [fsBold];
        Canvas.Font.Color := clNavy;
      end;
  end;
End;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/09/2001
Modifié le ... : 12/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCINTERET.OnClickListBox(Sender: TObject);
begin
  // Calcul des intérêts du compte sélectionné dans la listBox
  PrepareCalcul(ZListBox.Items[ZListBox.ItemIndex]);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/09/2001
Modifié le ... : 12/09/2001
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCINTERET.OnClickBImprimer(Sender: TObject);
begin
  inherited;
  LanceEtat('E', 'ICC', 'ICC', True, False, False, Onglet, ' '+SWhereEtat, '', False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/10/2001
Modifié le ... : 02/10/2001
Description .. : Mise en forme de la grille d' affichage des résultats
Mots clefs ... : THGRID
*****************************************************************}
procedure TOF_ICCINTERET.ParametrageGrille ;
begin
  GridEcriture := THGrid(GetControl('GRIDECRITURE'));

  GridEcriture.GetCellCanvas          := DessineCell;
  GridEcriture.PostDrawCell           := PostDrawCell;

  GridEcriture.Cells[ColDate,0]       := TraduireMemoire('Date');
  GridEcriture.ColAligns[ColDate]     := TaCenter;
  GridEcriture.ColWidths[ColDate]     := 70;

  GridEcriture.Cells[ColSolde,0]      := TraduireMemoire('Solde');
  GridEcriture.ColAligns[ColSolde]    := TaRightJustify;
  GridEcriture.ColWidths[ColSolde]    := LargeurColonne;
  GridEcriture.ColFormats[ColSolde]   := '#,##0.00';

  GridEcriture.Cells[ColCapital,0]    := TraduireMemoire('Capital');
  GridEcriture.ColAligns[ColCapital]  := TaRightJustify;
  GridEcriture.ColWidths[ColCapital]  := LargeurColonne;
  GridEcriture.ColFormats[ColCapital] := '#,##0.00';

  GridEcriture.Cells[ColTaux,0]       := TraduireMemoire('Taux');
  GridEcriture.ColAligns[ColTaux]     := TaRightJustify;
  GridEcriture.ColWidths[ColTaux]     := 45;
  GridEcriture.ColFormats[ColTaux]    := '#,##0.00';

  GridEcriture.Cells[ColMax,0]        := TraduireMemoire('Max');
  GridEcriture.ColAligns[ColMax]      := TaRightJustify;
  GridEcriture.ColWidths[ColMax]      := 45;
  GridEcriture.ColFormats[ColMax]     := '#,##0.00';

  GridEcriture.Cells[ColDuree,0]      := TraduireMemoire('Durée');
  GridEcriture.ColAligns[ColDuree]    := TaRightJustify;
  GridEcriture.ColWidths[ColDuree]    := 45;

  GridEcriture.Cells[ColInteret,0]    := TraduireMemoire('Intérêts');
  GridEcriture.ColAligns[ColInteret]  := TaRightJustify;
  GridEcriture.ColWidths[ColInteret]  := LargeurColonne2;
  GridEcriture.ColFormats[ColInteret] := '#,##0.00';

  // Non déductibles sur Capital
  GridEcriture.ColAligns[ColSurCapital]  := TaRightJustify;
  GridEcriture.ColWidths[ColSurCapital]  := LargeurColonne2;
  GridEcriture.Cells[ColSurCapital,0]    := TraduireMemoire('sur capital');
  GridEcriture.ColFormats[ColSurCapital] := '#,##0.00';

  // Non déductibles sur Taux
  GridEcriture.ColAligns[ColSurTaux]  := TaRightJustify;
  GridEcriture.ColWidths[ColSurTaux]  := LargeurColonne2;
  GridEcriture.Cells[ColSurTaux,0]    := TraduireMemoire('sur taux');
  GridEcriture.ColFormats[ColSurTaux] := '#,##0.00';
end;

Initialization
  registerclasses ( [ TOF_ICCINTERET ] ) ;
end.
