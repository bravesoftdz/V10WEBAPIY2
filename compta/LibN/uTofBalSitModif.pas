{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 24/10/2001
Modifié le ... : 18/06/2007
Description .. : Source TOF de la TABLE : BALSIT_MODIF ()
Suite ........ : - CA - 18/06/2007 - FQ 19814 - Ne pas autoriser la 
Suite ........ : modification d'une balance dynamique
Mots clefs ... : TOF;BALSIT_MODIF
*****************************************************************}
Unit uTofBalSitModif ;

Interface

Uses Windows ,StdCtrls, Controls, Classes, db, forms, sysutils,
     ComCtrls, dialogs, HCtrls, Grids, HEnt1, HMsgBox, UTOF, HSysMenu, HTB97,
     uTOB, Vierge, Ent1, CPGeneraux_TOM, BalSit,
{$IFDEF EAGLCLIENT}
     MainEagl, // AGLLanceFiche
{$ELSE}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,  // AGLLanceFiche
     Printdbg,
{$ENDIF}
     LookUp
{$IFDEF EAGLCLIENT}
     ,uObjEtats
{$ENDIF}
     ;

const
  COL_COMPTE1 = 0;
  COL_COMPTE2 = 1;
  COL_LIBELLE = 2;
  COL_DEBIT   = 3;
  COL_CREDIT  = 4;

Type
  TOF_BALSIT_MODIF = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
      FGrille       : THGrid;
      FCodeBal      : string;
      FListeEcr     : TOB;
      FBalSit       : TBalSit;
      FHSystemMenu  : THSystemMenu;
      FFindDialog   : TFindDialog;
      FFindFirst    : boolean;
      procedure   OnChercherClick ( Sender : TObject );
      procedure   OnNouveauCompteClick ( Sender : TObject );
      procedure   OnInsertionCompteClick ( Sender : TObject );
      procedure   OnSuppressionCompteClick ( Sender : TObject ) ;
      procedure   OnPurgeClick ( Sender : TOBject );
      procedure   OnImprimerClick ( Sender : TOBject );
      procedure   OnFindText(Sender: TObject);
      procedure   InitGrille ( TypeCum : string ; var ColName : string );
      procedure   OnGrilleRowExit (Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
      procedure   OnGrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
      procedure   OnFormKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
      function    EnregistreLigne ( Ligne : integer ) : boolean;
      procedure   PositionneSurLigne ( Compte1, Compte2 : string );
      procedure   AfficheSoldeBalance;
      procedure   RemplirGrille;
      procedure   SoldeBalance;
  end ;

procedure CPLanceFiche_BALSIT_MODIF( vStCodeBalance : string );

Implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  ParamSoc;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure CPLanceFiche_BALSIT_MODIF( vStCodeBalance : string );
begin
  AGLLanceFiche('CP', 'BALSIT_MODIF', '', '', vStCodeBalance);

end;
////////////////////////////////////////////////////////////////////////////////

procedure TOF_BALSIT_MODIF.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BALSIT_MODIF.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BALSIT_MODIF.OnUpdate ;
var okok : boolean;
  T : THNumEdit;
  co, li : integer;
  bCancel : boolean;
begin
  FBalSit.SetLibelle (GetControlText('FLIBELLE'));
  FBalSit.SetAbrege (GetControlText('FABREGE'));
  co := FGrille.Col;
  li := FGrille.Row;
  OnGrilleCellExit(FGrille,co,li,bCancel);
  if not EnregistreLigne ( FGrille.Row ) then
  begin
    LastError := -2;
    LastErrorMsg := 'Enregistrement impossible. Ligne incorrecte.';
  end else
  begin
    okok := True;
    AfficheSoldeBalance;
    T := THNumEdit(GetControl('FSOLDEBALANCE'));
    if (Arrondi(T.Value,V_PGI.OkDecV) <> 0) then
      okok := (PGIAsk('Balance non équilibrée !#10#13Voulez-vous continuer ?',ECRAN.Caption)=mrYes);
    if okok then
      if Transactions (FBalSit.Enregistre,3) <> oeOK then
      begin
        LastError := -1;
        LastErrorMsg := 'Enregistrement de la balance impossible';
      end else PGIInfo ('Enregistrement de la balance terminé',ECRAN.Caption);
  end;
  Inherited ;
end ;

procedure TOF_BALSIT_MODIF.OnLoad ;
begin
  Inherited ;
  Ecran.Caption := TraduireMemoire ('Balance - ')+
                   RechDom('CPBALSITTYPE',FBalSit.TypeCumul,False)+ ' - '+
                   TraduireMemoire('du ')+DateToStr(FBalsit.DateDebut)+
                   TraduireMemoire(' au ')+
                   DateToStr(FBalsit.DateFin);
end;

procedure TOF_BALSIT_MODIF.OnArgument (S : String ) ;
var   ColName : string;
begin
  Inherited ;

  // Initialiation des propriétés
  FGrille             := THGrid(GetControl('FGRILLE'));
  FGrille.OnRowExit   := OnGrilleRowExit;
  FGrille.OnCellExit  := OnGrilleCellExit;
//  FGrille.OnKeyDown   := OnGrilleKeyDown;
  ECRAN.OnKeyDown := OnFormKeyDown;

  TToolBarButton97(GetControl('BCHERCHER')).OnClick := OnChercherClick ;
  TToolBarButton97(GetControl('BINSERT')).OnClick := OnNouveauCompteClick ;
  TToolBarButton97(GetControl('BINSERTIONCOMPTE')).OnClick := OnInsertionCompteClick ;
  TToolBarButton97(GetControl('BSUPPRIMECOMPTE')).OnClick := OnSuppressionCompteClick ;
  TToolBarButton97(GetControl('BPURGE')).OnClick := OnPurgeClick ;
  TToolBarButton97(GetControl('BIMPRIMER')).OnClick := OnImprimerClick ;

  FHSystemMenu := THSystemMenu (TFVierge(ECRAN).HMTrad);

  FFindDialog := TFindDialog.Create ( ECRAN );
  FFindDialog.OnFind := OnFindText;

  FCodeBal := S;
  FBalSit := TBalSit.Create (FCodeBal);
  FBalSit.Charge ( True );

  // Affichage du titre de la balance
  if not (FBalSit.TypeCumul='GEN') then SetControlVisible('BINSERT',False);
  SetControlText('FCODEBAL',FCodeBal);
  SetControlText('FLIBELLE',FBalSit.Libelle);
  SetControlText('FABREGE',FBalSit.Abrege);
  SetControlProperty ( 'FCUMULDEBIT', 'Visible',  FBalSit.TypeCumul='GEN' );
  SetControlProperty ( 'FCUMULCREDIT', 'Visible', FBalSit.TypeCumul='GEN' );
  SetControlProperty ( 'FLCUMULCLASSE', 'Visible', FBalSit.TypeCumul='GEN');
  SetControlProperty ( 'FRESULTAT', 'Visible',  FBalSit.TypeCumul='GEN');
  SetControlProperty ( 'FLRESULTAT', 'Visible', FBalSit.TypeCumul='GEN');
  // Initialisation de la grille
  InitGrille ( FBalSit.TypeCumul, ColName );

  // Remplissage de la grille
  RemplirGrille;

  // Redimensionnement des colonnes
  TFVierge(Ecran).FormResize := True;
  FHSystemMenu.ResizeGridColumns ( FGrille );

  AfficheSoldeBalance;

  FGrille.SetFocus;
  FGrille.Col := COL_DEBIT;
  FGrille.Row := 1;

  // GCO - 06/07/2006
  if FBalSit.TypeBalance = 'DYN' then
  begin
    SetControlEnabled('BINSERT', False);
    SetControlEnabled('BPURGE', False);
    SetControlEnabled('BINSERTIONCOMPTE', False);
    SetControlEnabled('BSUPPRIMECOMPTE', False);
  end;
end ;

procedure TOF_BALSIT_MODIF.OnClose ;
begin
  FFindDialog.Free;
  FListeEcr.Free;
  FBalSit.Free;
  Inherited ;
end ;

procedure TOF_BALSIT_MODIF.InitGrille(TypeCum: string ; var ColName : string);
var   LibCompte1  : string;
      LibCompte2  : string;
begin
  LibCompte2 := '';
  if TypeCum = 'GEN' then
  begin
    ColName := 'BSE_COMPTE1;;G_LIBELLE;BSE_DEBIT;BSE_CREDIT';
    LibCompte1 := TraduireMemoire('Général');
  end else if TypeCum = 'TIE' then
  begin
    ColName := 'BSE_COMPTE1;;T_LIBELLE;BSE_DEBIT;BSE_CREDIT';
    LibCompte1 := TraduireMemoire('Auxiliaire');
  end else if TypeCum = 'ANA' then
  begin
    ColName := 'BSE_COMPTE1;;S_LIBELLE;BSE_DEBIT;BSE_CREDIT';
    LibCompte1 := TraduireMemoire('Section');
  end else if TypeCum = 'G/T' then
  begin
    ColName := 'BSE_COMPTE1;BSE_COMPTE2;T_LIBELLE;BSE_DEBIT;BSE_CREDIT';
    LibCompte1 := TraduireMemoire('Général');
    LibCompte2 := TraduireMemoire('Auxiliaire');
  end else if TypeCum = 'G/A' then
  begin
    ColName := 'BSE_COMPTE1;BSE_COMPTE2;S_LIBELLE;BSE_DEBIT;BSE_CREDIT';
    LibCompte1 := TraduireMemoire('Général');
    LibCompte2 := TraduireMemoire('Section');
  end;

  // Libellés des colonnes
  FGrille.Titres[COL_COMPTE1] := LibCompte1;
  FGrille.Titres[COL_COMPTE2] := LibCompte2;
  FGrille.Cells[COL_COMPTE1,0] := LibCompte1;
  FGrille.Cells[COL_COMPTE2,0] := LibCompte2;

  // Largeur des colonnes
  FGrille.ColWidths[COL_COMPTE1]  := 70;
  FGrille.ColLengths[COL_COMPTE1] := -1;
  FGrille.ColLengths[COL_COMPTE2] := -1;
  FGrille.ColLengths[COL_LIBELLE] := -1;
  if LibCompte2<>'' then FGrille.ColWidths[COL_COMPTE2]  := 70
  else FGrille.ColWidths[COL_COMPTE2]  := 0;
  FGrille.ColWidths[COL_LIBELLE]  := 140;
  FGrille.ColWidths[COL_DEBIT]    := 70;
  FGrille.ColWidths[COL_CREDIT]   := 70;

  // Positionnement du texte dans les colonnes
  FGrille.ColAligns[COL_COMPTE1]    :=  taCenter ;
  FGrille.ColAligns[COL_COMPTE2]    :=  taCenter ;
  FGrille.ColAligns[COL_LIBELLE]    :=  taLeftJustify ;
  FGrille.ColAligns[COL_DEBIT]      :=  taRightJustify ;
  FGrille.ColAligns[COL_CREDIT]     :=  taRightJustify ;

  // Formatage des colonnes
  FGrille.ColFormats[COL_DEBIT] := '#,##0.00';
  FGrille.ColFormats[COL_CREDIT] := '#,##0.00';
  FGrille.ColTypes[COL_DEBIT]:='R';
  FGrille.ColTypes[COL_CREDIT]:='R';

  // GCO - 30/06/2006
  if FBalSit.TypeBalance = 'DYN' then
    FGrille.Options := FGrille.Options - [goEditing] + [goRowSelect] - [goTabs]
  else
    FGrille.Options := FGrille.Options + [goEditing] + [goAlwaysShowEditor] + [goTabs] - [goRowSelect];
end;

procedure TOF_BALSIT_MODIF.AfficheSoldeBalance;
var i : integer;
    TotalD, TotalC : double;
    SoldeBalD, SoldeBalC : double;
    ClasseD, ClasseC, Resultat : double;
    Classe : string;
    AvecClasse : boolean;
begin
  TotalD := 0; TotalC := 0;
  SoldeBalD := 0; SoldeBalC := 0;
  ClasseD :=0; ClasseC := 0; Resultat := 0;
  AvecClasse := Copy(FBalSit.TypeCumul,1,1)='G';
  Classe:=Copy( FGrille.Cells[COL_COMPTE1, FGrille.Row], 1, 1 );
  for i:=1 to FGrille.RowCount - 1 do
  begin
    if i <= FGrille.Row then
    begin
      TotalD := TotalD + Valeur( FGrille.Cells[COL_DEBIT,i] );
      TotalC := TotalC + Valeur( FGrille.Cells[COL_CREDIT,i] );
      if (AvecClasse) and (Copy(FGrille.Cells[COL_COMPTE1, i], 1, 1)=Classe) then
      begin
        ClasseD := ClasseD + Valeur(FGrille.Cells[COL_DEBIT,  i]) ;
        ClasseC := ClasseC + Valeur(FGrille.Cells[COL_CREDIT, i]) ;
      end;
    end;
    // Calcul du solde de la balance
    SoldeBalD:=SoldeBalD + Valeur( FGrille.Cells[COL_DEBIT,   i]) ;
    SoldeBalC:=SoldeBalC + Valeur( FGrille.Cells[COL_CREDIT,  i]) ;
    // Calcul du résultat
    if AvecClasse then
    begin
      if (FGrille.Cells[COL_COMPTE1,i] = GetParamSocSecur('SO_OUVREPERTE','')) or
          (FGrille.Cells[COL_COMPTE1,i] = GetParamSocSecur('SO_OUVREBEN','')) or
          (Copy(FGrille.Cells[COL_COMPTE1, i], 1, 1)='6') or
          (Copy(FGrille.Cells[COL_COMPTE1, i], 1, 1)='7') then
        Resultat := Resultat - Valeur(FGrille.Cells[COL_DEBIT, i])
                         + Valeur(FGrille.Cells[COL_CREDIT, i]) ;
    end;
  end;
  SetControlProperty ( 'FTOTALDEBIT', 'Value', TotalD );
  SetControlProperty ( 'FTOTALCREDIT', 'Value', TotalC );
  SetControlProperty ( 'FCUMULDEBIT', 'Value', ClasseD );
  SetControlProperty ( 'FCUMULCREDIT', 'Value', ClasseC );
  SetControlProperty ( 'FRESULTAT', 'Value', Resultat );
  if AvecClasse then SetControlProperty ( 'FLCUMULCLASSE','Caption', TraduireMemoire('Cumul classe ')+Classe);
  AfficheLeSolde(THNumEdit(GetControl('FSOLDEPROGRESSIF')), TotalD, TotalC) ;
  AfficheLeSolde(THNumEdit(GetControl('FSOLDEBALANCE')), SoldeBalD, SoldeBalC) ;
end;

procedure TOF_BALSIT_MODIF.OnGrilleRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  if not EnregistreLigne ( Ou ) then
  begin
    MessageAlerte ('Ligne incorrecte.');
    Cancel := True;
    exit;
  end;
  AfficheSoldeBalance;
end;

procedure TOF_BALSIT_MODIF.OnChercherClick(Sender: TOBject);
begin
  FFindFirst := True;
  FFindDialog.Execute;
end;

procedure TOF_BALSIT_MODIF.OnFindText(Sender: TObject);
begin
  Rechercher (FGrille, FFindDialog, FFindFirst);
  if goEditing in FGrille.Options  then
    FGrille.Col := COL_DEBIT;
end;

procedure TOF_BALSIT_MODIF.OnGrilleCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var   RdMontant : double;
begin
  case ACol of
    COL_DEBIT, COL_CREDIT :
      if IsNumeric(FGrille.Cells[ACol,ARow]) then
      begin
        RdMontant := Valeur(FGrille.Cells[ACol,ARow]);
        Cancel    := RdMontant < 0; // on ne peut pas saisir de valeur negative
        if Cancel then
          PGIInfo('Montant négatif !',ECRAN.Caption);
        if Arrondi(RdMontant,V_PGI.OkDecV) = 0 then
          FGrille.Cells[ACol,ARow] := ''
        else FGrille.Cells[ACol,ARow] := STRFMONTANT ( RdMontant , 15 , V_PGI.OkDecV, '' , true);
      end;
  end;
// CA - 13/05/2002 - Inutile ! Pourquoi était-ce là ?
//  SetControlProperty ('BVALIDER','Enabled', ( FGrille.Cells[COL_DEBIT,ARow] <> '' ) or ( FGrille.Cells[COL_CREDIT,ARow] <> '' ));
end;

function TOF_BALSIT_MODIF.EnregistreLigne(Ligne: integer) : boolean;
var Debit, Credit : double;
begin
  Debit := Valeur (FGrille.Cells[COL_DEBIT,Ligne]);
  Credit := Valeur (FGrille.Cells[COL_CREDIT,Ligne]);
  if (Debit < 0) or (Credit < 0) then Result := False
  else
  begin
    FBalSit.MajEcriture ( Ligne-1 , Debit , Credit);
    Result := True;
  end;
end;

procedure TOF_BALSIT_MODIF.OnFormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit ;
 if Key = VK_ESCAPE then FGrille.CacheEdit;

  case Key of
    VK_RETURN :
      begin
        Key := VK_TAB;
      end;
    VK_TAB :
      begin
        if (Shift = [ssShift]) and (FGrille.Col = COL_DEBIT) and (FGrille.Row = 1) then
          THCritMaskEdit(GetControl('FABREGE')).SetFocus
        else if (FGrille.Col = COL_CREDIT) and (FGrille.Row = FGrille.RowCount-1) then
          THCritMaskEdit(GetControl('FLIBELLE')).SetFocus;
      end;
    Ord('F')  :
      if Shift = [ssCtrl] then
      begin
        TToolBarButton97(GetControl('BCHERCHER')).Click;
        Key := 0;
      end;
    VK_ESCAPE :
      begin
//        FGrille.CacheEdit;
        Key := 0;
      end;
    VK_F10 :
      begin
        TToolBarButton97(GetControl('BVALIDER')).Click;
        Key := 0;
      end;
    VK_F6 :
      if (FBalSit.TypeBalance<>'DYN') then
      begin
        SoldeBalance;
        Key := 0;
      end;
    VK_INSERT :
      // CA - 18/06/2007 - FQ 19814 - Ne pas autoriser la modification d'une balance dynamique
      if (Shift = [ssCtrl]) and (FBalSit.TypeBalance<>'DYN') then
      begin
        TToolBarButton97(GetControl('BINSERT')).Click;
        Key := 0;
      end else if (FBalSit.TypeBalance<>'DYN') then
      begin
        TToolBarButton97(GetControl('BINSERTIONCOMPTE')).Click;
        Key := 0;
      end;
  end;
end;

procedure TOF_BALSIT_MODIF.OnNouveauCompteClick(Sender: TOBject);
var St, Compte : string;
    Col, Row : integer;
begin
  St := FicheGeneCreateOne;
  if St <> '' then
  begin
    Compte := ReadTokenSt (St);
    if Presence ('GENERAUX','G_GENERAL',Compte) then
    begin
      Col := FGrille.Col;
      Row := FGrille.Row;
      FBalSit.AjouteEcriture(Compte,'',0,0);
      FBalSit.TriParCompte;
      RemplirGrille;
      FGrille.SetFocus;
      FGrille.Col := Col;
      FGrille.Row := Row;
    end;
  end;
end;

procedure TOF_BALSIT_MODIF.RemplirGrille;
var i : integer;
    TLigne : TOB;
    Table, Colonne, Where : string;
begin
  FGrille.SynEnabled := False;
  FGrille.VidePile ( False );

  if FBalSit.NbEcriture = 0 then
    FGrille.RowCount := 2
  else
    FGrille.RowCount := FBalSit.NbEcriture+1;

  for i:=0 to FBalSit.NbEcriture-1 do
  begin
    TLigne := FBalSit.GetEcriture ( i );
    FGrille.Cells[COL_COMPTE1, i+1] := TLigne.GetValue('BSE_COMPTE1');
    if (FBalSit.TypeCumul = 'G/T') or (FBalSit.TypeCumul = 'G/A') then
      FGrille.Cells[COL_COMPTE2, i+1] := TLigne.GetValue('BSE_COMPTE2');
    if not TLigne.FieldExists('LIBELLE') then
    begin // Cas du libellé absent : suite ajout d'une ligne dans la grille
          // On récupère le libellé dans la table ad-hoc et on renseigne TLigne
      if FBalSit.TypeCumul = 'GEN' then
      begin
        Table := 'GENERAUX';
        Colonne := 'G_LIBELLE';
        Where := 'G_GENERAL="'+TLigne.GetValue('BSE_COMPTE1')+'"';
      end else
      if (FBalSit.TypeCumul = 'G/T') or (FBalSit.TypeCumul = 'TIE') then
      begin
        Table := 'TIERS';
        Colonne := 'T_LIBELLE';
        if (FBalSit.TypeCumul = 'TIE') then
          Where := 'T_AUXILIAIRE="'+TLigne.GetValue('BSE_COMPTE1')+'"'
        else Where := 'T_AUXILIAIRE="'+TLigne.GetValue('BSE_COMPTE2')+'"';
      end else
      if (FBalSit.TypeCumul = 'G/A') or (FBalSit.TypeCumul = 'ANA') then
      begin
        Table := 'SECTION';
        Colonne := 'S_LIBELLE';
        if (FBalSit.TypeCumul = 'ANA') then
          Where := 'S_SECTION="'+TLigne.GetValue('BSE_COMPTE1')+'"'
        else Where := 'S_SECTION="'+TLigne.GetValue('BSE_COMPTE2')+'"';
      end;
      TLigne.AddChampSupValeur('LIBELLE', GetColonneSQL(Table,Colonne,Where));
    end;
    FGrille.Cells[COL_LIBELLE, i+1] := TLigne.GetValue('LIBELLE');
    if TLigne.GetValue('BSE_DEBIT') = 0 then
      FGrille.Cells[COL_DEBIT, i+1] := ''
    else FGrille.Cells[COL_DEBIT, i+1] := STRFMONTANT ( TLigne.GetValue('BSE_DEBIT') , 15 , V_PGI.OkDecV, '' , true);
    if TLigne.GetValue('BSE_CREDIT') = 0 then
      FGrille.Cells[COL_CREDIT, i+1] := ''
    else FGrille.Cells[COL_CREDIT, i+1] := STRFMONTANT ( TLigne.GetValue('BSE_CREDIT') , 15 , V_PGI.OkDecV, '' , true);
  end;
  FGrille.SynEnabled := True;
end;

procedure TOF_BALSIT_MODIF.OnInsertionCompteClick(Sender: TOBject);
var E : THCritMaskEdit;
    Compte1, Compte2 : string;
begin
  // Insertion d'une ligne de compte
  E := THCritMaskEdit.Create ( nil );
  try
    E.Parent := FGrille;
    E.Visible := False;
    E.Left := FGrille.Left+ (FGrille.Width div 2);
    if (FBalSit.TypeCumul='GEN') then
    begin
      LookupList(E,TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE','','G_GENERAL',TRUE, 1);
      Compte1 := E.Text;
      Compte2 := '';
    end
    else if (FBalSit.TypeCumul='TIE') or (FBalSit.TypeCumul='G/T') then
    begin
      LookupList(E,TraduireMemoire('Comptes'),'TIERS','T_AUXILIAIRE','T_LIBELLE','','T_AUXILIAIRE',TRUE, 2);
      if (FBalSit.TypeCumul='G/T') then
      begin
        Compte2 := E.Text;
        Compte1 := GetColonneSQL('TIERS','T_COLLECTIF','T_AUXILIAIRE="'+Compte2+'"');
      end else Compte1 := E.Text;
    end
    else if (FBalSit.TypeCumul='ANA') or (FBalSit.TypeCumul='G/A') then
    begin
      LookupList(E,TraduireMemoire('Sections'),'SECTION','S_SECTION','S_LIBELLE','','S_SECTION',TRUE, 3);
      if (FBalSit.TypeCumul='G/A') then
      begin
        PGIBox ('Non supporté',ECRAN.Caption);
        E.Free;
        exit;
      end else Compte1 := E.Text;
    end;
    if E.Text <> '' then
    begin
      if FBalSit.PresenceCompte(Compte1, Compte2)=-1 then  // Compte non trouvé
      begin
        FBalSit.AjouteEcriture(Compte1,Compte2,0,0);
        FBalSit.TriParCompte;
        RemplirGrille;
        PositionneSurLigne( Compte1, Compte2 );
      end else PGIInfo('Ce compte est déjà présent dans la balance.',ECRAN.Caption);
    end;
  finally
    E.Free;
  end;
end;

procedure TOF_BALSIT_MODIF.OnPurgeClick(Sender: TOBject);
begin
  // Purge des lignes à 0
  if PGIAsk ('Voulez-vous supprimer les lignes à 0 ?',ECRAN.Caption) = mrYes then
  begin
    FBalSit.SuppEcritureAZero;
    RemplirGrille;
    FGrille.SetFocus;
  end;
end;

procedure TOF_BALSIT_MODIF.OnImprimerClick(Sender: TOBject);
begin
{$IFDEF EAGLCLIENT}
  TObjEtats.GenereEtatGrille (FGrille,ECRAN.Caption,False);
{$ELSE}
  PrintDBGrid(FGRILLE,nil,TraduireMemoire(ECRAN.Caption),'');
{$ENDIF}
end;

procedure TOF_BALSIT_MODIF.OnSuppressionCompteClick(Sender: TObject);
var Row : integer;
    Compte2: string;
begin
  if (PGIAsk ('Confirmez-vous la suppression de la ligne courante ?',ECRAN.Caption) = mrYes) then
  begin
    Row := FGrille.Row;
    if (FBalSit.TypeCumul='GEN') or (FBalSit.TypeCumul='TIE') or
          (FBalSit.TypeCumul='ANA') then Compte2 := ''
    else Compte2 := FGrille.Cells[COL_COMPTE2,Row];
    if FBalSit.SupprimeLigne ( FGrille.Cells[COL_COMPTE1,Row], Compte2) then
    begin
      FGrille.DeleteRow(Row);
    end;
  end;
end;

procedure TOF_BALSIT_MODIF.SoldeBalance;
var Col,Row : integer;
    T : THNumEdit;
    Sld , SldBal : double;
begin
  Col := FGrille.Col;
  Row := FGrille.Row;
  if ((Col=COL_DEBIT) or (Col=COL_CREDIT)) and (Row > 0) and (Row < FGrille.RowCount) then
  begin
    AfficheSoldeBalance; // Pour prendre en compte la ligne en cours dans le solde de la balance
    T := THNumEdit(GetControl('FSOLDEBALANCE'));
    if T.Debit then SldBal := T.Value else SldBal := (-1)*T.Value;
    Sld := SldBal - Valeur(FGrille.Cells[COL_DEBIT,Row])+ Valeur(FGrille.Cells[COL_CREDIT,Row]);
    if Sld >=0 then
    begin
      FGrille.Cells[COL_DEBIT,Row]:='';
      FGrille.Cells[COL_CREDIT,Row]:= StrFMontant(Sld,15 , V_PGI.OkDecV, '' , true);
    end else
    begin
      FGrille.Cells[COL_CREDIT,Row]:='';
      FGrille.Cells[COL_DEBIT,Row]:= StrFMontant((-1)*Sld,15 , V_PGI.OkDecV, '' , true);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 26/12/2001
Modifié le ... :   /  /    
Description .. : Positionne le focus sur la ligne associée aux comptes 
Suite ........ : passés en paramètres.
Mots clefs ... : FOCUS;COMPTE
*****************************************************************}
procedure TOF_BALSIT_MODIF.PositionneSurLigne(Compte1, Compte2: string);
var l : integer;
begin
  l := FBalSit.PresenceCompte ( Compte1, Compte2);
  if l >= 0 then
  begin
    FGrille.Row := l+1;
    FGrille.Col := COL_DEBIT;
  end;
end;

Initialization
  registerclasses ( [ TOF_BALSIT_MODIF ] ) ;
end.
