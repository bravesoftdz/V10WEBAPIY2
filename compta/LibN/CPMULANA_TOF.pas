{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/03/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPMULANA ()
Mots clefs ... : TOF;CPMULANA
*****************************************************************}
Unit CPMULANA_TOF ;

Interface

Uses StdCtrls,
     Controls, 
     ParamSoc, //SG6 23/12/2004 Gestion croisaxe
     Classes,
{$IFDEF EAGLCLIENT}
     eMul,
     uTob,
     UtileAGL,
     MaineAGL, // AGLLanceFiche
{$ELSE}
     Db,
     Hdb,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     FE_Main,  // AGLLanceFiche
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox, 
     UTOF,
     HQry,    // RecupWhereCritere
     Htb97,   // TToolBarButton97
     Ent1,    // ExoToDates
     AglInit, // StringToAction
     CritEdt; // TCritEdt

Type
  TOF_CPMULANA = Class (TOF)

    Pages             : TPageControl;
    TablesLibres      : TTabSheet;
    
    TY_AXE            : THLabel;

    Y_GENERAL        : THEdit;
    Y_SECTION        : THEdit;
    Y_DATECOMPTABLE  : THEdit;
    Y_DATECOMPTABLE_ : THEdit;
    Y_DATECREATION   : THEdit;
    Y_DATECREATION_  : THEdit;
    Y_NUMVENTIL      : THEdit;
    Y_NUMVENTIL_     : THEdit;
    XX_WHERE         : THEdit;

    Y_EXERCICE       : THValComboBox;
    Y_JOURNAL        : THValComboBox;
    Y_TYPEANALYTIQUE : THValComboBox;
    Y_QUALIFPIECE    : THValComboBox;
    Y_DEVISE         : THValComboBox;
    Y_ETABLISSEMENT  : THValComboBox;
    Y_AXE            : THValComboBox;

    RUne             : TCheckBox;
    BTobV            : TToolBarbutton97;

{$IFDEF EAGLCLIENT}
    FListe : THGrid ;
{$ELSE}
    FListe : THDBGrid ;
{$ENDIF}

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    procedure OnChangeY_Exercice    ( Sender : TObject );
    procedure OnChangeY_Axe         ( Sender : TObject );
    procedure OnChangeY_QualifPiece ( Sender : TObject );
    procedure OnClickRUne           ( Sender : TObject );
    procedure OnDblClickFListe      ( Sender : TObject );
    procedure OnClickBTobV          ( Sender : TObject );

  private
    FBoAvecANouveau : Boolean;
    FBoOkZoom       : Boolean;
    FCritEdt        : TCritEdt ;
    procedure InitCriteres;
  end ;

procedure MultiCritereAna(Comment : TActionFiche);
procedure MultiCritereAnaANouveau(Comment : TActionFiche);
procedure MultiCritereAnaZoom(Comment : TActionFiche;var CritEdt : TCritEdt);
procedure ModeCreation( vBoAvecAnouveau : Boolean );

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  ULibExercice,
  CPProcMetier,
  {$ENDIF MODENT1}
{$IFDEF COMPTA}
  SaisODA,  // LanceSaisieODA
{$ENDIF}
  SaisUtil; // RMVT

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/03/2004
Modifié le ... :   /  /
Description .. : Appel du Multicritère Analytique sans les A-Nouveaux
Mots clefs ... :
*****************************************************************}
procedure MultiCritereAna(Comment : TActionFiche);
begin
  case Comment of
    taCreat   : ModeCreation( False );
    taConsult : AGLLanceFiche('CP', 'CPMULANA', '', '', 'ACTION=CONSULTATION;FALSE;FALSE');
    taModif   : AGLLanceFiche('CP', 'CPMULANA', '', '', 'ACTION=MODIFICATION;FALSE;FALSE');
  end ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/03/2004
Modifié le ... :   /  /
Description .. : Appel du Multicritère Analytique avec les A-Nouveaux
Mots clefs ... :
*****************************************************************}
procedure MultiCritereAnaANouveau(Comment : TActionFiche);
begin
  case Comment of
    taCreat   : ModeCreation( True );
    taConsult : AGLLanceFiche('CP', 'CPMULANA', '', '', 'ACTION=CONSULTATION;TRUE;FALSE');
    taModif   : AGLLanceFiche('CP', 'CPMULANA', '', '', 'ACTION=MODIFICATION;TRUE;FALSE');
  end ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure MultiCritereAnaZoom(Comment : TActionFiche ; var CritEdt : TCritEdt);
begin
  case Comment of
    taCreat   : ModeCreation( False );
    taConsult : AGLLanceFiche('CP', 'CPMULANA', '', '', 'ACTION=CONSULTATION;FALSE;TRUE');
    taModif   : AGLLanceFiche('CP', 'CPMULANA', '', '', 'ACTION=MODIFICATION;FALSE;TRUE');
  end ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/04/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure ModeCreation(vBoAvecAnouveau: Boolean);
{$IFDEF COMPTA}
var M : RMVT ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  FillChar(M,Sizeof(M),#0) ;
  M.Simul := 'N' ;
  M.CodeD := V_PGI.DevisePivot ;
  M.DateC := V_PGI.DateEntree ;
  M.TauxD := 1 ;
  M.DateTaux := M.DateC ;
  M.Valide := False ;
  M.Etabl := VH^.ETABLISDEFAUT ;
  if vBoAvecANouveau then M.Anouveau := TRUE ;
  LanceSaisieODA(Nil,taCreat,M) ;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULANA.OnArgument (S : String ) ;
var lStArgument : string;
    Date1,Date2 : TDateTime ;
begin
  Inherited ;
  if (TheData <> nil) and (TheData is (ClassCritEdt)) then
    FCritEdt := ClassCritEdt(TheData).CritEdt;
    
  // Récupération de l'argument
  lStArgument := S;

  // Lecture du TypeAction
  TFMUL(Ecran).TypeAction := StringToAction(ReadTokenSt( lStArgument ));
  // Avec ou Sans les A-Nouveaux
  FBoAvecANouveau := ReadTokenSt( lStArgument ) = 'TRUE';
  // Ok Zoom
  FBoOkZoom := ReadTokenSt( lStArgument ) = 'TRUE';

  if FBoAvecANouveau then
  begin
    case TFMUL(Ecran).TypeAction of
      taCreat   : Ecran.HelpContext := 7728000 ;

      taConsult : begin
                    Ecran.Caption := TraduireMemoire('Visualisation des écritures d''A-Nouveaux analytiques');
                    Ecran.HelpContext := 7729000 ;
                  end;

      taModif   : begin
                    Ecran.Caption := TraduireMemoire('Modification des écritures d''A-Nouveaux analytiques');
                    Ecran.HelpContext := 7730000 ;
                  end;
    end ;
  end
  else
  begin
    case TFMUL(Ecran).TypeAction of

      taConsult : begin
                    Ecran.Caption := TraduireMemoire('Visualisation des écritures analytiques');
                    Ecran.HelpContext := 7364000 ;
                  end;

      taModif   : begin
                    Ecran.Caption := TraduireMemoire('Modification des écritures analytiques');
                    Ecran.HelpContext := 7370000 ;
                  end;
    end ;
  end ;

  // Récupération des composants de la fiche
  Pages             := TPageControl(GetControl('PAGES', True));
  TablesLibres      := TTabSheet(GetControl('TABLESLIBRES', True));
  
  TY_AXE            := THLabel(GetControl('TY_AXE'           , True));

  Y_GENERAL        := THEdit(GetControl('Y_GENERAL'         , True));
  Y_SECTION        := THEdit(GetControl('Y_SECTION'         , True));
  Y_DATECOMPTABLE  := THEdit(GetControl('Y_DATECOMPTABLE'   , True));
  Y_DATECOMPTABLE_ := THEdit(GetControl('Y_DATECOMPTABLE_'  , True));
  Y_DATECREATION   := THEdit(GetControl('Y_DATECREATION'    , True));
  Y_DATECREATION_  := THEdit(GetControl('Y_DATECREATION_'   , True));
  XX_WHERE         := THEdit(GetControl('XX_WHERE'          , True));

  Y_EXERCICE       := THValComboBox(GetControl('Y_EXERCICE', True));
  Y_JOURNAL        := THValComboBox(GetControl('Y_JOURNAL',  True));
  Y_TYPEANALYTIQUE := THValComboBox(GetControl('Y_TYPEANALYTIQUE', True));
  Y_QUALIFPIECE    := THValComboBox(GetControl('Y_QUALIFPIECE', True));
  Y_DEVISE         := THValComboBox(GetControl('Y_DEVISE', True));
  Y_ETABLISSEMENT  := THValComboBox(GetControl('Y_ETABLISSEMENT', True));
  Y_AXE            := THValComboBox(GetControl('Y_AXE', True));

  RUne             := TCheckBox(GetControl('RUNE', True));
  BTobV            := TToolBarButton97(GetControl('BTOBV', True));

{$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FListe', True));
{$ELSE}
  FListe := THDBGrid(GetControl('FListe', True));
{$ENDIF}

  // Tables Libres
  LibellesTableLibre(TablesLibres,'TY_TABLE','Y_TABLE','Y') ;

  // branchement des événements
  Y_Exercice.OnChange    := OnChangeY_Exercice;
  Y_Axe.OnChange         := OnChangeY_Axe;
  Y_QualifPiece.OnChange := OnChangeY_QualifPiece;
  RUne.OnClick           := OnClickRUne;
  FListe.OnDblClick      := OnDblClickFListe;
  BTobV.OnClick          := OnClickBTobV;

  //
  Y_Journal.ItemIndex       := 0;
  Y_Axe.ItemIndex           := 0;
  Y_DEvise.ItemIndex        := 0;
  Y_Etablissement.ItemIndex := 0;

  if EstSerie(S3) then
  begin
    Y_TYPEANALYTIQUE.Visible  := False ;
    Y_AXE.Visible             := False ;
    TY_AXE.Visible            := False ;
  end;  

  TFMul( Ecran ).Q.Manuel := True ; // Evite l'exécution de la requête lors de la maj de Q.Liste
  // Affectation de la liste
  Case TFMul( Ecran ).TypeAction of
    taConsult : TFMul( Ecran ).Q.Liste := 'MULVANAL';
    taModif   : TFMul( Ecran ).Q.Liste := 'MULMANAL';
  end;
  TFMul( Ecran ).Q.Manuel:=False ;

  InitCriteres;

  if (FBoOkZoom and ( FCritEDT.LibreCodes1 <> '' )) then
    Y_Journal.Value := FCritEDT.LibreCodes1
  else if (FBoOkZoom and ( FCritEDT.sCpt1 <> '' )) then
  else
    Y_Journal.Value := SQLPremierDernier(fbJal,TRUE) ;

  Y_Exercice.Value := QuelDateExo(V_PGI.DateEntree, Date1, Date2) ;

  if FBoAvecANouveau Then
  begin
    Y_DateComptable.Text  := DateToStr(VH^.Entree.Deb) ;
    Y_DateComptable_.Text := DateToStr(VH^.Entree.Deb) ;
  end
  else
  begin
    Y_DateComptable.Text  := DateToStr(V_PGI.DateEntree) ;
    Y_DateComptable_.Text := DateToStr(V_PGI.DateEntree) ;
  end;

  Y_DateCreation.Text  := StDate1900 ;
  Y_DateCreation_.Text := StDate2099 ;
  Y_Axe.Value          := 'A1' ;

  PositionneEtabUser(Y_ETABLISSEMENT) ;
  if FBoAvecANouveau And (Y_Journal.Items.Count > 0) then
  begin
    if Y_Journal.Vide then
      Y_Journal.ItemIndex := 1
    else
      Y_Journal.ItemIndex := 0 ;
  end;

  if FBoOkZoom Then
  begin
    if FCritEdt.Exo.Code <> '' then
      Y_Exercice.Value := FCritEdt.Exo.Code ;

    Y_DateComptable.Text  := DateToStr(FCritEdt.Date1) ;
    Y_DateComptable_.Text := DateToStr(FCritEdt.Date2) ;
    if FCritEdt.Cpt1 <> '' then
    begin
      Y_Section.Text      := FCritEdt.Cpt1 ;
      Y_Axe.Value         := FCritEdt.Bal.Axe ;
      Y_Journal.ItemIndex := 0 ;
    end;
    if FCritEdt.sCpt1 <> '' then
    begin
      Y_General.Text := FCritEdt.sCpt1 ;
    end;
  end;

  //SG6 23/12/04 Gestion croisaxe
  if VH^.AnaCroisaxe then
  begin
    SetControlVisible('TY_SECTION',False);
    Y_SECTION.Visible := False;
    Y_AXE.Visible := False;
    TY_AXE.Visible := False;
  end;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/03/2004
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULANA.OnLoad ;
begin
  inherited;
  if (Y_GENERAL.Text <> '') or (Y_SECTION.Text <> '') then
    RUne.Checked := False ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/03/2004
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULANA.OnDisplay () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPMULANA.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPMULANA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPMULANA.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPMULANA.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPMULANA.OnClose ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/03/2004
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULANA.OnChangeY_Exercice(Sender: TObject);
begin
  inherited;
  ExoToDates(Y_EXERCICE.Value,Y_DATECOMPTABLE,Y_DATECOMPTABLE_);
  { FQ 21085 BVE 18.07.07 }
  if Y_EXERCICE.Value= '' then
  begin
    Y_DATECOMPTABLE.Text  := stDate1900;
    Y_DATECOMPTABLE_.Text := stDate2099;
  end;
  { END FQ 21085 }
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULANA.OnChangeY_Axe(Sender: TObject);
var i : Byte ;
begin
  inherited;
  if Y_AXE.Value = '' then Exit ;
  i := StrToInt(Copy(Y_AXE.Value,2,1)) ;
  if i = 1 then
    Y_SECTION.DataType := 'TZSECTION'
  else
    Y_SECTION.DataType := 'TZSECTION' + IntToStr(i);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULANA.OnChangeY_QualifPiece(Sender: TObject);
begin
  inherited;
  if Y_QUALIFPIECE.Value <> 'R' then Exit ;
  if Not V_PGI.Controleur then Y_QUALIFPIECE.Value := 'N' ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.*********************************************** 
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/03/2004
Modifié le ... :   /  /    
Description .. : Initialisation des critères
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULANA.InitCriteres;
begin
  if VH^.Precedent.Code <> '' then
    Y_DATECOMPTABLE.Text := DateToStr(VH^.Precedent.Deb)
  else
    Y_DATECOMPTABLE.Text := DateToStr(VH^.Encours.Deb) ;

  Y_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;

  if FBoAvecANouveau Then
  begin
    Y_JOURNAL.DataType       := 'TTJALANALAN' ;
    XX_WHERE.Text            := 'Y_ECRANOUVEAU = "OAN"';
    Y_TYPEANALYTIQUE.Enabled := False ;
    Y_QUALIFPIECE.Enabled    := False ;
    Y_DEVISE.Enabled         := False ;
    SetControlText('Y_QUALIFPIECE', 'N') ; // Sélection du N à la place du itemIndex:=0 car pas toujours la même lettre en premier !
    Y_DEVISE.ItemIndex       := 0 ;
  end
  else
    XX_WHERE.Text := 'Y_ECRANOUVEAU <> "OAN"' ;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULANA.OnClickRUne(Sender: TObject);
begin
  inherited;
  if RUne.Checked then
  begin
    Y_NUMVENTIL.Text  := '1' ;
    Y_NUMVENTIL_.Text := '1' ;
  end
  else
  begin
    Y_NUMVENTIL.Text  := '0' ;
    Y_NUMVENTIL_.text := '9999' ;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULANA.OnDblClickFListe(Sender: TObject);
begin
  inherited;
{$IFDEF COMPTA}
{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
  TheMulQ.Seek(FListe.Row-1);
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}
  if ( TheMulQ.Eof ) and ( TheMulQ.Bof ) then Exit ;
  TrouveEtLanceSaisieODA( TheMulQ, TFMul(Ecran).TypeAction) ;
{$ENDIF}  
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/04/2004
Modifié le ... : 23/09/2005
Description .. : l'accès n'était ouvert QUE si pas de journal
                 ET pas de General ET pas de section
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULANA.OnClickBTobV(Sender: TObject);
var lWhereSQL : String ;
begin
  inherited;
  if ((Y_Journal.Value <> '') or (Y_General.Text <>'') or (Y_Section.Text <> '')) or
    (PgiAsk('Vous n''avez pas renseigné de journal ni de section. La ' +
              'sélection peut être importante. ' + #13 + #10 +
              'Confirmez-vous l''analyse ?', Ecran.Caption) = MrYes) then
    begin
      lWhereSQL := RecupWhereCritere(Pages) ;
      AGLLanceFiche('CP','CPCONSULTREVIS','','',lWhereSQL) ;
    end;

end;

Initialization
  registerclasses ( [ TOF_CPMULANA ] ) ;
end.


