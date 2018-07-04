{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 02/11/2006
Modifié le ... : 23/05/2007
Description .. : Source TOF de la FICHE : CPTABLEAUVAR ()
Suite ........ : - FQ 20385 - CA - Les 2 colonnes Augmentation et 
Suite ........ : Diminution sont saisissables
Suite ........ : - FQ 20384 - CA -  Proposition d'enregistrement du tableau 
Suite ........ : pour obtenir les valeurs en temps réel
Suite ........ : - FQ 20367 - CA - On doit enregistrer toutes les lignes du 
Suite ........ : tableau
Mots clefs ... : TOF;CPTABLEAUVAR
*****************************************************************}
Unit UtofCPTABLEAUVAR ;

Interface

Uses StdCtrls,
     Controls,
     Windows,
     Graphics,
     Classes,
     vierge,
{$IFNDEF EAGLCLIENT}
     db,
     Fe_Main,   //Agllancefiche
     EdtREtat,       // LanceEtat
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     UtileAgl,    //LanceEtat
     MainEAgl,    //AglLancefiche
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,   //STRFMONTANT
     HMsgBox,
     CPREVDocTravaux_TOF,   // CPLanceFiche_CPRevDocTravaux
//     SaisUtil,
     HTb97,          //TToolBarButton97
     HSysMenu,       //THSystemMenu
     ulibWindows,    //SetGridGrise
     ParamDat,       //ParamDat
     Menus,          //TPopupMenu
     HrichOLE,       //THRichEditOle
     Dialogs,        //TSaveDialog
     Hxlspas,        //ExportGrid
     Ent1,           //VH^
     UTOF ;

const
//les colonnes du tableau
  clFixe                : integer = 0;
  clDate                : integer = 1;
  clCommentaire         : integer = 2;
  clDebut               : integer = 3;
  clAugmentation        : integer = 4;
  clDiminution          : integer = 5;
  clFin                 : integer = 6;
  clBlocNote            : integer = 7;


Type
  TOF_CPTABLEAUVAR = Class (TOF)
    private

    procedure RechercheDonnees;
    procedure CalculTotaux;
    procedure FormKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
    procedure AjouteLigne (Premiere : boolean = False ; EnDessous : boolean = True);
    procedure AjouteLigneOnClick (Sender : Tobject);
    procedure SortieSansSauvegarde;
    procedure SortieSansSauvegardeOnClick (Sender : Tobject);
    procedure SupprimeLigne;
    procedure SupprimeLigneOnClick (Sender : Tobject);
    Function  ValideLigne (LaLigne : integer ; SansCal : boolean = True) : boolean;
    procedure ValideLigneOnClick (Sender : Tobject);
    procedure FListeCellEnter (Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
    procedure FListeCellExit (Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
    procedure CalculEcart (LaLigne : integer);
    procedure DeleteOnClick (Sender : Tobject);
    Procedure CommMillesimeOnClick (Sender : Tobject);
    procedure CommMillesime;
    procedure CommCompteOnClick (Sender : Tobject);
    procedure CommCompte;
    procedure Repro (RTotal : Boolean = True);
    procedure BlocnoteOnClick (Sender : Tobject);
    procedure FListeOnRowEnter (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure FlisteOnRowExit (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure ExportOnClick (Sender : Tobject);
    procedure EnregTableau ( bFerme : boolean );
    procedure HPBOnClose (Sender : Tobject);
    procedure OnMouseUpFListe(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure ImprimeTableau;
    procedure BimprimeOnCLick (Sender : Tobject);
    procedure BlocNoteOnKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ReproTotalOnClick (Sender : Tobject);
    procedure ReproEleOnClick (Sender : Tobject);
    procedure TridateOnClick (Sender : Tobject);
    procedure TriMontantOnClick (Sender : Tobject);
    procedure TriLibelleOnClick (Sender : tobject);
    procedure LeTri (LeType : String) ;
    procedure BlocnoteOnclickF11(Sender: Tobject);
    procedure MajBlocNote(LaLigne: integer);
    procedure CycleCommentaireOnClick (Sender : Tobject);  //FQ20272
    procedure CycleObjectifOnClick (Sender : Tobject);
    procedure CycleSyntheseOnClick (Sender : Tobject);

    public
    CpGeneral           : String;               // compte général
    CodeExercice        : String;               // Code exercice
    StrLibCpt           : String;               // Libellé compte général
    CpSens              : String;               // Sens comptable
    DbTotCE             : Double;               // total solde crédit Exercice
    DbTotDE             : Double;               // total solde débit Exercice
    DbTotCP             : Double;               // total solde crédit Exercice n-1
    DbTotDP             : Double;               // total solde débit Exercice n-1
    SoldeComptable      : Double;              // solde comptable
    DateFinExercice     : String;              // date de fin d'exercice
    DateDebExercice     : String;              // date de début exercice
    TobVariation        : Tob;                  // la tob
    Tsolde              : Double;               // Solde comptable
    TriDateDesc         : Boolean;              // est un tri descendant sur date
    TriLibDesc          : Boolean;              // est un tri descendant sur libellé
    TriMtDesc           : Boolean;              // est un tri descendant sur montant

    ASaveDialog         : TSaveDialog;          // fenetre de dialogue pour export
    //Grille de saisie
    FListe              : THGrid;
    BValide             : TToolbarButton97;
    BInsert             : TToolbarButton97;
    BSupprime           : TToolbarButton97;
    BDelete             : TToolbarButton97;
    BFerme              : TToolbarButton97;
    BExport             : TToolbarButton97;
    BImprime            : TToolbarButton97;
    MenuCom             : TPopupMenu;
    PopF11              : TPopupMenu;
    PopTri              : TPopupMenu;

    //les totaux
    TH11                : THLabel;    // tot debut
    TH12                : THLabel;    // tot augmentation
    TH13                : THLabel;    // tot diminution
    TH14                : THLabel;    // tot fin
    TH21                : THLabel;    // solde debut
    TH24                : THLabel;    // solde fin
    TH31                : THLabel;    // ecart debut
    TH34                : THLabel;    // ecart fin
    //le bloc notes
    BBlocNote           : TToolBarButton97;
    FBlocNote           : THRichEditOle;
    HPB                 : TToolWindow97;
    CpCycleRevision     : String;               // Cycle de révision     //FQ20272

    procedure OnArgument (S : string ) ; override ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;


  end ;

procedure CPLanceFiche_CPTableauVar( vStCompte : string );

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  AGLInit, // ActionToString
  uLibRevision,
  Grids;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 10/11/2006
Modifié le ... : 13/06/2007
Description .. : Appel de la fiche
Suite ........ : 
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_CPTableauVar( vStCompte : string );
begin
  AGLLanceFiche('CP', 'CPTABLEAUVAR', '', '', vStCompte);
end;

procedure TOF_CPTABLEAUVAR.OnArgument(S : string);
var
  TExercice             : THLabel;
  Q                     : TQuery;
  i,j                   : integer;
  SoldePrec             : double;
  lStEtatCycle          : string;
  lMenu                 : TMenuItem;

begin
  Inherited ;
  // recupere les criteres
  CPGeneral             := ReadTokenSt(S);
  TriDateDesc           := True;
  TriLibDesc            := True;
  TriMtDesc             := True;

  try
    Q                   := OpenSQL('SELECT G_LIBELLE, G_SENS, ' +
                          'G_TOTDEBE, G_TOTCREE, G_TOTDEBP, G_TOTCREP, G_CYCLEREVISION ' +
                          'FROM GENERAUX WHERE G_GENERAL = "' + CPGeneral + '"', True);

    if not Q.Eof then
    begin
      StrLibCpt         := Q.FindField('G_LIBELLE').AsString;
      CpSens            := Q.FindField('G_SENS').AsString;
      DbTotDE           := Q.FindField('G_TOTDEBE').AsFloat;
      DbTotCE           := Q.FindField('G_TOTCREE').AsFloat;
      DbTotDP           := Q.FindField('G_TOTDEBP').AsFloat;
      DbTotCP           := Q.FindField('G_TOTCREP').AsFloat;
      CpCycleRevision   := Q.FindField('G_CYCLEREVISION').AsString;
    end;

    CodeExercice        := VH^.EnCours.Code;

    // GCO - 13/06/2007 - FQ 20480
    // GCO - 11/07/2007 - FQ 20946 - Blocage en consultation suivant l'état de la révision
    if CpCycleRevision <> '' then
    begin
      lStEtatCycle := GetColonneSQL('CREVCYCLE', 'CCY_ETATCYCLE',
                      'CCY_CODECYCLE = "' + CpCycleRevision + '" AND  ' +
                      'CCY_EXERCICE = "' + CodeExercice  + '"');

      if (VH^.Revision.DossierSupervise) or
         (VH^.Revision.DossierPretSupervise and (not JaileRoleCompta(rcSuperviseur))) or
         ((lStEtatCycle = cValide) and (not JaileRoleCompta(rcSuperviseur))) or
         (lStEtatCycle = cSupervise) then
        TFVierge(Ecran).TypeAction := TaConsult;
    end;

  finally
    Ferme(Q);
  end;

  //les controles
  FListe                := THGrid(GetControl('GRVAR', True));
  TExercice             := THLabel(GetControl('TEXERCICE', True));
  BValide               := TToolbarButton97(GetControl('BVALIDER', True));
  BInsert               := TToolbarButton97(GetControl('BINSERT', True));
  BSupprime             := TToolbarButton97(GetControl('BSUPPRIME', True));
  BDelete               := TToolbarButton97(GetControl('BDELETE', True));
  BFerme                := TToolbarButton97(GetControl('BFERME', True));
  BBlocNote             := TToolbarButton97(GetControl('BBLOCNOTE', True));
  BExport               := TToolbarButton97(GetControl('BEXPORT', True));
  BImprime              := TToolbarButton97(GetControl('BIMPRIMER', True));
  MenuCom               := TPopupMenu(GetControl('PM_COMM', True));
  PopF11                := TPopupMenu(GetControl('POPF11', True));
  PopTri                := TPopupMenu(GetControl('POPTRI', True));
  TH11                  := THLabel(GetControl('TTOTDEB', True));
  TH12                  := THLabel(GetControl('TTOTAUG', True));
  TH13                  := THLabel(GetControl('TTOTDIM', True));
  TH14                  := THLabel(GetControl('TTOTFIN', True));
  TH21                  := THLabel(GetControl('TSOLDEDEB', True));
  TH24                  := THLabel(GetControl('TSOLDEFIN', True));
  TH31                  := THLabel(GetControl('TECARTDEB', True));
  TH34                  := THLabel(GetControl('TECARTFIN', True));

  // Création de la fenêtre BLOC NOTE
  HPB                   := TToolWindow97.Create(nil);
  HPB.Parent            := Ecran;
  HPB.Caption           := 'Bloc note';
  HPB.Visible           := False;
  HPB.Left              := 350;
  HPB.Top               := 260;
  HPB.Height            := 140;
  HPB.width             := 250;

  FBlocNote             := THRichEditOle.Create(nil);
  FBlocNote.Parent      := HPB;
  FBlocNote.Align       := AlClient;
  FBlocNote.Clear;
  FBlocNote.OnKeyDown   := BlocNoteOnKeyDown;        
  { FQ 20826 BVE 25.06.07 }
  FBlocNote.Font.Height := -11;
  { END FQ 20826 }

  //Boite dialogue pour export
  ASaveDialog := TSaveDialog.Create(Ecran);
  ASaveDialog.Filter    := 'Fichier Texte (*.txt)|*.txt|Fichier Excel (*.xls)|*.xls|Fichier Ascii (*.asc)|*.asc|Fichier Lotus (*.wks)|*.wks|Fichier HTML (*.html)|*.html|Fichier XML (*.xml)|*.xml';
  ASaveDialog.DefaultExt  := 'XLS';
  ASaveDialog.FilterIndex := 1;
  ASaveDialog.Options     := ASaveDialog.Options + [ofOverwritePrompt, ofPathMustExist, ofNoReadonlyReturn, ofNoLongNames] - [ofEnableSizing];

  // les proprietes
  Ecran.OnKeyDown       := FormKeyDown;
  FListe.OnCellEnter    := FListeCellEnter;
  FListe.OnCellExit     := FListeCellExit;
  FListe.OnMouseUp      := OnMouseUpFListe;
//  FListe.OnKeyDown := FormKeyDown;
  FListe.OnRowEnter     := FListeOnRowEnter;
  Fliste.OnRowExit      := FlisteOnRowExit;
  BValide.Visible       := True;
  BValide.OnClick       := ValideLigneOnClick;
  BFerme.OnClick        := SortieSansSauvegardeOnClick;
  BInsert.OnClick       := AjouteLigneOnClick;
  BDelete.Visible       := True;
  BDelete.OnClick       := DeleteOnClick;
  BSupprime.OnClick     := SupprimeLigneOnClick;
  BBlocNote.OnClick     := BlocnoteOnClick;
  BExport.OnClick       := ExportOnClick;
  BImprime.Visible      := True;
  Bimprime.OnClick      := BimprimeOnCLick;
  HPB.OnClose           := HPBOnClose;

  //menu pop commentaires
  if Assigned(MenuCom) then
  begin
    for i := 0 to MenuCom.Items.Count -1 do
    begin
      //cycle de révision                   //FQ20272
      if MenuCom.Items[i].Name = 'COMCYCLE' then
        if CpCycleRevision = '' then
          MenuCom.Items[i].Visible  := False
        else
          MenuCom.Items[i].OnClick  := CycleCommentaireOnClick;
      if MenuCom.Items[i].Name = 'OBJREV' then
        if CpCycleRevision = '' then
          MenuCom.Items[i].Visible  := False
        else
          MenuCom.Items[i].OnClick  := CycleObjectifOnClick;
      if MenuCom.Items[i].Name = 'SYNTHCYCLE' then
        if CpCycleRevision = '' then
          MenuCom.Items[i].Visible  := False
        else
          MenuCom.Items[i].OnClick  := CycleSyntheseOnClick;

      if MenuCom.Items[i].Name = 'COM' then
        MenuCom.Items[i].OnClick  := CommMillesimeOnClick;
      if MenuCom.Items[i].Name = 'COC' then
        MenuCom.Items[i].OnClick  := CommcompteOnClick;
    end;
  end;

  //menu pop Tri
  if Assigned(PopTri) then
  begin
    for i := 0 to PopTri.Items.Count -1 do
    begin
      if PopTri.Items[i].Name = 'TRIDATE' then
        PopTri.Items[i].OnClick := TriDateOnClick;
      if PopTri.Items[i].Name = 'TRILIBELLE' then
        PopTri.Items[i].OnClick := TriLibelleOnClick;
      if PopTri.Items[i].Name = 'TRIMONTANT' then
        PopTri.Items[i].OnClick := TriMontantOnClick;
    end;
  end;

  //menu pop conextuel
  if Assigned(PopF11) then
  begin
    for i := 0 to PopF11.Items.Count -1 do
    begin
      if PopF11.Items[i].Name = 'NOUVLIG' then
      begin
        PopF11.Items[i].OnClick := AjouteLigneOnClick;
        PopF11.Items[i].Visible := TFVierge(Ecran).TypeAction <> TaConsult;
      end;

      if PopF11.Items[i].Name = 'SUPLIG' then
      begin
        PopF11.Items[i].OnClick := SupprimeLigneOnClick;
        PopF11.Items[i].Visible := TFVierge(Ecran).TypeAction <> TaConsult;
      end;

      if PopF11.Items[i].Name = 'DUPLIG' then
      begin
        PopF11.Items[i].OnClick := ReproTotalOnClick;
        PopF11.Items[i].Visible := TFVierge(Ecran).TypeAction <> TaConsult;
      end;

      if PopF11.Items[i].Name = 'DUPELE' then
      begin
        PopF11.Items[i].OnClick := ReproEleOnClick;
        PopF11.Items[i].Visible := TFVierge(Ecran).TypeAction <> TaConsult;
      end;  

      if PopF11.Items[i].Name = 'NOTES' then
        PopF11.Items[i].OnClick := BlocnoteOnclickF11;

      if PopF11.Items[i].Name = 'DOCTRAVAUXF11' then
      begin
        for j := 0 to PopF11.Items[i].Count -1 do
        begin
          lMenu := PopF11.Items[i].Items[j];

          // GCO - 26/07/2007 - FQ 20976
          if lMenu.Name = 'COMCYCLEF11' then
          begin
            lMenu.Visible := CpCycleRevision <> '';
            lMenu.OnClick := CycleCommentaireOnClick;
          end;

          if lMenu.Name = 'OBJREVF11' then
          begin
            lMenu.Visible := CpCycleRevision <> '';
            lMenu.OnClick := CycleObjectifOnClick;
          end;

          if lMenu.Name = 'SYNTHCYCLEF11' then
          begin
            lMenu.Visible := CpCycleRevision <> '';
            lMenu.OnClick := CycleSyntheseOnClick;
          end;

          if lMenu.Name = 'COMM' then
            lMenu.OnClick := CommMillesimeOnClick;

          if lMenu.Name = 'COMC' then
            lMenu.OnClick := CommcompteOnClick;
        end;
      end; // FIN GCO

      if PopF11.Items[i].Name = 'IMP' then
        PopF11.Items[i].OnClick := BimprimeOnCLick;
      if PopF11.Items[i].Name = 'VAL' then
        PopF11.Items[i].OnClick := ValideLigneOnClick;
      if PopF11.Items[i].Name = 'EXIT' then
        PopF11.Items[i].OnClick := SortieSansSauvegardeOnClick;
    end;
  end;

  ActivateXpPopUp( PopF11 );

  // l'entete
  TForm(Ecran).Caption  := 'Tableau des variations : ' + CpGeneral + ' - ' + StrLibCpt;
  UpdateCaption(TForm(Ecran));
  Q := OpenSQL('SELECT EX_LIBELLE, EX_DATEFIN, EX_DATEDEBUT FROM EXERCICE WHERE EX_EXERCICE="'+CodeExercice+'"', True);
  TExercice.Caption     := 'Solde de fin d''exercice : ';
  DateFinExercice       := Q.Fields[1].AsString;
  DateDebExercice       := Q.Fields[2].AsString;
  Ferme(Q);

  // les parametres
  FListe.ColWidths[clFixe]          := 10;
  FListe.ColWidths[clDate]          := 80;
  FListe.ColWidths[clCommentaire]   := 180;
  FListe.ColWidths[clDebut]         := 90;
  FListe.ColWidths[clAugmentation]  := 90;
  FListe.ColWidths[clDiminution]    := 90;
  FListe.ColWidths[clFin]           := 90;
  FListe.ColWidths[clBlocNote]      := 40;       
  { FQ 20645 BVE 18.06.07 }
  FListe.ColLengths[clCommentaire]  := 35;
  { END FQ 20645 }
  FListe.TwoColors                  := True;
  FListe.ColTypes[clDate]           := 'D';
//  FListe.ColFormats[clDate]         := ShortDateFormat;
  Fliste.ColFormats[clDate]         := 'dd/mm/yyyy';  //sinon effet de bord // ??   //FQ19964
  FListe.ColTypes[clCommentaire]    := 'S';
  { FQ 20576 BVE 11.06.07 }
  FListe.CalcInCell                 := True;          // pour permettre le calcul dans les cellules
  FListe.ColTypes[clDebut]          := 'K';
  FListe.ColFormats[clDebut]        := 'R';
  FListe.ColAligns[clDebut]         := taRightJustify;
  FListe.ColTypes[clAugmentation]   := 'K';
  FListe.ColFormats[clAugmentation] := 'R';
  FListe.ColAligns[clAugmentation]  := taRightJustify;
  FListe.ColTypes[clDiminution]     := 'K';
  FListe.ColFormats[clDiminution]   := 'R';
  FListe.ColAligns[clDiminution]    := taRightJustify;
  FListe.ColTypes[clFin]            := 'R';
  FListe.ColFormats[clFin]          := '##,##0.00';    
  { END FQ 20576 BVE 11.06.07 }
  FListe.ColAligns[clFin]           := taRightJustify;
  FListe.ColAligns[clBlocNote]      := taCenter;
  FListe.Options                    := FListe.Options - [goRowSelect] + [goEditing] + [goTabs];
  FListe.ColEditables[clDate]       := True;
  FListe.ColEditables[clCommentaire]  := True;
  FListe.ColEditables[clDebut]        := True;
  FListe.ColEditables[clAugmentation] := True;
  FListe.ColEditables[clDiminution] := True;
  FListe.ColEditables[clFin]        := False;
  FListe.ColEditables[clBlocNote]   := False;

  FListe.Cells[clDate, 0]           :=  TraduireMemoire('Date');
  FListe.Cells[clCommentaire, 0]    := TraduireMemoire('Commentaire');
  FListe.Cells[clDebut, 0]          := TraduireMemoire('Début');
  FListe.Cells[clAugmentation, 0]   := TraduireMemoire('Augmentation');
  FListe.Cells[clDiminution, 0]     := TraduireMemoire('Diminution');
  FListe.Cells[clFin, 0]            := TraduireMemoire('Fin');
  FListe.Cells[clBlocNote, 0]       := TraduireMemoire('Notes');



  TobVariation          := Tob.Create('CPTABLEAUVAR', nil, -1);

  // Soldes comptables n et n-1
  if CpSens = 'C' then
  begin
    SoldePrec           := DbTotCP - DbTotDP;
    TH21.Caption        := STRFMONTANT(SoldePrec, 15, V_PGI.OkDecV, '', True);
    SoldeComptable      := DbTotCE - DbTotDE;
    TH24.Caption        := STRFMONTANT(SoldeComptable, 15, V_PGI.OkDecV, '', True);
  end
  else
  begin
    SoldePrec           := DbTotDP - DbTotCP;
    TH21.Caption        := STRFMONTANT(SoldePrec, 15, V_PGI.OkDecV, '', True);
    SoldeComptable      := DbTotDE - DbTotCE;
    TH24.Caption        := STRFMONTANT(SoldeComptable, 15, V_PGI.OkDecV, '', True);
  end;

  if DbTotDE >= DbTotCE then
  begin
    Tsolde := DbTotDE - DbTotCE;
    TExercice.Caption   := TExercice.Caption + STRFMONTANT(tsolde, 15, V_PGI.OkDecV, '', True) +  ' D au ' + DateFinExercice  end
  else
  begin
    Tsolde := DbTotCE - DbTotDE;
    TExercice.Caption   := TExercice.Caption + STRFMONTANT(tsolde, 15, V_PGI.OkDecV, '', True) +  ' C au ' + DateFinExercice
  end;

  // GCO - 13/06/2007 - FQ 20480
  if (TFVierge(Ecran).TypeAction in [TaConsult]) then
  begin
    BInsert.Visible       := False;
    BSupprime.Visible     := False;
    BDelete.Visible       := False;
    FBlocNote.ReadOnly    := True;
    FListe.EditorMode     := False;
    FListe.Options := FListe.Options + [goRowSelect] - [goEditing] + [goTabs];
  end;

  // appel des donnees
  RechercheDonnees;

end ;

procedure TOF_CPTABLEAUVAR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPTABLEAUVAR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPTABLEAUVAR.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPTABLEAUVAR.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPTABLEAUVAR.OnClose ;
begin
  FreeAndNil(FBlocNote);
  FreeAndNil(HPB);
  FreeAndNil(TobVariation);
  FreeAndNil(ASaveDialog);
  Inherited ;
end ;

procedure TOF_CPTABLEAUVAR.OnDisplay () ;
begin
  Inherited ;
  THSystemMenu(TFVierge(ECRAN).HMTrad).ResizeGridColumns(fListe);
end ;

procedure TOF_CPTABLEAUVAR.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/11/2006
Modifié le ... :   /  /    
Description .. : Appel des données
Mots clefs ... : TABLEAU
*****************************************************************}
procedure TOF_CPTABLEAUVAR.RechercheDonnees;
var
  StrSql                : String;
  CalFin                : double;
  CalDebut              : double;
  //DateDeb               : TDateTime;
  //Year,Month,Day        : Word;
  Q                     : Tquery;
  i                     : integer;
  LaLigne               : integer;

begin
  StrSql                := 'SELECT * FROM CPTABLEAUVAR WHERE CTV_EXERCICE="'+CodeExercice+
                           '" AND CTV_GENERAL="'+CpGeneral+'" ORDER BY CTV_LIGNE';
  Q                     := OpenSQL(StrSql,False);
  TobVariation.LoadDetailDB('CPTABLEAUVAR','','',Q,False);
  Ferme(Q);
  { FQ 20386 BVE 13.06.07 }
  if TobVariation.Detail.Count = 0 then
  begin
     // Il n'y a pas de données sur l'exercice courant. Test sur l'exercice précédent.
     StrSql             := 'SELECT * FROM CPTABLEAUVAR ' +
                           'WHERE CTV_EXERCICE = "' + VH^.Precedent.Code + '" ' +
                           'AND CTV_GENERAL = "' + CpGeneral + '" ' +
                           'AND ( CTV_MONTANT + CTV_AUGMENTATION - CTV_DIMINUTION ) <> 0 ' +
                           'ORDER BY CTV_LIGNE';
     Q                  := OpenSQL(StrSql,False);
     TobVariation.LoadDetailDB('CPTABLEAUVAR','','',Q,False);
     Ferme(Q);
     if TobVariation.Detail.Count <> 0 then
     begin
        // Il y a des données sur l'exercice précédent
        if ( PGIAsk('Il n''y a pas de données sur l''exercice en cours. Voulez-vous reprendre les données de l''exercice précédent ?') = mrYes ) then
        begin
           // Pour le calcul de la date.                
           //DecodeDate(VH^.EnCours.Deb,Year,Month,Day);
           //DateDeb := EncodeDate(Year, 1, 1);
           // On récupere les données de l'exercice précédent
           for i := 0 to TobVariation.Detail.Count - 1 do
           begin
              // On adapte les valeurs
              // Exercice = Exercie en cours
              TobVariation.Detail[i].PutValue('CTV_EXERCICE',CodeExercice);
              // Debut = Fin
              CalDebut  := TobVariation.Detail[i].GetValue('CTV_MONTANT') +
                           TobVariation.Detail[i].GetValue('CTV_AUGMENTATION') -
                           TobVariation.Detail[i].GetValue('CTV_DIMINUTION');
              TobVariation.Detail[i].PutValue('CTV_MONTANT',CalDebut);
              // Augmentation et Diminution = vide     
              TobVariation.Detail[i].PutValue('CTV_AUGMENTATION',0);
              TobVariation.Detail[i].PutValue('CTV_DIMINUTION',0);
              // La date = 1 Janvier de l'Exercice courant
              // GCO - 26/07/2007 - FQ 21078
              //TobVariation.Detail[i].PutValue('CTV_DATE', DateDeb);
           end;
        end
        else
        begin
           // On vide la TOB car on ne veut pas reprendre les données de l'exercice précédent
           TobVariation.ClearDetail;
        end;
     end;
  end;
  { END FQ 20386 }
  LaLigne               := 1;
  for i := 0 to TobVariation.Detail.Count -1 do
  begin
    CalFin              := TobVariation.Detail[i].GetValue('CTV_MONTANT')+
                           TobVariation.Detail[i].GetValue('CTV_AUGMENTATION')-
                           TobVariation.Detail[i].GetValue('CTV_DIMINUTION');
    TobVariation.Detail[i].AddChampSupValeur('CALCULFIN', CalFin);

    if i > 0 then
    begin
      inc(LaLigne);
      FListe.RowCount   := LaLigne+1;
    end;
    FListe.Cells[clDate, LaLigne]           := TobVariation.Detail[i].GetValue('CTV_DATE');
    FListe.Cells[clCommentaire, LaLigne]    := TobVariation.Detail[i].GetValue('CTV_COMMENTAIRE');
    FListe.Cells[clDebut, LaLigne]          := STRFMONTANT(TobVariation.Detail[i].GetValue('CTV_MONTANT'), 15, V_PGI.OkDecV, '', True);
    if TobVariation.Detail[i].GetValue('CTV_AUGMENTATION') <> 0 then
      FListe.Cells[clAugmentation, LaLigne] := STRFMONTANT(TobVariation.Detail[i].GetValue('CTV_AUGMENTATION'), 15, V_PGI.OkDecV, '', True);

    if TobVariation.Detail[i].GetValue('CTV_DIMINUTION') <> 0 then
      FListe.Cells[clDiminution, LaLigne]   := STRFMONTANT(TobVariation.Detail[i].GetValue('CTV_DIMINUTION'), 15, V_PGI.OkDecV, '', True);

    FListe.Cells[clFin, LaLigne]            := STRFMONTANT(TobVariation.Detail[i].GetValue('CALCULFIN'), 15, V_PGI.OkDecV, '', True);
    Fliste.Cells[clBlocNote, LaLigne]       := TobVariation.Detail[i].GetValue('CTV_BLOCNOTE');
    ValideLigne(LaLigne, False);
  end;

  // GCO - 27/07/2007 - FQ 20946
  if (not (TFVierge(Ecran).TypeAction in [TaConsult])) and
     (TobVariation.Detail.Count = 0) then
      AjouteLigne(True);

  CalculTotaux;
  FBlocNote.Clear;

  if TobVariation.Detail.Count <> 0 then
    StringtoRich(FBlocNote, TobVariation.Detail[0].GetValue('CTV_BLOCNOTE'));
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/11/2006
Modifié le ... :   /  /
Description .. : Calcul des totaux
Suite ........ : - cumul des colonnes
Suite ........ : - calcul des ecarts
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTABLEAUVAR.CalculTotaux;
var
  i                     : integer;
  TotalDeb              : double;
  TotalAug              : double;
  TotalDim              : Double;
  TotalFin              : double;
  EcartFin              : double;
  EcartDeb              : double;


begin
  TotalDeb              := 0;
  TotalAug              := 0;
  TotalDim              := 0;
  TotalFin              := 0;
  EcartFin              := 0;
  EcartDeb              := 0;

  for i := 0 to TobVariation.Detail.Count -1 do
  begin
    TotalDeb            := TotalDeb + TobVariation.Detail[i].GetValue('CTV_MONTANT');
    TotalAug            := TotalAug + TobVariation.Detail[i].GetValue('CTV_AUGMENTATION');
    TotalDim            := TotalDim + TobVariation.Detail[i].GetValue('CTV_DIMINUTION');
    TotalFin            := TotalFin + TobVariation.Detail[i].GetValue('CALCULFIN');
    EcartFin            := TotalFin - SoldeComptable;
    EcartDeb            := TotalDeb - Valeur(TH21.Caption);
  end;

  TH11.Caption          := STRFMONTANT(TotalDeb, 15, V_PGI.OkDecV, '', True);
  TH12.Caption          := STRFMONTANT(TotalAug, 15, V_PGI.OkDecV, '', True);
  TH13.Caption          := STRFMONTANT(TotalDim, 15, V_PGI.OkDecV, '', True);
  TH14.Caption          := STRFMONTANT(TotalFin, 15, V_PGI.OkDecV, '', True);
  TH31.Caption          := STRFMONTANT(EcartDeb, 15, V_PGI.OkDecV, '', True);
  TH34.Caption          := STRFMONTANT(EcartFin, 15, V_PGI.OkDecV, '', True);

  if EcartFin < 0 then
    TH34.Font.Color     := clRed
  else
    TH34.Font.Color     := clBlack;

  if EcartDeb < 0 then
    TH31.Font.Color     := clRed
  else
    TH31.Font.Color     := clBlack;
    
end;



{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/11/2006
Modifié le ... :   /  /
Description .. : Gestion des touches
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTABLEAUVAR.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin

  Case Key of
      VK_ESCAPE : if Shift = [] then
                    SortieSansSauvegarde;

      VK_INSERT : if (Shift = []) and (IsValidDate(FListe.Cells[clDate, FListe.Row])) then
                    begin
                      // GCO - 02/10/2007 - FQ 21553
                      if TFVierge(Ecran).TypeAction = TaConsult then Exit;
                      key     := 0;
                      AjouteLigne(false,false); // FQ 19965
                    end;

      VK_F10 :  if Shift = [] then
                  //ValideLigne(FListe.Row);
                  ValideLigneOnClick(nil);


      VK_DOWN : begin
                  if Shift = [] then
                  begin
                    if (FListe.Row = FListe.RowCount-1) and (ValideLigne(FListe.Row)) then
                    begin
                      if Valeur(TobVariation.Detail[Fliste.Row -1].GetValue('CTV_MONTANT')) <> 0 then
                        AjouteLigne;
                    end;
                  end;
                end;

      VK_UP :   begin
                  if (not ValideLigne(Fliste.Row)) and (Fliste.row > 1) then
                  begin
//                    Fliste.Row := Fliste.Row -1;
                    key := 0;
                    SupprimeLigne;  //FQ19965
//                    TobVariation.Detail[FListe.Row].Free;
//                    FListe.DeleteRow(FListe.Row +1);
                  end
                  else
                  if (Valeur(TobVariation.Detail[Fliste.Row -1].GetValue('CALCULFIN')) = 0) then
                  begin
                    key := 0;
                    SupprimeLigne;
                  end;


                end;

      VK_RIGHT :  begin
                    if (Fliste.Col = Fliste.ColCount -1) and (Fliste.Row < Fliste.RowCount -1) then
                    begin
                      Fliste.Col := clDate;
                      Fliste.Row := Fliste.Row +1;
                        key := 0;
//                      Fliste.SetFocus;
                    end;
                  end;

      VK_LEFT :  begin
                    if (Fliste.Col = clDate) and (Fliste.Row > 1)  and (ValideLigne(Fliste.Row)) then
                    begin
                      Fliste.Col := Fliste.ColCount -1;
                      Fliste.Row := Fliste.Row -1;
                      key := 0;
//                      Fliste.SetFocus;
                    end;
                  end;

      VK_TAB :  if Shift = [] then
                  if (FListe.Row = FListe.RowCount-1) and (IsValidDate(FListe.Cells[clDate, FListe.Row]))
                    and (FListe.Col = clBlocNote) then
                    begin
                      key := 0;
                      if ValideLigne(FListe.Row) then
                        AjouteLigne;
                    end;

      VK_RETURN : ValideLigne(Fliste.Row);

      VK_DELETE : begin
                    // GCO - 02/10/2007 - FQ 21553
                    if TFVierge(Ecran).TypeAction = TaConsult then Exit;

                    if Shift = [ssCtrl] then
                      SupprimeLigne;

                    if (Shift = []) and (not FListe.EditorMode) then  //FQ19965
                    begin
                      FListe.Cells[Fliste.Col, Fliste.Row]  := '';
                    end;
                  end;

      VK_F11 :
          PopF11.Popup (Mouse.CursorPos.x, Mouse.CursorPos.y);

      VK_F4 : if Shift = [ssAlt] then
                SortieSansSauvegarde;

      VK_F5 : if (Shift = []) and (FListe.Col = clDate) then
              begin
                { FQ 19964 BVE 12.06.07 }
                FListe.ShowEditor;
                { END FQ 19664 }
                V_PGI.ParamDateproc(FListe);
              end;
      VK_F6 : if Shift = [] then
//              begin
//                if FListe.Col = clDebut then
//                begin
                  CalculEcart(FListe.Row);
//                end;
//              end;

      VK_F7 : Begin
                if Shift = [] then
                  Repro(False);
                if Shift = [ssAlt] then
                  Repro(True);
              end;

{^P}  80 :  if Shift = [ssCtrl] then
              ImprimeTableau;

{AB}  66 :  begin
              if Shift = [ssAlt] then
              begin
                BlocNoteOnCLick(nil);
              end;
            end;

{AO}  79 :  if Shift = [ssAlt] then
              CommMillesime;

{AC}  67 :  if Shift = [ssAlt] then
              CommCompte;
  end;



end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 02/10/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTABLEAUVAR.AjouteLigneOnClick(Sender: Tobject);
begin
  // GCO - 02/10/2007 - FQ 21553
  if TFVierge(Ecran).TypeAction = TaConsult then Exit;
  AjouteLigne;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/11/2006
Modifié le ... :   /  /    
Description .. : Ajout d'une ligne en fin de tableau avec ajout dans la Tob
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTABLEAUVAR.AjouteLigne (Premiere : boolean = False ; EnDessous : boolean = True);
Var
  LaLigne               : integer;
  TobV                  : Tob;
  i                     : integer;

begin                     
  { FQ 20963 BVE 12.07.07 }
  if HPB.Visible then
  begin
     HPB.OnClose(nil);
     HPB.Visible := false;
  end;
  { END FQ 20963 }
  
  // si insertion dans le tableau
  if FListe.Row < FListe.RowCount -1 then
  begin
    { FQ 19965 BVE 13.06.07 }
    if EnDessous then
       LaLigne             := FListe.Row +1
    else
       LaLigne             := FListe.Row;
    { END FQ 19965 }
  end
  else
    LaLigne             := TobVariation.Detail.Count +1;


  if Not Premiere then
    FListe.RowCount     := FListe.RowCount +1
  else
  begin
    FListe.RowCount     := 2;
    FListe.FixedRows    := 1;
    LaLigne             := 1;
  end;

  FListe.Options        := FListe.Options - [goRowSelect] + [goEditing] + [goTabs];
  FBlocNote.ReadOnly    := False;
  FBlocNote.Clear;

  TobV := Tob.Create('CPTABLEAUVAR', TobVariation, -1);
  TobV.PutValue('CTV_EXERCICE', CodeExercice);
  TobV.PutValue('CTV_GENERAL', CpGeneral);
  TobV.PutValue('CTV_LIGNE', LaLigne);
//  TobV.PutValue('CTV_DATE', iDate1900);
  if StrToDate(DateFinExercice) < V_PGI.DateEntree then
    TobV.PutValue('CTV_DATE', StrToDate(DateFinExercice))
  else
    TobV.PutValue('CTV_DATE', V_PGI.DateEntree);

  TobV.PutValue('CTV_COMMENTAIRE', '');
  TobV.PutValue('CTV_MONTANT', 0);
//  if Premiere then              //FQ20006
//    TobV.PutValue('CTV_MONTANT', Tsolde);
  TobV.PutValue('CTV_AUGMENTATION', 0);
  TobV.PutValue('CTV_DIMINUTION', 0);
  TobV.AddChampSupValeur('CALCULFIN', 0);
  TobV.PutValue('CTV_BLOCNOTE', '');

  ValideLigne(LaLigne, False);

  //si on intercale dans le tableau
  if  LaLigne < FListe.RowCount-1 then
  begin
    FListe.Rows[LaLigne].Clear;
    for i := LaLigne-1 to TobVariation.Detail.count -2 do
      TobVariation.Detail[i].PutValue('CTV_LIGNE', i+2);
    TobVariation.Detail.Sort('CTV_LIGNE');
    for i := LaLigne to FListe.RowCount-1 do
      ValideLigne(i, False);
  end;

  FListe.ColEditables[clAugmentation] := True;
  FListe.ColEditables[clDiminution]   := True;

  FListe.Col            := clDate;
  FListe.Row            := LaLigne;



end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/11/2006
Modifié le ... :   /  /    
Description .. : controle si les donnees de la ligne sont bonnes
Suite ........ : formate les valeurs
Suite ........ : maj de la tob
Suite ........ : appel calcul
Mots clefs ... : 
*****************************************************************}
Function TOF_CPTABLEAUVAR.ValideLigne (LaLigne : integer ; SansCal : boolean = True): boolean;
var
  MtDebut               : double;
  MtAugmentation        : double;
  MtDiminution          : double;
  TobV                  : Tob;
  CalFin                : double;

begin
  Result                := False;

  if SansCal then
  begin
    MtDebut             := Abs(Valeur(FListe.Cells[clDebut, LaLigne]));
    MtAugmentation      := Abs(Valeur(FListe.Cells[clAugmentation, LaLigne]));
    MtDiminution        := Abs(Valeur(FListe.Cells[clDiminution, LaLigne]));
    CalFin              := MtDebut + MtAugmentation - MtDiminution;

    if not IsValidDate(FListe.Cells[clDate, LaLigne]) then
    begin
//      PGIInfo('La date n''est pas valide', 'Tableau des variations');
      FListe.Col        := clDate;
      FListe.Row        := Laligne;
      exit;
    end;

    if StrToDate(FListe.Cells[clDate, LaLigne]) > StrToDate(DateFinExercice) then
    begin
      PGIInfo('La date ne peut être supérieure à celle de l''exercice');
      FListe.Cells[clDate, LaLigne] := '';
      FListe.ForceCellEnter(clDate, LaLigne);
      Exit;
    end;


    Tobv                := TobVariation.Detail[LaLigne-1];
    if Tobv = nil then
      exit;


    TobV.PutValue('CTV_EXERCICE', CodeExercice);
    TobV.PutValue('CTV_GENERAL', CpGeneral);
    TobV.PutValue('CTV_DATE', FListe.Cells[clDate, LaLigne]);
    TobV.PutValue('CTV_MONTANT', MtDebut);
    TobV.PutValue('CTV_AUGMENTATION', MtAugmentation);
    TobV.PutValue('CTV_DIMINUTION', MtDiminution);
    TobV.PutValue('CTV_COMMENTAIRE', FListe.Cells[clCommentaire, LaLigne]);
    Tobv.PutValue('CALCULFIN', CalFin);
(*
    if FBlocNote.LinesDisplayCount <> 0 then
      Tobv.PutValue('CTV_BLOCNOTE', RichToString(FBlocNote))
    else
      Tobv.PutValue('CTV_BLOCNOTE', '');
*)
  end
  else
  begin
    MtDebut             := TobVariation.Detail[LaLigne-1].GetValue('CTV_MONTANT');
    MtAugmentation      := TobVariation.Detail[LaLigne-1].GetValue('CTV_AUGMENTATION');
    MtDiminution        := TobVariation.Detail[LaLigne-1].GetValue('CTV_DIMINUTION');
    CalFin              := MtDebut + MtAugmentation - MtDiminution;
  end;

  FListe.Cells[clDate, LaLigne]           := TobVariation.Detail[LaLigne-1].GetValue('CTV_DATE');
  FListe.Cells[clCommentaire, LaLigne]    := TobVariation.Detail[LaLigne-1].GetValue('CTV_COMMENTAIRE');

  if MtDebut <> 0 then
    FListe.Cells[clDebut, LaLigne]        := STRFMONTANT(MtDebut, 15, V_PGI.OkDecV, '', True)
  else
    FListe.Cells[clDebut, LaLigne]        := '';

  if MtAugmentation <> 0 then
    FListe.Cells[clAugmentation, LaLigne] := STRFMONTANT(MtAugmentation, 15, V_PGI.OkDecV, '', True)
  else
    FListe.Cells[clAugmentation, LaLigne] := '';

  if MtDiminution <> 0 then
    FListe.Cells[ClDiminution, LaLigne]   := STRFMONTANT(MtDiminution, 15, V_PGI.OkDecV, '', True)
  else
    FListe.Cells[clDiminution, LaLigne]   := '';

  if CalFin <> 0 then
    FListe.Cells[clFin, LaLigne]          := STRFMONTANT(CalFin, 15, V_PGI.OkDecV, '', True);

  if TobVariation.Detail[LaLigne-1].GetValue('CTV_BLOCNOTE') <> '' then
    FListe.Cells[clBlocNote, LaLigne]     := '#ICO#7'
  else
    FListe.Cells[clBlocNote, LaLigne]     := '';

  // Erreur si la ligne est vide
  if (FListe.Row=FListe.RowCount-1) and (Fliste.Col=clBlocNote) and ((MtDebut=0) and (MtAugmentation=0) and (MtDiminution=0)) then exit;

  Result := True;
end;

procedure TOF_CPTABLEAUVAR.ValideLigneOnClick(Sender: Tobject);
begin
  EnregTableau (True);
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/11/2006
Modifié le ... :   /  /    
Description .. : Sur l'entree dans une cellule
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTABLEAUVAR.FListeCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
  { BVE 13.06.07
var
  MtAugmentation        : double;
  MtDiminution          : double;
  LaLigne               : integer;

begin
  LaLigne               := FListe.Row;
  MtAugmentation        := Valeur(FListe.Cells[clAugmentation, LaLigne]);
  MtDiminution          := Valeur(FListe.Cells[clDiminution, LaLigne]);
  }
var 
  date   : TDateTime;
begin
  // CA - FQ 20385 - Les 2 colonnes Augmentation et Diminution sont saisissables
  Fliste.ColEditables[clAugmentation]   := True;
  FListe.ColEditables[clDiminution]     := True;

  (*
  if (MtAugmentation = 0) and (MtDiminution = 0) then
  begin
    Fliste.ColEditables[clAugmentation]   := True;
    FListe.ColEditables[clDiminution]     := True;
  end;

  if MtAugmentation <> 0 then
  begin
    Fliste.ColEditables[clAugmentation]   := True;
    FListe.ColEditables[clDiminution]     := False;
  end;

  if MtDiminution <> 0 then
  begin
    Fliste.ColEditables[clAugmentation]   := False;
    FListe.ColEditables[clDiminution]     := True;
  end;
  *)        
  { FQ 20697 BVE 11.07.07 }
  if ( FListe.Col = clDebut ) then
  begin
     date := StrToDate(FListe.Cells[clDate,ARow]);
     if (VH^.EnCours.Deb <= date) and (date <= VH^.EnCours.Fin) then
     begin
        // On passe sur la colonne suivante
        PasseColSuivante(ACol,ARow,FListe);
        Cancel := true; 
     end;
  end;         
  { END FQ 20697 }
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/11/2006
Modifié le ... :   /  /
Description .. : En sortie d'une cellule
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTABLEAUVAR.FListeCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  val    : string;
begin
  If Acol = clCommentaire then
  begin
    If Length(FListe.Cells[clCommentaire, ARow]) > 35 then
      FListe.Cells[clCommentaire, ARow]   := Copy(FListe.Cells[clCommentaire, ARow], 1, 35);
  end;
  { FQ 19993 BVE 15.06.07 }
  if ( ACol = clDate ) then
  begin
     val := FListe.Cells[ACol,ARow];
     if IsNumeric(val) then
     begin
        if Length(val) = 8 then
        begin
           FListe.Cells[ACol,ARow] := Copy(val,1,2) + '/' + Copy(val,3,2) + '/' + Copy(val,5,4);
        end;
     end;
  end;                                                                                          
  { END FQ 19993 }

  if not ValideLigne(ARow) then
  begin
    FListe.Col          := ACol;
    FListe.Row          := ARow;
  end
  else
    CalculTotaux;
end;


procedure TOF_CPTABLEAUVAR.CalculEcart(LaLigne: integer);
Var
  CalEcart              : Double;

begin

  CalEcart              := Valeur(TH34.Caption);

  if CalEcart > 0 then
  begin
    FListe.Cells[clAugmentation, LaLigne]   := '';
    FListe.Cells[clDiminution, LaLigne]     := STRFMONTANT(CalEcart, 15, V_PGI.OkDecV, '', True);;
  end
  else
  begin
    FListe.Cells[clAugmentation, LaLigne]   := STRFMONTANT(Abs(CalEcart), 15, V_PGI.OkDecV, '', True);
    FListe.Cells[clDiminution, LaLigne]     := '';
  end;

end;



{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/11/2006
Modifié le ... :   /  /    
Description .. : Suppression definitive du tableau
Suite ........ : avec delete dans la table
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTABLEAUVAR.DeleteOnClick(Sender: Tobject);
var
  StrSql : String;
  i                     : integer;

begin
  if TobVariation.Detail.Count = 0 then
    exit;

  if PGIAsk('Voulez vous vraiment détruire ce tableau ?', TForm(Ecran).Caption) = mrYes then
  begin
    StrSql              := 'DELETE CPTABLEAUVAR WHERE CTV_EXERCICE="'+CodeExercice+'" AND CTV_GENERAL="'+CpGeneral+'"';
    ExecuteSQL(StrSql);
    TobVariation.Free;
    TobVariation        := Tob.Create('CPTABLEAUVAR', nil, -1);
    for i := FListe.RowCount -1  downto 1 do
    begin
      FListe.DeleteRow(i);
    end;

//    CalculTotaux;
//    FListe.RowCount := 2;
//    FListe.FixedRows := 1;
//    AjouteLigne(True);
    TForm(Ecran).Close;

  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/11/2006
Modifié le ... :   /  /
Description .. : Suppression de la ligne en cours
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTABLEAUVAR.SupprimeLigne;
var
  i                     : integer;

begin
  if TobVariation.Detail.Count = 0 then
    exit;


  if PGIAsk('Voulez vous supprimer la ligne ?', TForm(Ecran).Caption) = mrYes then
  begin
    TobVariation.Detail[FListe.Row -1].Free;
    FListe.DeleteRow(FListe.Row);
    CalculTotaux;


//    if FListe.RowCount = 2 then   //FQ19967
    if TobVariation.Detail.Count = 0 then     // on a tout supprimé
    begin
      AjouteLigne(True);
    end;

    for i := 0 to TobVariation.Detail.Count -1 do
    begin
      TobVariation.Detail[i].PutValue('CTV_LIGNE', i+1);
    end;
    FBLocNote.Clear;
    if TobVariation.Detail.Count -1 >= 0 then    // iln'y a que le descriptif général
      StringtoRich(FBlocNote, TobVariation.Detail[FListe.Row -1].GetValue('CTV_BLOCNOTE'));


  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTABLEAUVAR.SupprimeLigneOnClick(Sender: Tobject);
begin
  // GCO - 02/10/2007 - FQ 21553
  if TFVierge(Ecran).TypeAction = TaConsult then Exit;
  SupprimeLigne;
end;

procedure TOF_CPTABLEAUVAR.CommMillesimeOnClick(Sender: Tobject);
begin
  CommMillesime;
end;

procedure TOF_CPTABLEAUVAR.CommMillesime;
begin
  CPLanceFiche_CPREVDOCTRAVAUX(CpGeneral, '', CodeExercice, 3);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : __/__/____
Modifié le ... : 30/05/2007
Description .. : GCO - 30/05/2007 - FQ 20446
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTABLEAUVAR.CommCompte;
begin
  CPLanceFiche_CPREVDOCTRAVAUX(CpGeneral, '', CodeExercice, 4)
end;

procedure TOF_CPTABLEAUVAR.CommCompteOnClick(Sender: Tobject);
begin
  CommCompte;
end;

procedure TOF_CPTABLEAUVAR.SortieSansSauvegarde;
var
  Rep                   : integer;

begin
  Rep                   := mrNo;
  
  if TobVariation.IsOneModifie(True) then
    Rep                 := PGIAskCancel('Voulez-vous enregistrer les modifications ?', TForm(Ecran).Caption);

  Case Rep of
    mrNo                : TForm(Ecran).Close;
    mrCancel            : Exit;
    mrYes               : EnregTableau (True);
  end;

end;

procedure TOF_CPTABLEAUVAR.SortieSansSauvegardeOnClick(Sender: Tobject);
begin
  SortieSansSauvegarde;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/11/2006
Modifié le ... :   /  /    
Description .. : Recopie les donnees selon combinaison de touche :
Suite ........ : - [F7] : recopie la même zone de la ligne au dessus
Suite ........ : - [Alt]-[F7] : recopie la ligne au dessus
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTABLEAUVAR.Repro(RTotal: Boolean);
Var
  LaCol                 : integer;
  LaLigne               : integer;

begin
  LaLigne               := FListe.Row;

  if not RTotal then
  begin
    LaCol               := FListe.Col;
//    if LaCol = clDate then      //FQ19964
//      exit;
    FListe.Cells[LaCol, LaLigne]          := FListe.Cells[LaCol, LaLigne -1];
  end
  else
  begin
    Fliste.Cells[clDate, LaLigne]         := Fliste.Cells[clDate, LaLigne -1];
    FListe.Cells[clCommentaire, LaLigne]  := FListe.Cells[clCommentaire, LaLigne -1];
    FListe.Cells[clDebut, LaLigne]        := FListe.Cells[clDebut, LaLigne -1];
    FListe.Cells[clAugmentation, LaLigne] := FListe.Cells[clAugmentation, LaLigne -1];
    FListe.Cells[clDiminution, LaLigne]   := FListe.Cells[clDiminution, LaLigne -1];
  end;

  if Valeur(FListe.Cells[clAugmentation, LaLigne]) <> 0 then
  begin
    FListe.ColEditables[clAugmentation]   := True;
    FListe.ColEditables[clDiminution]     := False;
  end
  else
  begin
    FListe.ColEditables[clAugmentation]   := False;
    FListe.ColEditables[clDiminution]     := True;
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/11/2006
Modifié le ... :   /  /    
Description .. : affichage de la fentre du bloc notes
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTABLEAUVAR.BlocnoteOnClick(Sender: Tobject);
Var LaLigne               : integer;
begin
  if TobVariation.Detail.Count = 0 then Exit;

  if Sender = nil then         // si appel depuis alt+B
  begin
    if HPB.visible = False then
    begin
      FBLocNote.Clear;
      StringtoRich(FBlocNote, TobVariation.Detail[FListe.Row-1].GetValue('CTV_BLOCNOTE'));
      HPB.Visible       := True;
      BBlocNote.Down    := True;
      FBlocNote.SetFocus;
    end
    else
    begin
      HPB.Visible       := False;
      BBlocNote.Down    := False;
      FListe.SetFocus;
    end;
  end
  else                         // appuie sur bouton
    begin
    If BBlocNote.Down then
    begin
      FBLocNote.Clear;
      StringtoRich(FBlocNote, TobVariation.Detail[FListe.Row-1].GetValue('CTV_BLOCNOTE'));
      HPB.Visible       := True;
      FBlocNote.SetFocus;
    end
    else
    begin
      HPB.Visible       := False;
      FListe.SetFocus;
    end;
  end;

  if not BBlocNote.Down then
  begin
    LaLigne             := FListe.Row;
    if FBlocNote.LinesDisplayCount <> 0 then
    begin
      TobVariation.Detail[LaLigne-1].PutValue('CTV_BLOCNOTE', RichToString(FBlocNote));
      FListe.Cells[clBlocNote, LaLigne]   := '#ICO#7';
    end
    else
    begin
      TobVariation.Detail[LaLigne-1].PutValue('CTV_BLOCNOTE', '');
      FListe.Cells[clBlocNote, LaLigne]   := '';
    end;
  end;

//  FBLocNote.Clear;
//  StringtoRich(FBlocNote, TobVariation.Detail[FListe.Row-1].GetValue('CTV_BLOCNOTE'));
  
end;

procedure TOF_CPTABLEAUVAR.FListeOnRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  HPB.Visible           := BBlocNote.Down;
  FBLocNote.Clear;
  StringtoRich(FBlocNote, TobVariation.Detail[FListe.Row -1].GetValue('CTV_BLOCNOTE'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/11/2006
Modifié le ... :   /  /    
Description .. : Exportation du tableau avec appel d'une boite de dialogue 
Suite ........ : et choix du format
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTABLEAUVAR.ExportOnClick(Sender: Tobject);
Var
  SauveHint             : String;

begin
  if ASaveDialog.Execute then
  begin
    if ASaveDialog.FilterIndex = 5 then //html
    begin
      SauveHint         := FListe.Hint;
      FListe.Hint       := Ecran.Caption;
      ExportGrid(FListe, nil, ASaveDialog.FileName, ASaveDialog.FilterIndex, True);
      FListe.Hint       := SauveHint;
    end
    else
      ExportGrid(FListe, nil, ASaveDialog.FileName, ASaveDialog.FilterIndex, True);
  end;


end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/11/2006
Modifié le ... : 08/11/2006
Description .. : Sauvegarde du tableau
Suite ........ : procedure :
Suite ........ : 1) delete de la table 
Suite ........ : 2) insert de la tob
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTABLEAUVAR.EnregTableau (bFerme : boolean);
var
  StrSql                : String;
  i                     : integer;

begin
  ValideLigne(FListe.Row);

  Try
    BEGINTRANS;
    StrSql              := 'DELETE CPTABLEAUVAR WHERE CTV_EXERCICE="'+CodeExercice+'" AND CTV_GENERAL="'+CpGeneral+'"';
    ExecuteSQL(StrSql);
    for i := 0 to TobVariation.Detail.Count -1 do
    begin
      // FQ 20367 - CA - On doit enregistrer toutes les lignes du tableau
      // if Valeur(TobVariation.Detail[i].GetValue('CTV_MONTANT')) <> 0 then
      // begin
      TobVariation.Detail[i].SetAllModifie(True);
      TobVariation.Detail[i].InsertOrUpdateDB(False);
      // end;
    end;
    COMMITTRANS;
  Except
    on E: Exception do
    begin
      ROLLBACK;
    end;
  end;
  if bFerme then TForm(Ecran).Close;
end;


procedure TOF_CPTABLEAUVAR.HPBOnClose(Sender: Tobject);
var
  LaLigne               : integer;

begin
  LaLigne               := FListe.Row;
  TobVariation.Detail[LaLigne-1].PutValue('CTV_BLOCNOTE', RichToString(FBlocNote));
  if FBlocNote.LinesDisplayCount <> 0 then
    FListe.Cells[clBlocNote, LaLigne]   := '#ICO#7'
  else
    FListe.Cells[clBlocNote, LaLigne]   := '';
  BBlocNote.Down        := False;
end;

procedure TOF_CPTABLEAUVAR.OnMouseUpFListe(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  ACol                  : integer;
  ARow                  : integer;

begin
  if (Button = mbLeft) then
  begin
    FListe.MouseToCell(X, Y, ACol, ARow);
    if (Acol = 0) and (ARow = 0) then
      ExportGridAutoToExcel(FListe);
  end;
end;

procedure TOF_CPTABLEAUVAR.BimprimeOnCLick(Sender: Tobject);
begin
  ImprimeTableau;
end;

procedure TOF_CPTABLEAUVAR.ImprimeTableau;
begin
  // FQ 20384 - CA -  Proposition d'enregistrement du tableau pour obtenir les valeurs en temps réel
  if TobVariation.IsOneModifie(True) then
  begin
    if (PGIAsk('Voulez-vous enregistrer les modifications ?')=mrYes) then
      EnregTableau (False);
  end;

  TRic.CPLanceEtat_TabVariation( CPGeneral, CodeExercice, 0);
end;

procedure TOF_CPTABLEAUVAR.BlocNoteOnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case key of

      66 :  begin
              if Shift = [ssAlt] then
              begin
                BlocNoteOnCLick(nil);
              end;
            end;
  end;

end;

procedure TOF_CPTABLEAUVAR.ReproEleOnClick(Sender: Tobject);
begin
  Repro(False);
end;

procedure TOF_CPTABLEAUVAR.ReproTotalOnClick(Sender: Tobject);
begin
  Repro(True);
end;

procedure TOF_CPTABLEAUVAR.LeTri(LeType: String);
var
  LatobTri              : Tob;
  TobT                  : Tob;
  TobS                  : Tob;
  LeChamp               : String;
  i                     : integer;

begin
  LaTobTri := Tob.Create('tob detri', nil, -1);
  LaTobTri.Dupliquer(TobVariation, True, True);

  if LaTobTri.Detail.count <1 then
    exit;

  if LeType = 'DATE' then
  begin
    TriLibDesc  := True;
    TriMtDesc   := True;
    LeChamp     := 'CTV_DATE';
    if not TriDateDesc then
    begin
      LeChamp := '-' + LeChamp;
      TriDateDesc := True;
    end
    else
      TriDateDesc := False;
  end

  else if LeType = 'LIB' then
  begin
    TriDateDesc := True;
    TriMtDesc   := True;
    LeChamp     := 'CTV_LIBELLE';
    if not TriLibDesc then
    begin
      LeChamp := '-' + LeChamp;
      TriLibDesc := True;
    end
    else
      TriLibDesc := False;
  end
  else if LeType = 'MT' then
  begin
    TriDateDesc := True;
    TriLibDesc   := True;
    LeChamp     := 'CTV_MONTANT';
    if not TriMtDesc then
    begin
      LeChamp := '-' + LeChamp;
      TriMtDesc := True;
    end
    else
      TriMtDesc := False;
  end;

  LaTobTri.Detail.Sort(LeChamp);

  for i := 0 to LaTobTri.Detail.Count -1 do
  begin
    TobS  := TobVariation.Detail[i];
    TobT  := LaTobTri.Detail[i];
    tobS.PutValue('CTV_DATE', TobT.GetValue('CTV_DATE'));
//    tobS.PutValue('CTV_LIGNE', TobT.GetValue('CTV_LIGNE'));
    tobS.PutValue('CTV_COMMENTAIRE', TobT.GetValue('CTV_COMMENTAIRE'));
    tobS.PutValue('CTV_MONTANT', TobT.GetValue('CTV_MONTANT'));
    tobS.PutValue('CTV_AUGMENTATION', TobT.GetValue('CTV_AUGMENTATION'));
    tobS.PutValue('CTV_DIMINUTION', TobT.GetValue('CTV_DIMINUTION'));
    tobS.PutValue('CTV_BLOCNOTE', TobT.GetValue('CTV_BLOCNOTE'));
    ValideLigne(i+1, False);
  end;
  CalculTotaux;

  

end;

procedure TOF_CPTABLEAUVAR.TridateOnClick(Sender: Tobject);
begin
  LeTri('DATE');

end;

procedure TOF_CPTABLEAUVAR.TriLibelleOnClick(Sender: tobject);
begin
  LeTri('LIB');

end;

procedure TOF_CPTABLEAUVAR.TriMontantOnClick(Sender: Tobject);
begin
  LeTri('MT');

end;

procedure TOF_CPTABLEAUVAR.BlocnoteOnclickF11(Sender: Tobject);
begin
  FBlocNote.Clear;
  StringtoRich(FBlocNote, TobVariation.Detail[FListe.Row-1].GetValue('CTV_BLOCNOTE'));
  HPB.Visible           := True;
  BBlocNote.Down        := True;
end;

procedure TOF_CPTABLEAUVAR.FlisteOnRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  MajBlocNote(Ou);
end;

procedure TOF_CPTABLEAUVAR.MajBlocNote(LaLigne: integer);
begin
  if FBlocNote.LinesDisplayCount <> 0 then
    TobVariation.Detail[LaLigne -1].PutValue('CTV_BLOCNOTE', RichToString(FBlocNote))
  else
    TobVariation.Detail[LaLigne -1].PutValue('CTV_BLOCNOTE', '');

end;

procedure TOF_CPTABLEAUVAR.CycleCommentaireOnClick(Sender: Tobject);
begin
  CPLanceFiche_CPRevDocTravaux( CpGeneral, '', CodeExercice, 0 );
end;

procedure TOF_CPTABLEAUVAR.CycleObjectifOnClick(Sender: Tobject);
begin
  CPLanceFiche_CPRevDocTravaux( CpGeneral, '', CodeExercice, 1 );
end;

procedure TOF_CPTABLEAUVAR.CycleSyntheseOnClick(Sender: Tobject);
begin
  CPLanceFiche_CPRevDocTravaux( CpGeneral, '', CodeExercice, 2 );
end;

Initialization
  registerclasses ( [ TOF_CPTABLEAUVAR ] ) ;
end.
