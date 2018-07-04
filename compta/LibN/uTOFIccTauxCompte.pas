{***********UNITE*************************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 30/08/2001
Modifié le ... : 30/08/2001
Description .. : Source TOF de la TABLE : ICCTAUXCOMPTE ()
Mots clefs ... : TOF;ICCTAUXCOMPTE
*****************************************************************}
Unit uTOFIccTauxCompte ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_MAIN,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTof,
     IccGlobale,
     UTob,        // TOB
     Windows,     // VK_
     Htb97,       // TToolBarButton97
     ParamDat;    // ParamDate

procedure CPLanceFiche_FicheICCTAUX(psz : String);

Type
  TOF_ICCTAUXCOMPTE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

  private
    FCompte    : string;
    FTob       : Tob;
    FValider   : TToolBarButton97;
    GrilleTaux : THGrid;

    procedure OnCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
    procedure OnKeyPress(Sender: TObject; var Key: Char);
    procedure OnElipsisClick(Sender: TObject);
    procedure OnRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure OnKeyDownForm(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure UpdateGrille(Sender : TObject);
    procedure AddRow(Sender : TObject);
    procedure DeleteRow(Sender : TObject);
    procedure ShowLookupList(lGrille : THGrid);

    function  CheckDataOk : Boolean;
    function  IsRowValid(TheRow : integer) : Boolean;

  public

  end ;

const
  LargeurColonne = 80 ;
  ColDate1 = 0 ;
  ColDate2 = 1 ;
  ColTaux  = 2 ;

Implementation

////////////////////////////////////////////////////////////////////////////////
procedure CPLanceFiche_FicheICCTAUX(psz : String);
begin
  AGLLanceFiche('CP', 'ICCTAUXCOMPTE', '', '', psz);
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_ICCTAUXCOMPTE.OnNew ;
begin
  Inherited ;
  AddRow(nil);
end ;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_ICCTAUXCOMPTE.OnDelete ;
begin
  Inherited ;
  DeleteRow(nil);
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 04/09/2001
Modifié le ... : 04/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCTAUXCOMPTE.OnUpdate ;
begin
  Inherited ;

  if not IsRowValid(GrilleTaux.Row) then
    DeleteRow(nil);

  if not CheckDataOk then
  begin
    PgiInfo('Enregistrement impossible : Erreur sur les intervalles des dates.','Vérification des dates');
    Exit;
  end;

  if not BlocageMonoposte(True) then Abort;
  try
    try
      BeginTrans;
      ExecuteSQL('Delete from ICCTauxCompte where ICD_GENERAL = "' + FCompte + '"');

      if FTob.Detail.Count <> 0 then
        FTob.InsertDB(nil,False);

      CommitTrans;
      Ecran.Close;
    except
      on E : Exception do
      begin
        Pgiinfo('Une erreur s''est produite lors de la sauvegarde.' + #10#13 + E.Message,'Enregistrement des taux du compte ' + FCompte);
        RollBack;
      end;
    end;

  finally
    DeblocageMonoPoste(True);
  end;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 04/09/2001
Modifié le ... : 04/09/2001
Description .. : Création de l' affiche
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCTAUXCOMPTE.OnArgument (S : String ) ;
var StArg : String;
begin
  Inherited ;
  StArg := S;
  FCompte := ReadTokenSt(StArg);

  FValider := TToolBarButton97(GetControl('BValider'));

  GrilleTaux := THGrid(GetControl('GrilleTaux')) ;

  GrilleTaux.FixedCols := 0 ;
  GrilleTaux.FixedRows := 1 ;

  Grilletaux.Cells[ColDate1, 0]     := TraduireMemoire('Date du');
  GrilleTaux.ColTypes[ColDate1]     := 'D';
  GrilleTaux.ColFormats[Coldate1]   := ShortDateFormat;
  GrilleTaux.ColAligns[ColDate1]    := TaCenter;
  GrilleTaux.ColWidths[ColDate1]    := LargeurColonne;

  Grilletaux.Cells[ColDate2, 0]     := TraduireMemoire('Date au');
  GrilleTaux.ColTypes[ColDate2]     := 'D';
  GrilleTaux.ColFormats[Coldate2]   := ShortDateFormat;
  GrilleTaux.ColAligns[ColDate2]    := TaCenter;
  GrilleTaux.ColWidths[ColDate2]    := LargeurColonne;

  GrilleTaux.Cells[ColTaux, 0]      := TraduireMemoire('Taux en %');
  GrilleTaux.ColAligns[ColTaux]     := TaRightJustify;
  GrilleTaux.ColTypes[ColTaux]      := 'R'; // GCO - 06/11/2007 - FQ 21781
  GrilleTaux.ColFormats[ColTaux]    := StrfMask(V_PGI.OkDecV, '', True);
  GrilleTaux.ColWidths[ColTaux]     := LargeurColonne;

  // Evénements de la grille
  GrilleTaux.OnKeyPress             := OnKeyPress;
  GrilleTaux.OnElipsisClick         := OnElipsisClick;
  GrilleTaux.OnCellEnter            := OnCellEnter;
  GrilleTaux.OnRowExit              := OnRowExit;
  GrilleTaux.OnCellExit             := OnCellExit;

  Ecran.OnKeyDown := OnKeyDownForm;

  FTob := Tob.Create('',nil,-1);
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 04/09/2001
Modifié le ... : 04/09/2001
Description .. : Chargement de l' enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCTAUXCOMPTE.OnLoad ;
var SQL : TQuery;
begin
  Inherited ;

  SQL := nil;
  try
    SQL := OpenSQL('Select ICD_DATEDU, ICD_DATEAU, ICD_TAUX from IccTauxCompte where ICD_GENERAL = "' + FCompte + '"',True);

    if not SQL.EOF then
      FTob.LoadDetailDB('ICCTAUX', '', '', SQL, True, False);

    UpdateGrille(nil);

  finally
    if Assigned(SQL) then Ferme(SQL);
  end;

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 03/09/2001
Modifié le ... : 03/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCTAUXCOMPTE.OnKeyDownForm(Sender: TObject; var Key: Word;Shift: TShiftState);
begin
  case Key of

    VK_F5 : if (Ecran.ActiveControl = GrilleTaux) and
               ((GrilleTaux.Col = ColDate1) or (GrilleTaux.Col = ColDate2)) then
              ShowLookupList(GrilleTaux);

    VK_F7 : if GrilleTaux.Row > 1 then
              GrilleTaux.Cells[GrilleTaux.Col, GrilleTaux.Row] := GrilleTaux.Cells[GrilleTaux.Col, GrilleTaux.Row - 1];

    VK_F10 : TToolBarButton97(GetControl('BVALIDER')).Click;

    VK_INSERT : AddRow(nil);

    VK_DELETE : if Shift = [ssCtrl] then DeleteRow(nil);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 30/08/2001
Modifié le ... : 30/08/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCTAUXCOMPTE.OnCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  GrilleTaux.ElipsisButton := (GrilleTaux.Col = ColDate1) or (GrilleTaux.Col = ColDate2);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 03/09/2001
Modifié le ... : 03/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCTAUXCOMPTE.OnElipsisClick(Sender: TObject);
begin
  V_PGI.ParamDateproc(GrilleTaux);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 03/09/2001
Modifié le ... : 03/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_ICCTAUXCOMPTE.IsRowValid(TheRow: integer): Boolean;
begin
  Result := False;
  if (Valeur(GrilleTaux.Cells[ColTaux, TheRow]) > 0) and
     (Valeur(GrilleTaux.Cells[ColTaux, TheRow]) < 100) then
   Result := True;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 03/09/2001
Modifié le ... : 03/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCTAUXCOMPTE.OnKeyPress(Sender: TObject; var Key: Char);
begin
  // GCO - 06/11/2007 - FQ 21781
  if key = #9 then Exit;

  if (GrilleTaux.Col = ColDate1) or (GrilleTaux.Col = ColDate2) then
    ParamDate(Ecran, Sender, Key);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 03/09/2001
Modifié le ... : 03/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCTAUXCOMPTE.OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  case ACol of

    ColDate1 :
      begin
        if not IsValidDate(GrilleTaux.Cells[ACol, ARow]) then
        begin
          Cancel := True;
          Exit;
        end;
      end;

    ColDate2 :
      begin
        if not IsValidDate(GrilleTaux.Cells[ACol, ARow]) then
        begin
          Cancel := True;
          Exit;
        end
        else
          if StrToDate(GrilleTaux.Cells[ColDate2,ARow]) <=
             StrToDate(GrilleTaux.Cells[ColDate1,ARow]) then
          begin
            PgiInfo('La date de fin doit être supérieure à la date de début.','Saisie de date');
            Cancel := True;
            Exit;
          end;
      end;

    ColTaux:
      begin
        if not IsNumeric(GrilleTaux.Cells[ColTaux, ARow]) then
        begin
          Cancel := True;
          Exit;
        end
        else
        begin
          if (Valeur(GrilleTaux.Cells[ColTaux, ARow]) > 0) and
             (Valeur(GrilleTaux.Cells[ColTaux, ARow]) < 100) then
            GrilleTaux.Cells[ACol, ARow] := StrfMontant(Valeur(GrilleTaux.Cells[ACol, ARow]), 20, 2, '', True)
          else
          begin
            PgiInfo('Le taux de rémunération doit être compris entre 0 et 100.','Saisie du taux');
            Cancel := True;
            Exit;
          end;

          if (ARow = GrilleTaux.RowCount - 1) then AddRow(nil);
        end;
      end;

  end; // End du Case
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 03/09/2001
Modifié le ... : 03/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCTAUXCOMPTE.OnRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  Cancel := not IsRowValid(Ou);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 03/09/2001
Modifié le ... : 03/09/2001
Description .. :
Mots clefs ... : Charge la grille avec les enregistrements de FTOB
*****************************************************************}
procedure TOF_ICCTAUXCOMPTE.UpdateGrille(Sender: TObject);
begin
  FTob.PutGridDetail(GrilleTaux, False, False, 'ICD_DATEDU; ICD_DATEAU; ICD_TAUX');

  if FTob.Detail.Count = 0 then
    GrilleTaux.RowCount := 2
  else
    GrilleTaux.RowCount := FTob.Detail.Count + 1;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 03/09/2001
Modifié le ... : 03/09/2001
Description .. : Ajouter une ligne dans la grille des taux
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCTAUXCOMPTE.AddRow(Sender: TObject);
begin
  if not isRowValid(GrilleTaux.Row) then Exit;

  GrilleTaux.EditorMode := False;
  GrilleTaux.RowCount := GrilleTaux.RowCount + 1;
  GrilleTaux.Row := GrilleTaux.RowCount - 1;
  GrilleTaux.Col := ColDate1;
  GrilleTaux.EditorMode := True;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 03/09/2001
Modifié le ... : 03/09/2001
Description .. : Efface une ligne dans la grille des taux
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCTAUXCOMPTE.DeleteRow(Sender: TObject);
begin
  if GrilleTaux.RowCount = 2 then
  begin
    GrilleTaux.Cells[ColDate1, GrilleTAux.Row] := '';
    GrilleTaux.Cells[ColDate2, GrilleTAux.Row] := '';
    GrilleTaux.Cells[ColTaux, GrilleTAux.Row] := '';
    GrilleTaux.Col := ColDate1;
    GrilleTaux.EditorMode := True;
  end
  else
    GrilleTaux.DeleteRow(GrilleTaux.Row);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 04/09/2001
Modifié le ... : 04/09/2001
Description .. : Affiche l'écran de sasie de date avec F5
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCTAUXCOMPTE.ShowLookupList(lGrille: THGrid);
begin
  lGrille.EditorMode := True;
  V_PGI.ParamDateproc(lGrille);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 04/09/2001
Modifié le ... : 04/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_ICCTAUXCOMPTE.CheckDataOk : Boolean;
var lTob : Tob;
    i : integer;
begin
  FTob.ClearDetail;

  if (not IsRowValid(GrilleTaux.Row)) and ( GrilleTaux.RowCount = 2) then
  begin
    //  Les données sont bonnes mais sont vides
    Result := True;
    Exit;
  end
  else
  begin
    for i := 1 to (GrilleTaux.RowCount-1) do
    begin
      lTob := Tob.Create('ICCTAUXCOMPTE', FTob,-1);
      lTob.PutValue('ICD_GENERAL', FCompte);
      lTob.PutValue('ICD_DATEDU',  StrToDate(GrilleTaux.Cells[ColDate1, i]));
      lTob.PutValue('ICD_DATEAU',  StrToDate(GrilleTaux.Cells[ColDate2, i]));
      lTob.PutValue('ICD_TAUX',    GrilleTaux.Cells[ColTaux , i]);
    end;

    FTob.Detail.Sort('ICD_DATEDU');
    for i := 0 to FTob.Detail.Count -1 do
    begin
      if (i <> 0) and (FTob.Detail[i].GetValue('ICD_DATEDU') <= FTob.Detail[i-1].GetValue('ICD_DATEAU')) then
      begin
        FTob.ClearDetail;
        Result := False;
        Exit;
      end;
    end;

    // L' ordre de saisie des dates est correctes
    Result := True;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 03/09/2001
Modifié le ... : 03/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCTAUXCOMPTE.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_ICCTAUXCOMPTE ] ) ;
end.

