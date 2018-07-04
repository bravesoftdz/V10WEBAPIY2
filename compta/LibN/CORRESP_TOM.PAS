{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 24/07/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CORRESP (CORRESP)
Mots clefs ... : TOM;CORRESP
*****************************************************************}
Unit CORRESP_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
     Dialogs,
{$IFNDEF EAGLCLIENT}
     db,
     HDB,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,
     Fiche,
     FichList,
     FichGrid,
     Mul,
{$else}
     MaineAGL,
     eFiche,
     eFichList,
     eFichGrid,
     eMul,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     Ent1,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM, 
     UTob  ,UentCommun ;
     
Const	MsgBox: array[1..12] of string = (
		{1} 'des comptes généraux',
		{2}	'des comptes auxiliaires',
		{3}	'de l''axe analytique n° 1',
		{4}	'de l''axe analytique n° 2',
		{5}	'de l''axe analytique n° 3',
		{6}	'de l''axe analytique n° 4',
		{7}	'de l''axe analytique n° 5',
		{8}	'des comptes budgétaires',
    {9} 'Journaux',
    {10}'Mode de paiement',
    {11}'Paramétrage SISCOII',
    {12}'Etablissement');

procedure CCLanceFiche_Correspondance(pszType : String);
procedure ZoomCorresp(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);

Type
  TOM_CORRESP = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    procedure OnLoadAlerte               ; override ;
    procedure BValiderClick(Sender: TObject);
  private
    gszFType          : string ;
    gszNoType         : string ;
    FFindFirst        : Boolean;
  {$IFDEF EAGLCLIENT}
    FListe            : THGrid ;
  {$ELSE}
    FListe            : THDBGrid ;
  {$ENDIF}
    FMouseX, FMouseY  : Integer;
    rbtPlan1          : TRadioButton;
    rbtPlan2          : TRadioButton;
    AFindDialog       : TFindDialog;
    ValiderClick      : TNotifyEvent;

    // Initialisation :
    procedure ChargeControl;
    procedure ChargeEvenements;
    procedure ChargeTitreEcran;
    procedure InitialiseGrille(OkInitGrid : Boolean = TRUE);
    // Evenements :
    procedure PlanClick(Sender : TObject);
    procedure ExportClick(Sender : TObject);
    procedure ImportClick(Sender : TObject);
    procedure ListeKeyPress(Sender: TObject; var Key: Char);
    procedure ListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ListeDblClick(Sender: TObject);
    procedure ListeMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
{$IFDEF EAGLCLIENT}
    procedure DoCollExit(Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
{$ELSE}
    procedure DoCollExit(Sender: TObject);
{$ENDIF}                                 
    procedure OnClickBRechercher(Sender: Tobject);
    procedure OnFindAFindDialog(Sender : TObject);

    // Autres
    procedure ControlChamp;
    procedure MajTitreColGrille;
    function  GetNumPlan : Char;
    function  ExisteLien(Compte : String ; Plan : string ; var Mess : string) : boolean;
    function  VerifParamSoc : Boolean;
    function  IsCompensation: Boolean;
    end ;

Implementation


uses
{$IFDEF MODENT1}
  CPObjetGen,
  CPTypeCons,
  CPProcMetier,
{$ENDIF MODENT1}
{$IFDEF EAGLCLIENT}
     Grids,
{$ELSE}
     DBGrids,
{$ENDIF}
     HTB97,
     HStatus,
     HSysMenu,
     Hxlspas,
     LookUp,
     ParamSoc,
     uLibEcriture,
     UtilSoc,
     Windows, TntDBGrids;

procedure CCLanceFiche_Correspondance(pszType : String);
begin
  AGLLanceFiche('CP','CPCORRESP','','',pszType);
end;

procedure ZoomCorresp(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
begin
  AGLLanceFiche('CP','CPCORRESP','','',Axe);
end;

procedure TOM_CORRESP.OnArgument ( S: String ) ;
begin
  Inherited ;
  // On sauvegarde le type des correspondances.
  gszFType := S;

  if (gszFType[1] = 'I') or (gszFType = 'SIS') then
  begin
{$IFDEF EAGLCLIENT}
     TFFicheGrid(Ecran).FRange := gszFType;
{$ELSE}
     DS.Filtered := True;
     DS.Filter := 'CR_TYPE=''' + gszFType + '''';
{$ENDIF}
     gszNoType := '';
     TRadioButton(GetControl('RBTN_PLAN')).visible := FALSE;
     TRadioButton(GetControl('RBTN_PLAN_')).visible := FALSE;
  end
  else
  begin
     gszNoType := GetNumPlan;
{$IFDEF EAGLCLIENT}
     TFFicheGrid(Ecran).FRange := gszFType + gszNoType;
{$ELSE}
     DS.Filtered := True;
     DS.Filter := 'CR_TYPE=''' + gszFType + gszNoType + '''';
{$ENDIF}
  end;
  // Ajout me
  TToolBarButton97(GetControl('BEXPORT')).ShowHint := TRUE;
  TToolBarButton97(GetControl('BEXPORT')).Hint := 'Exporter';

  TToolBarButton97(GetControl('BIMPORT')).ShowHint := TRUE;
  TToolBarButton97(GetControl('BIMPORT')).Hint := 'Importer';

  // Initialisation
  ChargeControl;
  ChargeEvenements;
  ChargeTitreEcran;
  // AJOUT ME plantage dans Comsx MajTitreColGrille;
  InitialiseGrille(FALSE);

  if not VerifParamSoc then
  begin
     PgIInfo ('Attention, vous devez activer au moins un plan de correspondance dans Paramètres Société - Comptables');
     SetControlVisible('BINSERT', FALSE);
     ParamSociete(False, BrancheParamSocAVirer, 'SCO_COMPTABLES', '',
                  RechargeParamSoc, ChargePageSoc, SauvePageSoc, InterfaceSoc, 1105000);
     { FQ 21292 BVE 29.08.07
     Self.OnClose;
     { END FQ 21292 }
  end;
end ;

procedure TOM_CORRESP.OnNewRecord ;
begin
  Inherited ;
  // Renseigne le champ CR_TYPE automatiquement
  SetField('CR_TYPE', gszFType + gszNoType);
  SetControlProperty('CR_CORRESP', 'EDITMASK', '>aaaaaaaaaaaaaaaaa;0;');
end ;

procedure TOM_CORRESP.OnDeleteRecord ;
var
  SQL     : string;
  Corresp : string;
  Pl      : string;
  Mess    : string;
begin
  if (gszFType='AU') and TCompensation.IsCompensation then
  begin
     SQL := 'UPDATE TIERS SET T_CORRESP' + gszNoType + ' = "" '+
            'WHERE T_AUXILIAIRE = "' + GetField('CR_LIBRETEXTE1') + '"'+
            '   OR T_AUXILIAIRE = "' + GetField('CR_LIBRETEXTE2') + '"';
     ExecuteSQL(SQL);
  end;
  // ajout me 20028
  Pl := '';
  if GetControlVisible('RBTN_PLAN') then
  begin
        if (rbtPlan1 <> nil) and rbtPlan1.Checked then
           Pl := '1'
        else if (rbtPlan2 <> nil) and rbtPlan2.Checked then
           Pl := '2';
  end
  else
     Pl := gszNoType;

  Corresp := GetField('CR_CORRESP');
  if (pl = '') or (Corresp = '') then Exit;
  if ExisteLien(Corresp,Pl,Mess) then
  begin             
     LastErrorMsg := Mess;
     LastError := 1;
     Exit;
  end;

  inherited;
end ;

procedure TOM_CORRESP.OnUpdateRecord ;
  function VerifCompte(NatureCompte, Compte: String): Boolean;
  var
    SQL:          String;
    Ok:           Boolean;
    LibelNature:  String;
    Q:            TQuery;
  begin
    Ok := True;

    if NatureCompte = 'CLI' then
      LibelNature := 'client'
    else
      LibelNature := 'fournisseur';

    if Ok and (Trim(Compte) = '') then
      begin
      PgiBox(traduireMemoire('Le compte ' + LibelNature + ' n''est pas renseigné.')) ;
      Ok := False;
      end ;

    SQL := 'SELECT CR_TYPE,CR_LIBRETEXTE2,CR_CORRESP'+
           ' FROM CORRESP'+
           ' WHERE CR_TYPE="'+GetField('CR_TYPE')+'"'+
           '   AND CR_CORRESP <> "'+GetField('CR_CORRESP')+'"'+
           '   AND '+TCompensation.GetSQLCorrespCompte(NatureCompte, Compte);
    Q := OpenSQL(SQL, True);
    if ok and not Q.Eof then
      begin
      PgiBox(traduireMemoire('Le compte ' + LibelNature + ' ' + Compte + ' est déjà associé au compte ' + Q.FindField('CR_LIBRETEXTE2').AsString + ' sous le compte de correspondance ' + Q.FindField('CR_CORRESP').AsString + '.')) ;
      Ok := False;
      end ;
    Ferme(Q);

    if Ok then
      begin
      SQL := 'SELECT T_AUXILIAIRE, T_PAYEUR'+
             ' FROM TIERS'+
             ' WHERE T_AUXILIAIRE="'+Compte+'"';
      if NatureCompte = 'CLI' then
        SQL := SQL + '   AND (T_NATUREAUXI="CLI" OR T_NATUREAUXI="AUD")'
      else
        SQL := SQL + '   AND (T_NATUREAUXI="FOU" OR T_NATUREAUXI="AUC")';
      Q := OpenSQL(SQL, True);
      if Q.Eof then
        begin
        PgiBox(TraduireMemoire('Le compte '+Compte+' n''existe pas ou n''est pas un '+LibelNature+'.')) ;
        Ok := False;
        end;
      Ferme(Q);
      end;
    Result := Ok;
  end;
var
  SQL: String;
  Ok:  Boolean;
begin
{$IFNDEF EAGLCLIENT}
  if Trim(FListe.Fields[0].AsString) = '' then
{$ELSE}
  if Trim(FListe.Cells[0,FListe.Row]) = '' then
{$ENDIF}
  begin
     LastError := 1;
     LastErrorMsg := TraduireMemoire('Vous devez renseignez un code.');
     Exit;
  end;
  Ok := True;
  ControlChamp;
  if IsCompensation then
  begin
    if GetField('CR_LIBRETEXTE1') = GetField('CR_LIBRETEXTE2') then
    begin
      PgiBox(traduireMemoire('Les comptes client et fournisseur doivent être différents.')) ;
      Ok := False;
    end;
    Ok := Ok and VerifCompte('CLI', GetField('CR_LIBRETEXTE1'));
    Ok := Ok and VerifCompte('FOU', GetField('CR_LIBRETEXTE2'));
    if Ok then
    begin
      {b FP 13/07/2006 FQ18589: Efface les anciens comptes ayant le plan de correspondance}
      SQL := 'UPDATE TIERS SET T_CORRESP'+gszNoType+'=""'+
             ' WHERE T_CORRESP'+gszNoType+'="'+GetField('CR_CORRESP')+'"';
      ExecuteSQL(SQL);
      {e FP 13/07/2006}
      SQL := 'UPDATE TIERS SET T_CORRESP'+gszNoType+'="'+GetField('CR_CORRESP')+'"'+
             ' WHERE T_AUXILIAIRE="'+GetField('CR_LIBRETEXTE1')+'"'+
             '    OR T_AUXILIAIRE="'+GetField('CR_LIBRETEXTE2')+'"';
      ExecuteSQL(SQL);
    end;
  end;

  if Ok then OK := VerifParamSoc;

  if Ok then
    Inherited
  else
    LastError := -1;
end ;

procedure TOM_CORRESP.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_CORRESP.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_CORRESP.OnLoadRecord ;
begin
  Inherited ;
  MajTitreColGrille;
end ;

procedure TOM_CORRESP.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_CORRESP.OnClose ;
begin
  Inherited ;
  FreeAndNil(AFindDialog);
end ;

procedure TOM_CORRESP.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_CORRESP.OnLoadAlerte;
begin
  Inherited ;
end ;

procedure TOM_CORRESP.BValiderClick(Sender: TObject);
var
  Lg       : integer;
  NotOK    : Boolean;
  Q        : TQuery;
begin
  NOTOk := FALSE;
  if (gszFType = 'IGE') or (gszFType = 'IAU') then
  begin
     if (gszFType = 'IAU') then
        lg :=  GetInfoCpta (fbAux).Lg
     else
        lg :=  GetInfoCpta (fbGene).Lg;
{$IFDEF EAGLCLIENT}
     Q := TFMul(Ecran).Q.TQ;
{$ELSE}
     Q := TFMul(Ecran).Q;
{$ENDIF}
     Q.First;
     while not Q.EOF do
     begin
        if Length(Q.FindField('CR_CORRESP').AsString ) <> lg then
        begin
           NOTOk := TRUE;
           Break;
        end;
        Q.Next;
     end;
  end;
  if NOTOk then
  begin
     PgiInfo ('La longueur des comptes entrants  est différente de celle définie dans le dossier. '+
      #10#13+ ' Vous pouvez utiliser des caractères jockers % ou $');
     Exit;
  end;
  ValiderClick(Sender);
end;

procedure TOM_CORRESP.PlanClick(Sender : TObject);
begin
  { FQ 20948 BVE 18.10.07 }
  if (TFFicheGrid(Ecran).TypeAction = taModif) or (TFFicheGrid(Ecran).TypeAction = taCreat) then
     BValiderClick(nil);
  { END FQ 20948 }
  if Sender = nil then Exit
  else if TRadioButton(Sender).Name = 'RBTN_PLAN' then
     gszNoType := '1'
  else if TRadioButton(Sender).Name = 'RBTN_PLAN_' then
     gszNoType := '2'
  else
     Exit;
  if TFFicheGrid(Ecran).TypeAction = taModif then
     TToolBarButton97(GetControl('BDefaire')).OnClick(nil);
  InitialiseGrille ;
{$IFDEF EAGLCLIENT}
  TFFicheGrid(Ecran).FRange := gszFType + gszNoType;
  TFFicheGrid(Ecran).ReloadDB;
{$ELSE}
  /// Affiche l'autre plan de correspondance 
  DS.Filtered := True;
  DS.Filter   := 'CR_TYPE=''' + gszFType + gszNoType + '''';
  DS.Refresh;
{$ENDIF}
end;

procedure TOM_CORRESP.ExportClick(Sender: TObject);
var SD : THSaveDialog;
begin
  SD := THSaveDialog.Create(Application);
  SD.DefaultExt := '*.txt';
  SD.Title := 'Exportation de fichiers';
  SD.Filter := 'Fichier Texte (*.txt)|*.txt';
  if SD.execute then
     ExportGrid(Nil, FListe, SD.FileName, SD.FilterIndex, TRUE);
  SD.free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Compta
Créé le ...... : 14/08/2007
Modifié le ... : 14/08/2007
Description .. :
Suite ........ : SBO 14/08/2007 : Traitement des problèmes d'import des 
Suite ........ : comptes de correspondance FQ 19074
Mots clefs ... : 
*****************************************************************}
procedure TOM_CORRESP.ImportClick(Sender: TObject);
var SD    : TSaveDialog;
    F     : TextFile ;
    Cod,Lib,St : String ;
    car        : string;
    lTobCor    : Tob ;
    lBoResult  : boolean ;
    abreg      : string;
begin
  SD := TSaveDialog.Create(Application);
  SD.DefaultExt := '*.txt';
  SD.Title := 'Import de fichiers';
  SD.Filter := 'Fichier Texte (*.txt)|*.txt|Fichier Ascii (*.csv)|*.csv';

  lTobCor    := Tob.Create('CORRESP', nil, -1) ;
  lBoResult  := False ;

  try // finally

    try // except

      Begintrans ;

      if SD.execute then
      begin
          AssignFile(F,  SD.FileName) ;
         {$i-} Reset(F,  SD.FileName) ; {$i+}
          If Ioresult<>0 Then
            BEGIN
            Exit ;
            END ;
          InitMove(100,'') ;
         if ExisteSQL('SELECT * FROM CORRESP WHERE CR_TYPE="' + gszFType + gszNoType + '"') then
         begin
              if PGIAsk('Voulez-vous Supprimez le plan de correspondance déjà existant ?','Import')=mrYes then
                ExecuteSQL('DELETE FROM CORRESP WHERE CR_TYPE="' + gszFType + gszNoType + '"') ;
         end;
          While Not Eof(F) Do
            BEGIN
            Readln(F,St) ;
            if pos('Compte de correspondance', St) <> 0 then Continue;
            if pos('Compte dans CEGID Comptabilité', St) <> 0 then Continue;
            St:=Trim(St) ;
            MoveCur(FALSE) ;
            if pos(';', St) <> 0 then Car := ';';
            if pos(#9, St) <> 0 then Car := #9;

            If St<>'' Then
              BEGIN
              Cod:='' ; Lib:='' ;
              If St[Length(St)]<>CAr Then St:=St+Car ;
              Cod := ReadTokenPipe(St, Car) ;
              if (gszFType='IGE') and (pos('%', Cod) =0)  and (pos('$', Cod) =0) then Cod:=BourreOuTronque(Cod,fbGene) else
              if (gszFType='IAU') and (pos('%', Cod) =0)  and (pos('$', Cod) =0) then Cod:= BourreOuTronque(Cod,fbAux) ;
              If St<>'' Then Lib        := ReadTokenPipe(St, Car) ;
              If St<>'' Then abreg      := ReadTokenPipe(St, Car) ;
              If Lib<>'' Then
                begin
                if ExisteSQL('SELECT CR_CORRESP FROM CORRESP WHERE CR_TYPE="' + gszFType + gszNoType + '" AND CR_CORRESP="'+Cod+'" ') then
                  ExecuteSQL('UPDATE CORRESP SET CR_LIBELLE = "' + Lib + '" WHERE CR_TYPE="' + gszFType + gszNoType + '" AND CR_CORRESP="'+Cod+'"  ' )
                else
                  begin
                  lTobCor.PutValue('CR_TYPE',    gszFType + gszNoType ) ;
                  lTobCor.PutValue('CR_CORRESP', Cod ) ;
                  lTobCor.PutValue('CR_LIBELLE', Lib ) ;
                  lTobCor.PutValue('CR_ABREGE', abreg ) ;
                  lTObCor.InsertDB( nil ) ;
                  end ;
                end ;
              END ;
            END ;
          CloseFile(F) ;
          FiniMove ;

      end;

    CommitTrans ;
    lBoResult := True ;

    except

      on E: Exception do
        begin
        rollback ;
        PgiInfo( TraduireMemoire('Une erreur est survenue pendant l''import : ') + E.Message ) ;
        lBoResult := False ;
        end;

      end ;

  finally
    SD.free;
    FreeAndNil( lTobCor ) ;

    if lBoResult then
      begin
      PgiInfo( TraduireMemoire('L''import s''est terminé avec succès. Le module va être fermé pour mettre à jour les données.') ) ;
      TFFicheGrid(Ecran).bDefaireClick(nil) ; //CancelRecord ;
      TFFicheGrid(Ecran).BFermeClick(nil);
      end ;
  end ;

end;

procedure TOM_CORRESP.ControlChamp;
begin
{$IFDEF EAGLCLIENT}
  if (gszFType='IGE') or (gszFType='ZG') then FListe.Cells[1, FListe.Row]:= BourreOuTronque(FListe.Cells[1, FListe.Row], fbGene)
  else if (gszFType='IAU') then FListe.Cells[1, FListe.Row] := BourreOuTronque(FListe.Cells[1, FListe.Row], fbAux) // des comptes auxiliaires
  else if (gszFType='IA1') then FListe.Cells[1, FListe.Row] := BourreOuTronque(FListe.Cells[1, FListe.Row], fbAxe1) // de l'axe analytique n° 1
  else if (gszFType='IA2') then FListe.Cells[1, FListe.Row] := BourreOuTronque(FListe.Cells[1, FListe.Row], fbAxe2) // de l'axe analytique n° 2
  else if (gszFType='IA3') then FListe.Cells[1, FListe.Row] := BourreOuTronque(FListe.Cells[1, FListe.Row], fbAxe3) // de l'axe analytique n° 3
  else if (gszFType='IA4') then FListe.Cells[1, FListe.Row] := BourreOuTronque(FListe.Cells[1, FListe.Row], fbAxe4) // de l'axe analytique n° 4
  else if (gszFType='IA5') then FListe.Cells[1, FListe.Row] := BourreOuTronque(FListe.Cells[1, FListe.Row], fbAxe5) // de l'axe analytique n° 5
  else if (gszFType='BU') then FListe.Cells[1, FListe.Row] := BourreOuTronque(FListe.Cells[1, FListe.Row], fbBudGen) // des comptes budgétaires
  else if (length(FListe.Cells[1, FListe.Row]) > 3 ) and ((gszFType='IJA') or (gszFType='IPM') or (gszFType='IET')) then
  begin
     PgiBox(traduireMemoire('Longueur maximum : 3 .')) ;
     FListe.Cells[1, FListe.Row] := Copy(FListe.Cells[1, FListe.Row], 1, 3);
  end
  else
     if (gszFType='SIS') then
     begin
        FListe.Cells[1, FListe.Row] := BourreOuTronque(FListe.Cells[1, FListe.Row], fbAux);
        FListe.Cells[2, FListe.Row] := BourreOuTronque(FListe.Cells[1, FListe.Row], fbAux);
     end;
{$ELSE}
  // Ajout me
  if not(DS.State in [dsInsert]) then exit;

  if (gszFType='IGE') or (gszFType='ZG') then FListe.Fields[1].AsString := BourreOuTronque(FListe.Fields[1].AsString, fbGene)
  else if (gszFType='IAU') then FListe.Fields[1].AsString := BourreOuTronque(FListe.Fields[1].AsString, fbAux) // des comptes auxiliaires
  else if (gszFType='IA1') then FListe.Fields[1].AsString := BourreOuTronque(FListe.Fields[1].AsString, fbAxe1) // de l'axe analytique n° 1
  else if (gszFType='IA2') then FListe.Fields[1].AsString := BourreOuTronque(FListe.Fields[1].AsString, fbAxe2) // de l'axe analytique n° 2
  else if (gszFType='IA3') then FListe.Fields[1].AsString := BourreOuTronque(FListe.Fields[1].AsString, fbAxe3) // de l'axe analytique n° 3
  else if (gszFType='IA4') then FListe.Fields[1].AsString := BourreOuTronque(FListe.Fields[1].AsString, fbAxe4) // de l'axe analytique n° 4
  else if (gszFType='IA5') then FListe.Fields[1].AsString := BourreOuTronque(FListe.Fields[1].AsString, fbAxe5) // de l'axe analytique n° 5
  else if (gszFType='BU') then FListe.Fields[1].AsString := BourreOuTronque(FListe.Fields[1].AsString, fbBudGen) // des comptes budgétaires
  else if (length(FListe.Fields[1].AsString) > 3 ) and ((gszFType='IJA') or (gszFType='IPM') or (gszFType='IET')) then
  begin
     PgiBox(traduireMemoire('Longueur maximum : 3 .')) ;
     FListe.Fields[1].AsString := Copy(FListe.Fields[1].AsString, 1, 3);
  end
  else
     if (gszFType='SIS') then
     begin
        FListe.Fields[1].AsString := BourreOuTronque(FListe.Fields[1].AsString, fbAux);
        FListe.Fields[2].AsString := BourreOuTronque(FListe.Fields[2].AsString, fbAux);
     end;
{$ENDIF}
end;

procedure TOM_CORRESP.ListeDblClick(Sender: TObject);
var
  iCol:        Integer;
  TempControl: THEdit;
  Ok:          Boolean;
  lStWhere,lStColonne,lStOrder,lStSelect: String;
begin
  Ok := False;
  {$IFDEF EAGLCLIENT}
  iCol := FListe.Col;
  {$ELSE}
  iCol := FListe.SelectedIndex;
  {$ENDIF}

  if IsCompensation and (iCol in [2, 3]) then    {FP 25/04/2006 FQ17907}
    begin
    TempControl := THEdit.Create(Ecran);
    try
      TempControl.Visible := False;
      TempControl.Parent  := Ecran;
      TempControl.Top     := FMouseY;
      TempControl.Left    := FMouseX;
      {$IFDEF EAGLCLIENT}
      TempControl.Text := FListe.Cells[iCol, FListe.Row];
      {$ELSE}
      TempControl.Text := FListe.InplaceEditor.Text;      {FP 25/04/2006 FQ17907}
      {$ENDIF}

      if (iCol  = 2) then          {Compte Client}
        begin
         CMakeSQLLookupAux(lStWhere,lStColonne,lStOrder,lStSelect,'','COC');
         lStWhere:=lStWhere+' AND T_NATUREAUXI <> "DIV"';
         Ok := LookupList(TempControl,TraduireMemoire('Comptes clients'),'TIERS',lStColonne,lStSelect,lStWhere,lStOrder,true, 0);
        end
      else if (iCol  = 3) then     {Compte Fournisseur}
        begin
         CMakeSQLLookupAux(lStWhere,lStColonne,lStOrder,lStSelect,'','COF');
         lStWhere:=lStWhere+' AND T_NATUREAUXI <> "DIV"';
         Ok := LookupList(TempControl,TraduireMemoire('Comptes fournisseurs'),'TIERS',lStColonne,lStSelect,lStWhere,lStOrder,true, 0);
        end;
      if Ok then
        begin
        {$IFDEF EAGLCLIENT}
        FListe.Cells[iCol, FListe.Row] := TempControl.Text;
        {$ELSE}
        If DS.State=dsBrowse Then DS.Edit ;
        FListe.Fields[iCol].AsString := TempControl.Text;
        {$ENDIF}
        end;
    finally
      TempControl.Free;
      end;
    end;
end;

procedure TOM_CORRESP.ListeKeyPress(Sender: TObject; var Key: Char);
var
  iCol: Integer;
begin
  {$IFDEF EAGLCLIENT}
  iCol := FListe.Col;
  {$ELSE}
  iCol := FListe.SelectedIndex;
  {$ENDIF}

  if IsCompensation and (iCol in [2, 3]) then
  begin
    Key := UpCase(Key);
  end;
end;

procedure TOM_CORRESP.ListeMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
begin  
  FMouseX := X;
  FMouseY := Y;
end;

procedure TOM_CORRESP.ListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin    
  if (Key = VK_F5) and (shift = []) then
  begin
    ListeDblClick(Sender);
    Exit;
  end;
end;

{$IFDEF EAGLCLIENT}
procedure TOM_CORRESP.DoCollExit(Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
{$ELSE}
procedure TOM_CORRESP.DoCollExit(Sender: TObject);
{$ENDIF}
begin
  ControlChamp;
end;

procedure TOM_CORRESP.OnClickBRechercher(Sender: Tobject);
begin
  FFindFirst := True;
  AFindDialog.Execute;
end;

procedure TOM_CORRESP.OnFindAFindDialog(Sender: TObject);
begin
  Rechercher(FListe, AFindDialog, FFindFirst);
end;


function TOM_CORRESP.IsCompensation: Boolean;
begin
  Result := False;
  if gszFType = 'AU' then   {Auxiliaire}
    Result := TCompensation.IsPlanCorresp(StrToInt(gszNoType));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 24/07/2007
Modifié le ... :   /  /
Description .. : Retourne le plan actif
Mots clefs ... :
*****************************************************************}
function TOM_CORRESP.GetNumPlan : Char;
var
  Param1 : string;
  Param2 : string;
  Visibl : Boolean;
begin
  Param1 := 'SO_CORS' + gszFType + '1';
  Param2 := 'SO_CORS' + gszFType + '2';
  {Les radios ne sont visibles que si les deux paramsocs sont cochés}
  Visibl := GetParamSocSecur(Param1, True, True) and
            GetParamSocSecur(Param2, True, True);

  SetControlVisible('RBTN_PLAN', Visibl);
  SetControlVisible('RBTN_PLAN_', Visibl);

  {Par défaut, on affiche le plan de correspondance 1, sauf s'il n'est pas paramétré}
  if Visibl or GetParamSocSecur(Param1, True, True) then Result := '1'
                                                    else Result := '2';
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 24/07/2007
Modifié le ... :   /  /    
Description .. : Permet de verifier si le plan de correspondance dans les 
Suite ........ : paramsoc est activé.
Mots clefs ... : 
*****************************************************************}
function TOM_CORRESP.VerifParamSoc : Boolean;
var
  Param1 : string;
  Param2 : string;
begin
  if (gszFType <> 'GE') and (gszFType <> 'AU') and (gszFType <> 'A1') and (gszFType <> 'A2') and
     (gszFType <> 'A3') and (gszFType <> 'A4') and (gszFType <> 'A5') then Result := TRUE
  else
  begin
     Param1 := 'SO_CORS' + gszFType + '1';
     Param2 := 'SO_CORS' + gszFType + '2';
     Result := GetParamSocSecur(Param1, True, True) or
               GetParamSocSecur(Param2, True, True);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 24/07/2007
Modifié le ... :   /  /    
Description .. : Permet de verifier s'il y a un lien entre un compte existant et 
Suite ........ : le compte de correspondance
Mots clefs ... : 
*****************************************************************}
function TOM_CORRESP.ExisteLien(Compte : String ; Plan : string ; var Mess : string) : boolean;
var
  Q       : TQuery;
begin
  Result := false;
  if (gszFType='GE') then
  begin
     Q      := OpenSQL ('SELECT G_CORRESP' + Plan + ', G_GENERAL FROM GENERAUX WHERE G_CORRESP' + Plan + ' = "' + Compte + '"', TRUE);
     result := not(Q.EOF);
     if not(Q.Eof) then
        Mess   := 'Le compte de correspondance ' + Compte + ' est associé au compte ' + Q.FindField ('G_GENERAL').asString + '. Suppression est impossible';
  end
  else
  if (gszFType='AU') then
  begin
     Q      := OpenSQL ('SELECT T_CORRESP' + Plan + ', T_AUXILIAIRE FROM TIERS WHERE T_CORRESP' + Plan + '="' + Compte + '"', TRUE);
     result := not(Q.EOF);
     if not(Q.Eof) then
        Mess   := 'Le compte de correspondance ' + Compte + ' est associé au compte ' + Q.FindField ('T_AUXILIAIRE').asString + ', suppression est impossible';
  end
  else
  if (gszFType='A1') or (gszFType='A2') or (gszFType='A3') or (gszFType='A4') or (gszFType='A5') then
  begin
     Q      := OpenSQL ('select S_CORRESP' + Plan + ', S_SECTION from SECTION where S_CORRESP' + Plan + '="' + Compte + '" AND S_AXE = "' + gszFType + '"', TRUE);
     result := not(Q.EOF);
     if not(Q.Eof) then
        Mess   := 'La section de correspondance ' + Compte + ' est associée à la section ' + Q.FindField ('S_SECTION').asString + '. Suppression est impossible';
  end;
  if Assigned(Q) then Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 24/07/2007
Modifié le ... :   /  /    
Description .. : Permet de charger les différents controle de la fiche
Mots clefs ... : 
*****************************************************************}
procedure TOM_CORRESP.ChargeControl;
begin
  // Le TRadioButton des deux plans.
  rbtPlan1 := nil; 
  rbtPlan2 := nil;
  if gszFType[1] <> 'I' then
  begin
    rbtPlan1 := TRadioButton(GetControl('RBTN_PLAN'));
    rbtPlan2 := TRadioButton(GetControl('RBTN_PLAN_'));
  end;

  AFindDialog := TFindDialog.Create(Ecran);

  // La grille
{$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FListe'));
  FListe.Options := [goEditing] + [goAlwaysShowEditor] + [goVertLine] + [goHorzLine] + [goTabs];
{$ELSE}
  FListe := THDBGrid(GetControl('FListe')); 
  FListe.Options := [dgEditing] + [dgAlwaysShowEditor] + [dgTitles] + [dgIndicator] + [dgColLines] + [dgRowLines] + [dgTabs];
{$ENDIF}
end;


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 24/07/2007
Modifié le ... :   /  /    
Description .. : Assigne les evenements aux controles de la fiche
Mots clefs ... : 
*****************************************************************}
procedure TOM_CORRESP.ChargeEvenements;
begin
  // Les plan de correspondance
  if (rbtPlan1 <> nil) then
  begin
     rbtPlan1.OnClick := PlanClick;
     rbtPlan1.TabStop := false;
  end;
  if (rbtPlan2 <> nil) then
  begin
     rbtPlan2.OnClick := PlanClick;
     rbtPlan2.TabStop := false;
  end;

  // L'import et l'export
  TRadioButton(GetControl('BExport')).OnClick := ExportClick;
  TRadioButton(GetControl('BImport')).OnClick := ImportClick;

  // La grille
  FListe.OnDblClick   := ListeDblClick;
  FListe.OnKeyPress   := ListeKeyPress;
  FListe.OnKeyDown    := ListeKeyDown;
  FListe.OnMouseMove  := ListeMouseMove;

{$IFNDEF EAGLCLIENT}
  FListe.OnColExit   	:= DoCollExit;
{$ELSE}
  FListe.OnCellExit  	:= DoCollExit;
{$ENDIF}

  AFindDialog.OnFind  := OnFindAFindDialog;
  TToolBarButton97(GetControl('BRechercher', True)).OnClick := OnClickBRechercher;
  ValiderClick := TToolbarButton97(GetControl('BVALIDER', True)).OnClick;
  TToolbarButton97(GetControl('BVALIDER', True)).OnClick := BValiderClick;
end;

procedure TOM_CORRESP.InitialiseGrille(OkInitGrid : Boolean = TRUE);
var
  HM :    THSystemMenu ;
  i  :    Integer;
  TotalWidth:  Integer;
begin
  if IsCompensation then
  begin
     TFFicheGrid(Ecran).ChampsName := 'CR_CORRESP;CR_LIBELLE;CR_LIBRETEXTE1;CR_LIBRETEXTE2';
  end
  else
  begin
     TFFicheGrid(Ecran).ChampsName := 'CR_CORRESP;CR_LIBELLE;CR_ABREGE';
  end;

{$IFDEF EAGLCLIENT}
  if OkInitGrid then
  begin
     FListe.ColCount := 1;
     TFFicheGrid(Ecran).InitGrille ;
  end;
{$ELSE}
  TFFicheGrid(Ecran).InitGrille ;
{$ENDIF}

 // AJOUT ME plantage dans Comsx MajTitreColGrille;
  HM:=THSystemMenu(GetControl('HMTrad')) ;
{$IFDEF EAGLCLIENT}
  if ( HM <> NIL ) and ( FListe <> Nil ) Then HM.ResizeGridColumns(FListe) ;
{$ELSE}
  if ( HM <> NIL ) and ( FListe <> Nil ) Then HM.ResizeDBGridColumns(FListe) ;
{$ENDIF}

  if (gszFType = 'AU') and TCompensation.IsCompensation then
  begin
     TotalWidth := 0;
{$IFDEF EAGLCLIENT}
     for i:=0 to FListe.ColCount-1 do
        Inc(TotalWidth, FListe.ColWidths[i]);
{$ELSE}
     for i:=0 to FListe.columns.Count-1 do
        Inc(TotalWidth, FListe.Columns.Items[i].Width);
{$ENDIF}

{$IFDEF EAGLCLIENT}
    for i:=0 to FListe.ColCount-1 do
       FListe.ColWidths[i] := (TotalWidth div FListe.ColCount);
{$ELSE}
    for i:=0 to FListe.columns.Count-1 do
       FListe.Columns.Items[i].Width := (TotalWidth div FListe.columns.Count);
{$ENDIF}
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 24/07/2007
Modifié le ... :   /  /
Description .. : Met à jour le titre des colonnes de la grille
Mots clefs ... :
*****************************************************************}
procedure TOM_CORRESP.MajTitreColGrille;
begin
{$IFNDEF EAGLCLIENT}
  if gszFType='SIS' then
  begin
    FListe.columns[0].Title.caption := TraduireMemoire(' Compte collectif ');
    FListe.columns[1].Title.caption := TraduireMemoire(' Auxiliaire SISCO de ');
    FListe.columns[2].Title.caption := TraduireMemoire(' à ');
  end
  else if (gszFType='IGE') or (gszFType='IAU') then
  begin
    FListe.columns[0].Title.caption := TraduireMemoire(' Compte Entrant ');
    FListe.columns[1].Title.caption := TraduireMemoire(' Compte dans CEGID Comptabilité ');
    FListe.columns[2].Title.caption := TraduireMemoire(' Intitulé ');
  end
  else if (gszFType='IA1') or (gszFType='IA2') or (gszFType='IA3') or (gszFType='IA4') or (gszFType='IA5') then
  begin
    FListe.columns[0].Title.caption := TraduireMemoire(' Section Entrant ');
    FListe.columns[1].Title.caption := TraduireMemoire(' Section dans CEGID Comptabilité ');
    FListe.columns[2].Title.caption := TraduireMemoire(' Intitulé ');
  end
  else if gszFType='IPM' then
  begin
    FListe.columns[0].Title.caption := TraduireMemoire(' Mode de paiement Entrant ');
    FListe.columns[1].Title.caption := TraduireMemoire(' Mode de paiement dans CEGID Comptabilité ');
    FListe.columns[2].Title.caption := TraduireMemoire(' Intitulé ');
  end
  else if gszFType='IJA' then
  begin
    FListe.columns[0].Title.caption := TraduireMemoire(' Journal Entrant ');
    FListe.columns[1].Title.caption := TraduireMemoire(' Journal dans CEGID Comptabilité ');
    FListe.columns[2].Title.caption := TraduireMemoire(' Intitulé ');
  end
  else if gszFType='IET' then
  begin
    FListe.columns[0].Title.caption := TraduireMemoire(' Etablissement Entrant ');
    FListe.columns[1].Title.caption := TraduireMemoire(' Etablissement dans CEGID Comptabilité ');
    FListe.columns[2].Title.caption := TraduireMemoire(' Intitulé ');
  end;

  if IsCompensation and (FListe.columns.Count > 3) then
  begin
    FListe.columns[2].Title.caption := 'Compte Client et Autres débiteurs';
    FListe.columns[3].Title.caption := 'Compte Fourn. et Autres créditeurs';
  end;
{$ELSE}
  if gszFType='SIS' then
  begin
    FListe.Cells[0, 0] := TraduireMemoire(' Compte collectif ');
    FListe.Cells[1, 0] := TraduireMemoire(' Auxiliaire SISCO de ');
    FListe.Cells[2, 0] := TraduireMemoire(' à ');
  end
  else if (gszFType='IGE') or (gszFType='IAU') then
  begin
    FListe.Cells[0, 0] := TraduireMemoire(' Compte Entrant ');
    FListe.Cells[1, 0] := TraduireMemoire(' Compte dans CEGID Comptabilité ');
    FListe.Cells[2, 0] := TraduireMemoire(' Intitulé ');
  end
  else if (gszFType='IA1') or (gszFType='IA2') or (gszFType='IA3') or (gszFType='IA4') or (gszFType='IA5') then
  begin
    FListe.Cells[0, 0] := TraduireMemoire(' Section Entrant ');
    FListe.Cells[1, 0] := TraduireMemoire(' Section dans CEGID Comptabilité ');
    FListe.Cells[2, 0] := TraduireMemoire(' Intitulé ');
  end
  else if gszFType='IPM' then
  begin
    FListe.Cells[0, 0] := TraduireMemoire(' Mode de paiement Entrant ');
    FListe.Cells[1, 0] := TraduireMemoire(' Mode de paiement dans CEGID Comptabilité ');
    FListe.Cells[2, 0] := TraduireMemoire(' Intitulé ');
  end
  else if gszFType='IJA' then
  begin
    FListe.Cells[0, 0] := TraduireMemoire(' Journal Entrant ');
    FListe.Cells[1, 0] := TraduireMemoire(' Journal dans CEGID Comptabilité ');
    FListe.Cells[2, 0] := TraduireMemoire(' Intitulé ');
  end
  else if gszFType='IET' then
  begin
    FListe.Cells[0, 0] := TraduireMemoire(' Etablissement Entrant ');
    FListe.Cells[1, 0] := TraduireMemoire(' Etablissement dans CEGID Comptabilité ');
    FListe.Cells[2, 0] := TraduireMemoire(' Intitulé ');
  end;

  if IsCompensation and (FListe.ColCount > 3) then
  begin
    FListe.Cells[2, 0] := 'Compte Client et Autres débiteurs';
    FListe.Cells[3, 0] := 'Compte Fourn. et Autres créditeurs';
  end;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 24/07/2007
Modifié le ... :   /  /    
Description .. : Permet d'indiquer le bon titre à la fiche en fonction des 
Suite ........ : paramètres de départ
Mots clefs ... : 
*****************************************************************}
procedure TOM_CORRESP.ChargeTitreEcran;
begin
  // S5 : GE:Généraux ; AU:Auxiliaire ; A1:Axe analytique 1 ; A2:Axe analytique 2
  // S7 : A3:Axe analytique 3 ; A4:Axe analytique 4 ; A5:Axe analytique 5

  // Contexte d aide
  if (gszFType='GE') then
     Ecran.HelpContext := 1325000
  else if (gszFType='AU') then
     Ecran.HelpContext := 1330000
  else if (gszFType='A1') or (gszFType='A2') or (gszFType='A3') or (gszFType='A4') or (gszFType='A5') then
     Ecran.HelpContext := 1134500 ;

  // Titre de la feuille
  if (gszFType='GE') or (gszFType='IGE') then // des comptes généraux
     Ecran.Caption:=Ecran.Caption+' '+MsgBox[1] 
  else
     if (gszFType='AU') or (gszFType='IAU') then // des comptes auxiliaires
        Ecran.Caption:=Ecran.Caption+' '+MsgBox[2] 
     else
        if (gszFType='A1') or (gszFType='IA1') then // de l axe analytique n° 1
	         Ecran.Caption:=Ecran.Caption+' '+MsgBox[3]
        else
   	       if (gszFType='A2') or (gszFType='IA2') then // de l axe analytique n° 2
	            Ecran.Caption:=Ecran.Caption+' '+MsgBox[4]
           else
      	      if (gszFType='A3') or (gszFType='IA3') then // de l axe analytique n° 3
	               Ecran.Caption:=Ecran.Caption+' '+MsgBox[5]
	            else
                 if (gszFType='A4') or (gszFType='IA4') then // de l axe analytique n° 4
		                Ecran.Caption:=Ecran.Caption+' '+MsgBox[6]
	               else
		                if (gszFType='A5') or (gszFType='IA5') then // de l axe analytique n° 5
		                   Ecran.Caption:=Ecran.Caption+' '+MsgBox[7]
	                  else
                       if gszFType='BU' then // des comptes budgétaires
		                      Ecran.Caption:=Ecran.Caption+' '+MsgBox[8]
	                     else
                          if gszFType='IJA' then
			                       Ecran.Caption:=Ecran.Caption+' '+MsgBox[9]
			                    else
                             if gszFType='IPM' then
				                        Ecran.Caption:=Ecran.Caption+' '+MsgBox[10]
			                       else
                                if gszFType='SIS' then
				                           Ecran.Caption:=Ecran.Caption+' '+MsgBox[11]
			                          else
                                   if gszFType='IET' then
				                              Ecran.Caption:=Ecran.Caption+' '+MsgBox[12]
				                           else
                                      Ecran.Caption:=TraduireMemoire(Ecran.Caption) ;
  UpdateCaption(Ecran);
end;

Initialization
  registerclasses ( [ TOM_CORRESP ] ) ; 
end.

