{***********UNITE*************************************************
Auteur  ...... : Sylvie DE ALMEIDA
Créé le ...... : 02/03/2007
Modifié le ... : 02/03/2007
Description .. : Source TOF de la FICHE : YYCOMPTESPARCAT ()
Mots clefs ... : TOF;YYCOMPTESPARCAT
*****************************************************************}
Unit YYCOMPTESPARCAT_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
     UtilPGI,
{$IFNDEF EAGLCLIENT}
     db,dbTables,Fiche,FE_main,FichList, EdtREtat, HDB,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eFiche,maineagl,eFichList,HPdfPrev,UtileAGL,
     eMul, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,
     UTOB,
     Windows,
     {$IFDEF TAXES}
     YYDETAILTAXES_TOF,
     {$ENDIF}
     HTB97;
{$IFDEF TAXES}
procedure YYLanceFiche_CompteParCatTaxe;
Type
  TOF_YYCOMPTESPARCAT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
      {$IFNDEF EAGLCLIENT}
      FListe     : THDBGrid;
      {$ELSE}
      FListe     : THGrid;
      {$ENDIF}
      CbModele   : THValComboBox;
      CbCategorie: THValComboBox;
      CbType     : THValComboBox;
      BDetail     : TToolbarButton97;
      procedure DoDblClick(Sender: TObject);
      procedure OnChangeCategorie (Sender: TObject);
      procedure OnChangeModele (Sender: TObject);
      procedure MAJComboCat;
      procedure MAJComboType;
      procedure OuvreDetail(strModele, strCategorie, strType, strRegime : string);
      function  GetColIndex    ( vStColName : String ) : Integer ;
      procedure OuvreDet(Sender: TObject);

  end ;
{$ENDIF}
Implementation
{$IFDEF TAXES}
uses TntGrids, TntDBGrids, DBGrids, Grids;

//==============================================================================

procedure YYLanceFiche_CompteParCatTaxe;
begin

    AglLanceFiche('YY','YYCOMPTESPARCAT','','','') ;
end;
//==============================================================================

procedure TOF_YYCOMPTESPARCAT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_YYCOMPTESPARCAT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_YYCOMPTESPARCAT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_YYCOMPTESPARCAT.OnLoad ;
begin
  Inherited ;
  setcontrolVisible ('PCOMPLEMENT', false); // onglet compléments non visible
end ;

procedure TOF_YYCOMPTESPARCAT.OnArgument (S : String ) ;
begin
  Inherited ;
  // Récupération des contrôles
  {$IFNDEF EAGLCLIENT}
  FListe := THDBGrid(GetControl('FListe'));
  {$ELSE}
  FListe := THGrid(GetControl('FListe'));
  {$ENDIF}
  if (assigned(FListe)) then
    FListe.OnDblClick := DoDblClick;
  cbModele := THValComboBox(GetControl('CCT_CODEMODELE'));
  CbModele.OnChange := OnChangeModele;
  CbCategorie := THValComboBox(GetControl('CCT_CODECAT'));
  CbCategorie.OnChange := OnChangeCategorie;
  CbType := THValComboBox(GetControl('CCT_CODETYP'));
  BDetail := TToolbarButton97 (GetControl('BDETAIL'));
  BDetail.OnClick := OuvreDet;
end ;

procedure TOF_YYCOMPTESPARCAT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_YYCOMPTESPARCAT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_YYCOMPTESPARCAT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_YYCOMPTESPARCAT.OnChangeModele(Sender: TObject);
begin
   MAJComboCat;
end;

procedure TOF_YYCOMPTESPARCAT.OnChangeCategorie(Sender: TObject);
begin
   MAJComboType;
end;

procedure TOF_YYCOMPTESPARCAT.MAJComboCat;
var
   strSQL, St, St1, va, it : string;
   OQuery : TQuery ;
   strModele : string;
begin

   //Mise à jour combo catégorie
   cbCategorie.Clear; //on initialise la combo des catégories
   if CbModele.itemindex = -1 then
    exit;
   strModele := CbModele.values[CbModele.itemindex];

   strSQL := 'SELECT "TX1", YMODELETAXE.YMT_LIBCAT1 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT1 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT1 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';
   strSQL := strSQL + ' UNION ';
   strSQL := strSQL + 'SELECT "TX2", YMODELETAXE.YMT_LIBCAT2 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT2 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT2 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';
   strSQL := strSQL + ' UNION ';
   strSQL := strSQL + 'SELECT "TX3", YMODELETAXE.YMT_LIBCAT3 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT3 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT3 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';
   strSQL := strSQL + ' UNION ';
   strSQL := strSQL + 'SELECT "TX4", YMODELETAXE.YMT_LIBCAT4 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT4 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT4 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';
   strSQL := strSQL + ' UNION ';
   strSQL := strSQL + 'SELECT "TX5", YMODELETAXE.YMT_LIBCAT5 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT5 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT5 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';

  OQuery := OpenSql (strSQL, TRUE);
  try
  While not OQuery.Eof do
  begin
    St := OQuery.Fields[0].AsString;
    St1 := OQuery.Fields[1].AsString;
    va:=ReadTokenSt(St) ;
    it:=TraduireMemoire(ReadTokenSt(St1)) ;
    cbCategorie.Values.Add(va) ;
    cbcategorie.Items.Add(it) ;
    OQuery.next;
  end;
  finally
    OQuery.free;
  end;

end;

procedure TOF_YYCOMPTESPARCAT.MAJComboType;
var
   strSQL, St, St1, va, it : string;
   OQuery : TQuery ;
   strModele : string;
begin
   //Mise à jour combo types de taxes 
   cbType.Clear; //on initialise la combo des types de taxes
   if CbModele.itemindex = -1 then
    exit;
   strModele := CbModele.values[CbModele.itemindex];

  strSQL := 'SELECT YMODELECATTYP.YCY_CODETYP, YTYPETAUX.TYP_LIBELLE FROM YMODELECATTYP INNER JOIN YTYPETAUX ON YMODELECATTYP.YCY_CODETYP = YTYPETAUX.TYP_CODE WHERE YMODELECATTYP.YCY_TYPGERE = "X" AND YMODELECATTYP.YCY_CODECAT  = "'+CbCategorie.values[CbCategorie.itemindex]+'" AND YMODELECATTYP.YCY_CODEMODELE = "'+strModele+'"';
  OQuery := OpenSql (strSQL, TRUE);
  try
  While not OQuery.Eof do
  begin
    St := OQuery.Fields[0].AsString;
    St1 := OQuery.Fields[1].AsString;
    va:=ReadTokenSt(St) ;
    it:=TraduireMemoire(ReadTokenSt(St1)) ;
    cbType.Values.Add(va) ;
    cbType.Items.Add(it) ;
    OQuery.next;
  end;
  finally
    OQuery.free;
  end;
end;

procedure TOF_YYCOMPTESPARCAT.DoDblClick(Sender: TObject);
var
  iModele, iCategorie, iType, iRegime : integer;
begin
if GetDataSet.Bof and GetDataSet.Eof then Exit ; // si rien on sort
//ouverture fenêtre détails
iModele := GetColIndex('CCT_CODEMODELE');
iCategorie := GetColIndex('CCT_CODECAT');
iType := GetColIndex('CCT_CODETYP');
iRegime := GetColIndex('CCT_CODEREG');
{$IFNDEF EAGLCLIENT}
if (FListe.Columns[iModele].Field.AsString <> '') then
  OuvreDetail (FListe.Columns[iModele].Field.AsString, FListe.Columns[iCategorie].Field.AsString,FListe.Columns[iType].Field.AsString,FListe.Columns[iRegime].Field.AsString);
{$ELSE}
if (Fliste.Cells[iModele, Fliste.Row] <> '') then
  OuvreDetail (Fliste.CellValues[iModele, Fliste.Row],Fliste.CellValues[iCategorie, Fliste.Row],Fliste.CellValues[iType, Fliste.Row],Fliste.CellValues[iRegime, Fliste.Row]);
{$ENDIF}
end;

procedure TOF_YYCOMPTESPARCAT.OuvreDetail(strModele, strCategorie, strType, strRegime : string);
var
  Code, stWhere: STRING;
begin
//ouverture fenêtre détails
  stWhere := 'CODEMOD='+ strModele+ ';CODECAT=' + strCategorie + ';CODETYP=' + strType +';CODEREG='+ strRegime;
  Code := AGLLanceFiche ('YY', 'YYDETAILTAXES','YMT_CODE='+strModele, '', ''+stWhere );

end;

function TOF_YYCOMPTESPARCAT.GetColIndex(vStColName: String): Integer;
var
i : Integer ;
begin
  result := -1 ;
{$IFNDEF EAGLCLIENT}
   for i := 0 to FListe.Columns.Count - 1 do
    begin
    if FListe.Columns[i].FieldName = vStColName then
      begin
      result := i ;
      Exit ;
      end ;
    end ;
{$ELSE}
  for i := 0 to FListe.ColCount - 1  do
    begin
    if FListe.ColNames[i] = vStColName then
      begin
      result := i ;
      Exit ;
      end ;
    end ;
{$ENDIF}
end;

procedure TOF_YYCOMPTESPARCAT.OuvreDet(Sender: TObject);
begin
//ouverture fenêtre détails
AglLanceFiche('YY','YYDETAILTAXES','','','') ;

end;



Initialization
  registerclasses ( [ TOF_YYCOMPTESPARCAT ] ) ;

{$ENDIF} 
end.

