{***********UNITE*************************************************
Auteur  ...... : Franck VAUTRAIN
Créé le ...... : 28/06/2011
Modifié le ... :   /  /
Description .. : Class destiné à gérer les grille et leur affichage
Suite ........ : autre que QR1
Mots clefs ... :
*****************************************************************}
unit UtilsGrille;

interface

uses  M3FP,
      StdCtrls,
      Controls,
      Classes,
      Forms,
      SysUtils,
      ComCtrls,
      HCtrls,
      HEnt1,
      HMsgBox,
      UTOB,
      UTOF,
      AglInit,
      Agent,
      EntGC,
      StrUtils,
{$IFDEF EAGLCLIENT}
      MaineAGL,HPdfPrev,UtileAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ENDIF}
      HTB97;


Type
  TGestionGS = class(Tobject)


    private
      FF              : THForm;          
      fNomListe       : String;
      fColNamesGS     : string;
      fLibColNameGS   : String;
      FalignementGS   : string;
      FtitreGS        : string;
      fLargeurGS      : string;
      fGS             : THGrid;
      fTOBG           : TOB;
      fRowHeight      : Integer;
      fColEditableGS  : String;

      function ChargementCombo(Suffixe: string): string;

    public
      property    Ecran         : THForm  read FF             write FF;
      property    NomListe      : String  read fNomListe      write fNomListe;
      property    ColNamesGS    : String  read fColNamesGS    write fColNamesGS;
      property    LibColNameGS  : String  read fLibColNameGS  write fLibColNameGS;
      property    AlignementGS  : String  read FalignementGS  write FalignementGS;
      property    titreGS       : String  read FtitreGS       write FtitreGS;
      property    LargeurGS     : String  read fLargeurGS     write fLargeurGS;
      property    GS            : THGrid  read fGS            write fGS;
      property    TOBG          : TOB     read fTOBG          write fTOBG;
      property    RowHeight     : integer read fRowHeight     write fRowHeight;
      Property    ColEditableGS : String  Read fColEditableGS write fColEditableGS;
      //
      constructor Create;  overload;
      destructor  Destroy; override;
      //
      procedure ChargementGrille;
      Procedure ChargeInfoListe;
      procedure DessineGrille;
      //
      function  GetTOBGrille : TOB;
      function  ZoneAccessible(var ACol, ARow: Integer): boolean;
    end;

implementation

constructor TGestionGS.Create;
begin

  Ecran           := Nil;

  fNomListe       := '';

  fColNamesGS     := '';
  fLibColNameGS   := '';
  FalignementGS   := '';
  FtitreGS        := '';
  fLargeurGS      := '';
  fRowHeight      := 18;
  GS              := nil;
  TOBG            := nil;

end;

procedure TGestionGS.DessineGrille();
var st              : String;
    lestitres       : String;
    lesalignements  : String;
    FF              : String;
    alignement      : String;
    Nam             : String;
    leslargeurs     : String;
    lalargeur       : String;
    letitre         : String;
    lelement        : string;
    TypeFormat      : string;
    LesEditables    : String;
    //
    OkEdit          : Boolean;
    Obli            : Boolean;
    OkLib           : Boolean;
    OkVisu          : Boolean;
    OkNulle         : Boolean;
    OkCumul         : Boolean;
    Sep             : Boolean;
    //Okimg           : boolean;
    //
    dec             : Integer;
    NbCols          : integer;
    indice          : Integer;
    //
    NoCol           : Integer;
begin

  fGS.VidePile(false);
   
  //Calcul du nombre de Colonnes du Tableau en fonction des noms
  st := fColNamesGS;

  NbCols := 0;

  repeat
    lelement := READTOKENST (st);
    //if st='' then continue;
    if lelement <> '' then
    begin
      inc(NbCols);
    end;
  until lelement = '';
  //
  If fGS.DBIndicator then
    fGS.ColCount    := Nbcols+1
  else
    fGS.ColCount    := Nbcols;
  //
  st              := fColNamesGS;
  lesalignements  := falignementGS;
  lestitres       := fLibColNameGS;
  leslargeurs     := flargeurGS;
  LesEditables    := fColEditableGS;

  //Mise en forme des colonnes
  for indice := 0 to Nbcols-1 do
  begin
    Nam := ReadTokenSt (St); // nom
    if Trim(Nam) = '' Then continue;
    //
    alignement  := ReadTokenSt(lesalignements);
    lalargeur   := readtokenst(leslargeurs);
    letitre     := readtokenst(lestitres);
    if LesEditables = '' then
      OkEdit := True
    else
      OkEdit := StrToBool(readtokenst(LesEditables));
    //
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;

    //Permet de gérer si la fiche à un indicateur ou non en première colonne
    If fGS.DBIndicator then
      NoCol := Indice + 1
    else
      NoCol := Indice;

    //
    fGS.cells[NoCol,0]     := leTitre;
    fGS.ColNames [NoCol]   := Nam;
    //
    fGS.Colformats[NoCol]  := '';
    //
    //Alignement des cellules
    if copy(Alignement,1,1)='G'       then //Cadré à Gauche
      fGS.ColAligns[NoCol] := taLeftJustify
    else if copy(Alignement,1,1)='D'  then  //Cadré à Droite
      fGS.ColAligns[NoCol] := taRightJustify
    else if copy(Alignement,1,1)='C'  then
      fGS.ColAligns[NoCol] := taCenter; //Cadré au centre
    //
    if OkLib then
    begin
      TypeFormat := '';
      TypeFormat := Copy(Nam,Pos('_', Nam)+1,Length(Nam));
      if TypeFormat <> ''  then
      begin
        If      TypeFormat ='BTETAT'      then fGS.Colformats[NoCol] := 'CB=BTTYPEACTIONMAT'
        else if TypeFormat ='RESSOURCE'   then fGS.Colformats[NoCol] := 'CB=BTRESSOURCE'
        else if TypeFormat ='CODEFAMILLE' then fGS.Colformats[NoCol] := 'CB=BTFAMILLEMAT'
        else if TypeFormat ='TIERS'       then fGS.Colformats[NoCol] := 'CB=BTTIERS'
        else                                   fGS.Colformats[NoCol] := 'CB=' + ChargementCombo(TypeFormat);
      end;
    end;

    //Colonne visible ou non
    if OkVisu then
      fGS.ColWidths[NoCol] := strtoint(lalargeur) * fGS.Canvas.TextWidth('W')
    else
      fGS.ColWidths[NoCol] := -1;

    //Colonne Editable ou non
    if OkEdit then
      fGS.ColEditables[NoCol] := True
    else
    begin
      fGS.ColEditables[NoCol] := False;
      fGS.ColLengths[NoCol]   := 0;
    end;

    //Affichage d'une image ou du texte
    //if OkImg then FGS.ColDrawingModes[Indice]:= 'IMAGE';

    if (Dec<>0) or (Sep) then
    begin
      if OkNulle then
        fGS.ColFormats[NoCol] := FF+';; ;' //'#'
      else
        fGS.ColFormats[NoCol] := FF; //'#';
    end;

    if fRowHeight <> 0 then GS.DefaultRowHeight := fRowHeight;

  end ;

end;

Function TGestionGS.ChargementCombo(Suffixe : string) : string;
Var StSQl     : string;
    QQ        : TQuery;
begin

  StSQl := 'Select DO_COMBO FROM DECOMBOS WHERE DO_NOMCHAMP LIKE "' + Suffixe + '%"';
  QQ    := OpenSQL(StSQl, True);
  if not QQ.Eof then
    Result := QQ.findfield('DO_COMBO').AsString
  else
    Result := '';

end;

Procedure TGestionGS.ChargementGrille;
var indice  : integer;
    TOBL    : TOB;
    LibChamp: String;
begin

    if fTOBG.Detail.count <> 0 then
      fGS.RowCount := TOBG.detail.count + 1
    else
      Exit;      
    //
    fGS.DoubleBuffered := true;
    fGS.BeginUpdate;

    //Récupération du dernier caractère de la zone Libellé de colonne si ; on le supprime pour ne pas créer une colonne à vide à la fin
    LibChamp := RightStr(Trim(fColNamesGS), 1);
    if Libchamp = ';' then
      Libchamp := copy(Trim(fColNamesGS),1,length(Trim(fColNamesGS))-1)
    else
      Libchamp := fColNamesGS;
    //
    TRY
      fGS.SynEnabled := false;
      for Indice := 0 to fTOBG.detail.count -1 do
      begin
        TOBL := fTOBG.Detail[Indice];
        //fGS.row := Indice+1;
        TOBL.PutLigneGrid(fGS,Indice+1,False,False,LibChamp);
      end;
    FINALLY
      fGS.SynEnabled := true;
      fGS.EndUpdate;
      fGS.Invalidate;
    END;

end;

function TGestionGS.ZoneAccessible(var ACol, ARow: Integer): boolean;
Var TOBTempo : TOB;
begin

  Result := false;

  if ARow = 0                   then Exit;

  if not GS.ColEditables [Acol] then Exit;

  TOBTempo := GetTobGrille;

  If TOBTempo = nil             then Exit;

  Result := True;

end;

function TGestionGS.GetTOBGrille : TOB;
Var Arow : Integer;
begin

  result := nil;

  Arow := fGS.Row;
  
  if (fTOBG = nil) Or (fTOBG.detail.count = 0) Or (fTOBG.detail.count < ARow) then exit;

  if Arow > fTOBG.detail.count Then exit;

  Result := fTOBG.detail[Arow -1];

end;

destructor TGestionGS.Destroy;
begin

  inherited;
end;


procedure TGestionGS.ChargeInfoListe;
Var StSQL     : string;
    QQ        : TQuery;
    TobListe  : TOB;
    Dataliste : Tstringlist;
begin

  if fNomListe = '' then Exit;

  TobListe := tob.Create('LISTE', nil, -1);
  DataListe:= TStringList.Create;

  StSQL := 'SELECT * FROM LISTE WHERE LI_LISTE="' + fNomListe;
  StSQL := StSQL + '" AND LI_LANGUE="'            + V_PGI.LangueDataRef;
  StSQl := StSQL + '" AND LI_UTILISATEUR="'       + V_PGI.User + '"';

  QQ := OpenSQL(StSQL, True);

  // si pas de liste personnalisée
  if QQ.Eof then
  begin
    StSQL := 'SELECT * FROM LISTE WHERE LI_LISTE="' + FNomListe;
    StSQL := StSQL + '" AND LI_LANGUE="'            + V_PGI.LangueDataRef;
    StSQL := StSQL + '" AND LI_UTILISATEUR="---"';
    QQ := OpenSQL(StSQL, True);
    if QQ.Eof then
    begin
      Ferme(QQ);
      FreeAndNil(TobListe);
      fNomListe := '';
      exit;
    end;
  end;

  TobListe.SelectDB('LISTE GRILLE',QQ, False);

  Ferme(QQ);

  FTitreGS := TobListe.GetString('LI_LIBELLE');

  DataListe.SetText(PChar(TobListe.GetString('LI_DATA')));

  //récupèration des champs de la liste
  fColNamesGS    := DataListe.Strings[1];
  //Le libellé en cours
  fLibColNameGS  := Dataliste.Strings[4];
  //la mise en forme de la grille (Taille des colonnes et alignement
  fLargeurGS     := Dataliste.Strings[5];
  FalignementGS  := Dataliste.Strings[6];

  FreeAndNil(DataListe);
  FreeAndNil(TobListe);

end;

end.
