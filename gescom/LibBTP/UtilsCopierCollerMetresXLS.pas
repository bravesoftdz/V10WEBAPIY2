{***********UNITE*************************************************
Auteur  ...... : Franck VAUTRAIN
Créé le ...... : 24/11/2016
Modifié le ... : 24/11/2016
Description .. : Objet servant à gérer les copier/coller
Suite......... : dans les cadre de la gestion des métrés
Suite..........:
Mots clefs ... : Excel;OLE;Outils
*****************************************************************}
unit UtilsCopierCollerMetresXLS;

interface

Uses  Windows,
      Messages,
      SysUtils,
      Classes,
      CBPPath,
      ComObj,
      ComCtrls,
      Controls,
      FileCtrl,
      Hctrls,
      HMsgBox,
      HEnt1,
      db,
      EntGC,
      FactComm,
      UtilsMetresXLS,
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
      uTOB;

    procedure CopieTobVarDoc(TTObPiece, TobDest, TOBL, TobVardoc : Tob);
    procedure CopieTobMetres(TTObPiece, TobDest, TOBL, TobMetres : Tob);
    procedure ColleTobVarDoc(NewLine, OldLine, TobVardoc, TOBL : Tob);
    procedure ColleTobMetres(NewLine, OldLine, TOBL : TOb; TheMetreDoc : TMetreArt);

    function FindLigneOuvrageinDoc(TTObPiece, TOBD: TOB): TOB;

    function IsOuvrage (TOBL : TOB) : boolean;

    function StreamToHex(Stream: TStream; var Hex: string): Boolean;


implementation
uses  EBizUtil,
      UtilFichiers,
      FactUtil,
      IdStrings;



Procedure CopieTobVarDoc(TTObPiece, TobDest, TOBL, TobVardoc : Tob);
var StSQL     : string;
    QQ        : TQuery;
    NatPiece  : string;
    Souche    : String;
    NumPiece  : Integer;
    IndPiece  : Integer;
    NumOrdre  : Integer;
    UniqueBlo : Integer;
    TobEntOuv : TOB;
    TobVar    : TOB;
    TobLVar   : TOB;
    TobTempo  : TOB;
begin

  Try
    If TobDest = nil then exit;

    TobDest.AddChampSupValeur('VARMETRELIGNE', '');

    if IsSousDetail(TOBL) then TobEntOuv := FindLigneOuvrageinDoc(TTObPiece, TOBL);
    //
    if TobEntOuv = nil then TobEntOuv := TOBL;
    //
    NatPiece  := TobEntOuv.GetString('GL_NATUREPIECEG');
    Souche    := TobEntOuv.GetString('GL_SOUCHE');
    NumPiece  := TobEntOuv.GetValue('GL_NUMERO');
    IndPiece  := TobEntOuv.GetValue('GL_INDICEG');
    NumOrdre  := TobEntOuv.GetValue('GL_NUMORDRE');
    UniqueBlo := TobEntOuv.GetValue('UNIQUEBLO');
    //
    if GetPrefixeTable(TobDest)='BLO' THEN
      Uniqueblo := TobDest.GetValue('BLO_UNIQUEBLO')
    else
      Uniqueblo := TobDest.GetValue('UNIQUEBLO');
    //
    TobVar := tob.Create('Les VARIABLES DOC', nil, -1);

    if (TobVardoc = nil) Or (TobVarDoc.Detail.count = 0) then
    begin
      //lecture de la table BTVARDOC pour chargement des cette dernière dans VARMETRELIGNE
      StSQL := 'SELECT * FROM BVARDOC ';
      StSQl := StSQL + 'WHERE BVD_NATUREPIECE ="' + NatPiece + '" ';
      StSQL := StSQL + '  AND BVD_SOUCHE      ="' + Souche       + '" ';
      StSQL := StSQL + '  AND BVD_NUMERO      ='  + IntToStr(NumPiece);
      StSQL := StSQL + '  AND BVD_INDICE      ='  + IntToStr(IndPiece);
      StSQL := StSQL + '  AND BVD_NUMORDRE    ='  + IntToStr(NumOrdre);
      StSQL := StSQL + '  AND BVD_UNIQUEBLO   ='  + IntToStr(UniqueBlo);

      QQ := OpenSQL(StSQL, False,-1,'',True);
      if Not QQ.eof then
      Begin
        TobVar.LoadDetailDB('BVARDOC','','',QQ, False);
      end;

      Ferme(QQ);
    end
    Else
    begin
      //On charge la tobtempo avec les informations de la Tob
      TobTempo := TobVarDoc.FindFirst(['BVD_NATUREPIECE','BVD_SOUCHE','BVD_NUMERO','BVD_INDICE','BVD_NUMORDRE','BVD_UNIQUEBLO'],
                                      [NatPiece,Souche,NumPiece,IndPiece,NumOrdre,UniqueBlo], false);
      while TobTempo <> nil do
      begin
        TobLVar := tob.Create('BVARDOC', TobVar, -1);
        ToblVar.Dupliquer(TobTempo, True, true);
        TobTempo := TobVarDoc.FindNext(['BVD_NATUREPIECE','BVD_SOUCHE','BVD_NUMERO','BVD_INDICE','BVD_NUMORDRE','BVD_UNIQUEBLO'],
                                       [NatPiece,Souche,NumPiece,IndPiece,NumOrdre,UniqueBlo], false);
      end;
    end;

    //On Sauvegarde la TobVar dans la zone associé de la tob ligne
    if TobVar.Detail.count <> 0 then
      TobDest.PutValue('VARMETRELIGNE', TOBtoBINString(TobVar))
    else
      TobDest.PutValue('VARMETRELIGNE', '');

  finally
    FreeAndNil(TobVar);
  end;

end;

Procedure CopieTobMetres(TTObPiece, TobDest, TOBL, TobMetres : Tob);
var StSQL     : string;
    QQ        : TQuery;
    NatPiece  : string;
    Souche    : String;
    NumPiece  : Integer;
    IndPiece  : Integer;
    NumOrdre  : Integer;
    UniqueBlo : Integer;
    FileName  : string;
    TheChaine : string;
    RepMetre  : string;
    FileMetre : TmemoryStream;
    StFile    : string;
    //
    TobEntOuv : TOB;
    TobTempo  : TOB;
begin

  Try
    //L'enregistrement à copier est vide...
    If TobDest = nil then Exit;

    TobDest.AddChampSupValeur('VARFILELIGNE' , '');

    //Nous sommes sur une ligne de sous détail et nous devons remonter à l'ouvrage initial
    if IsSousDetail(TOBL) then TobEntOuv := FindLigneOuvrageinDoc(TTObPiece, TOBL);
    //
    //Si on ne trouve pas d'ouvrage de départ on garde la ligne de l'ouvrage initial du sous-détail.
    if TobEntOuv = nil then TobEntOuv := TOBL;
    //
    NatPiece     := TobEntOuv.GetString('GL_NATUREPIECEG');
    Souche       := TobEntOuv.GetString('GL_SOUCHE');
    NumPiece     := TobEntOuv.GetValue('GL_NUMERO');
    IndPiece     := TobEntOuv.GetValue('GL_INDICEG');
    NumOrdre     := TobEntOuv.GetValue('GL_NUMORDRE');
    UniqueBlo    := TobEntOuv.GetValue('UNIQUEBLO');
    //
    if GetPrefixeTable(TobDest)='BLO' THEN
      Uniqueblo := TobDest.GetValue('BLO_UNIQUEBLO')
    else
      Uniqueblo := TobDest.GetValue('UNIQUEBLO');
    //
    if (TobMetres = nil) Or (TobMetres.Detail.count = 0) then
    begin
      StSQL := 'SELECT BME_NOMDUFICHIER FROM BMETRE ';
      StSQl := StSQL + 'WHERE BME_NATUREPIECEG="' + NatPiece + '" ';
      StSQL := StSQL + '  AND BME_SOUCHE      ="' + Souche       + '" ';
      StSQL := StSQL + '  AND BME_NUMERO      ='  + IntToStr(NumPiece);
      StSQL := StSQL + '  AND BME_INDICEG     ='  + IntToStr(IndPiece);
      StSQL := StSQL + '  AND BME_NUMORDRE    ='  + IntToStr(NumOrdre);
      StSQL := StSQL + '  AND BME_UNIQUEBLO   ='  + IntToStr(UniqueBlo);
      //
      QQ := OpenSQL(StSQL, False,-1,'',True);
      if Not QQ.eof then FileName := QQ.FindField('BME_NOMDUFICHIER').AsString;
      //
      Ferme(QQ);
    end
    else
    begin
      Filename := '';
      //On charge la tobtempo avec les informations de la Tob
      TobTempo := TobMetres.FindFirst(['BME_NATUREPIECEG','BME_SOUCHE','BME_NUMERO','BME_INDICEG','BME_NUMORDRE','BME_UNIQUEBLO'],
                                      [NatPiece,Souche,NumPiece,IndPiece,NumOrdre,UniqueBlo], false);
      if tobTempo <> nil then
        FileName := TobTempo.GetString('BME_NOMDUFICHIER')
      else
        Filename := '';
    end;

    if Filename = '' then Exit;

    //Vérification si le fichier existe dans le répertoire local
    if not FileExists(Filename) then
    begin
      //Vérification si fichier existe sur le serveur
      RepMetre := RecupRepertbase;
      RepMetre := RepMetre + '\' + 'Documents' + '\';
      RepMetre := RepMetre + NatPiece + '-' + Souche + '-' + IntToStr(NumPiece) + '-' + IntToStr(IndPiece) +'\';
      FileName := RepMetre + ExtractFileName(FileName);
      if not FileExists(FileName) then Exit;
    end;

    //Chargement des lignes lues dans une zone Stream
    FileMetre := TMemoryStream.Create;
    FileMetre.LoadFromFile(fileName);
    SetLength(StFile,FileMetre.Size);
    FileMetre.Seek(0,0);
    FileMetre.Read(pchar(Stfile)^,FileMetre.Size);
    if StreamToHex (FileMetre,TheChaine) then
    begin
      //stockage dans champs VARFILELIGNE
      TobDest.PutValue('VARFILELIGNE',TheChaine );
    end;
  Finally
    FileMetre.Free;
  end;

end;


Procedure ColleTobVarDoc(NewLine, OldLine, TobVarDoc, TOBL : TOB);
Var TobVar    : Tob;
    Tobtempo  : TOB;
    TobLVar   : TOB;
    ind       : Integer;
    NatPiece  : string;
    Souche    : String;
    NumPiece  : Integer;
    IndPiece  : Integer;
    NumOrdre  : Integer;
    UniqueBlo : Integer;
begin

  //on récupère la valeur de VARMETRELIGNE Dans une TOBTempo.....
  If not OldLine.FieldExists('VARMETRELIGNE') then Exit;

  if OldLine.GetValue('VARMETRELIGNE') = '' then Exit;

  TobVar := tob.Create('Les VARIABLES DOC', nil, -1);

  BINStringToTOB(OldLine.GetValue('VARMETRELIGNE'), TobVar);

  If TobVar = nil then exit;

  if GetPrefixeTable(NewLine) = 'GL' then
  Begin
    NatPiece  := NewLine.GetValue('GL_NATUREPIECEG') ;
    Souche    := NewLine.GetValue('GL_SOUCHE') ;
    NumPiece  := StrToInt(NewLine.GetValue('GL_NUMERO'));
    IndPiece  := StrToInt(NewLine.GetValue('GL_INDICEG'));
    NumOrdre  := StrToInt(TOBL.GetValue('GL_NUMORDRE'));
    UniqueBlo := StrToInt(NewLine.GetValue('UNIQUEBLO'));
  end
  else
  begin
    NatPiece  := NewLine.GetValue('BLO_NATUREPIECEG') ;
    Souche    := NewLine.GetValue('BLO_SOUCHE') ;
    NumPiece  := StrToInt(NewLine.GetValue('BLO_NUMERO'));
    IndPiece  := StrToInt(NewLine.GetValue('BLO_INDICEG'));
    NumOrdre  := StrToInt(TOBL.GetValue('GL_NUMORDRE'));
    UniqueBlo := StrToInt(NewLine.GetValue('BLO_UNIQUEBLO'));
  end;

  //on charge la tobVardoc associée à la pièce...
  For ind := 0 to TobVar.detail[0].detail.count -1 do
  begin
     TobTempo := TobVar.detail[0].detail[ind];
     TobLVar  := tob.Create('BVARDOC', TobVarDoc, -1);
     TobTempo.PutValue('BVD_NATUREPIECE',  NatPiece);
     TobTempo.PutValue('BVD_SOUCHE',       Souche);
     TobTempo.PutValue('BVD_NUMERO',       NumPiece);
     TobTempo.PutValue('BVD_INDICE',       IndPiece);
     TobTempo.PutValue('BVD_NUMORDRE',     NumOrdre);
     TobTempo.PutValue('BVD_UNIQUEBLO',    UniqueBlo);
     if Not NewLine.fieldExists('GL_NUMORDRE') then TobTempo.PutValue('BVD_INDEXFILE', 999999);
     TobLvar.Dupliquer(TobTempo, True, True);
  end;

end;

Procedure ColleTobMetres(NewLine, OldLine, TOBL : Tob; TheMetreDoc : TMetreArt);
Var StFile    : String;
    FileName  : String;
    CodeArt   : String;
    FileMetre : TMemoryStream;
    NatPiece  : string;
    Souche    : String;
    NumPiece  : Integer;
    IndPiece  : Integer;
    NumOrdre  : Integer;
    UniqueBlo : Integer;
    TobLMetre : Tob;
begin

  If not OldLine.FieldExists('VARFILELIGNE') then Exit;

  //on récupère la valeur de VARMETRELIGNE Dans une TOBTempo.....
  if OldLine.GetValue('VARFILELIGNE') = '' then Exit;

  NatPiece  := TOBL.GetValue('GL_NATUREPIECEG') ;
  Souche    := TOBL.GetValue('GL_SOUCHE') ;
  NumPiece  := StrToInt(TOBL.GetValue('GL_NUMERO'));
  IndPiece  := StrToInt(TOBL.GetValue('GL_INDICEG'));
  NumOrdre  := StrToInt(TOBL.GetValue('GL_NUMORDRE'));

  if GetPrefixeTable(NewLine) = 'GL' then
  Begin
    UniqueBlo := StrToInt(NewLine.GetValue('UNIQUEBLO'));
    CodeArt   := NewLine.GetValue('GL_ARTICLE');
  end
  else
  begin
    UniqueBlo := StrToInt(NewLine.GetValue('BLO_UNIQUEBLO'));
    CodeArt   := NewLine.GetValue('BLO_ARTICLE');
  end;

  //Récupération chemin du fichier
  StFile := OldLine.GetValue('VARFILELIGNE');

  if DirectoryExists(TheMetreDoc.frepMetreDocLocal) then
  begin
    FileName := TheMetredoc.fRepMetreDocLocal + CodeArt + '-' + IntToStr(NumOrdre) + '-' + IntToStr(UniqueBlo) + '.xlsx';
    //Chargement des lignes lues dans une zone Stream
    if not FileExists(FileName) then
    begin
      FileMetre := TMemoryStream.Create;
      FileMetre.Size := Length(StFile) div 2;
      if FileMetre.Size > 0 then
      begin
        HexToBin(PChar(StFile),FileMetre.Memory,FileMetre.size);
        FileMetre.SaveToFile (FileName);
      end;
      FileMetre.Free;
      //Création de l'enregistrement associé dans la TOBMetres
      TobLMetre := Tob.Create('BMETRE', TheMetreDoc.TTobMetres, -1);
      TobLMetre.PutValue('BME_NATUREPIECEG', NatPiece);
      TobLMetre.PutValue('BME_SOUCHE', Souche);
      TobLMetre.PutValue('BME_NUMERO', IntToStr(NumPiece));
      TobLMetre.PutValue('BME_INDICEG', IntToStr(IndPiece));
      TobLMetre.PutValue('BME_NUMORDRE', IntToStr(NumOrdre));
      TobLMetre.PutValue('BME_UNIQUEBLO', IntToStr(UniqueBlo));
      TobLMetre.PutValue('BME_ARTICLE', CodeArt);
      TobLMetre.PutValue('BME_NOMDUFICHIER', FileName);
    end;
  end;

end;

function FindLigneOuvrageinDoc(TTObPiece, TOBD: TOB): TOB;
var Inddep      : Integer;
    Indice      : integer;
		indOuvrage  : integer;
    Prefixe     : String;
begin

	result  := nil;

  Prefixe := GetPrefixeTable(TOBD);

  Inddep  := TOBD.GetIndex ;

  IndOuvrage := TOBD.GetValue(Prefixe + '_INDICENOMEN');

  for Indice := Inddep-1 downto 0 do
  begin
    if ISOuvrage(TTObPiece.detail[Indice]) then
    begin
      if prefixe <> 'BLO_' then
      begin
        if (TTObPiece.detail[Indice].getValue('GL_INDICENOMEN')=IndOuvrage) then
        begin
          if (IsArticle(TTObPiece.detail[Indice])) then
          begin
            Result := TTObPiece.detail[Indice];
            break;
          end;
        end;
      end
      else
      begin
        if (IsArticle(TTObPiece.detail[Indice])) then
        begin
          Result := TTObPiece.detail[Indice];
          break;
        end;
      end;
    end;
  end;

end;


function IsOuvrage (TOBL : TOB) : boolean;
var TypeArticle : string;
		prefixe     : string;
begin

  result := false;

  prefixe := GetPrefixeTable(TOBL);

  TypeArticle := TOBL.getValue(Prefixe+'_TYPEARTICLE');

  if (TypeArticle = 'ARP') or (TypeArticle = 'OUV') then result := true;

end;

function StreamToHex(Stream: TStream; var Hex: string): Boolean;
var
  Buffer: Byte;
begin
  Hex := EmptyStr;
  Stream.Seek(0, soFromBeginning);
  while Stream.Read(Buffer, 1) = 1
    do Hex := Hex + IntToHex(Buffer, 2);
  Result := not SameText(Hex, EmptyStr)
end;

end.
