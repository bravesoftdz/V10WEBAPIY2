unit UtilBatiprixFile;

interface

uses Windows,
     classes,
     Graphics,
     FileCtrl,
     SysUtils,
     registry,
     IniFiles,
     forms,
     HCtrls,
     UTOB,
     Hent1,
     HmsgBox,
     math,
     CBPPath,
     UtilFichiers;

type

  TFichier = class;

  // Classe de description d'un champ du fichier
  TdescField = class
    private
      fTypeChamp : string;
      fSeparator : string;
      flongueur : integer;
      fformat : string;
      fnbDecimales : integer;
    public
      procedure SetDescription(const Value: string);
      property TypeChamp : string read fTypeChamp write fTypeChamp;
      property Longueur : integer read flongueur write flongueur;
      property NbDecimales : integer read fnbDecimales write fnbDecimales;
      property format : string read fformat write fformat;
      property Separateur : string read fSeparator write fSeparator;
      // --
      property Descripteur : string write SetDescription;
  end;

  // champ d'un fichier
  TFileField = class
    private
      fparent : Tfichier;
      fNomChamp : string;
      fDescription  : TdescField;
      function Getlongueur: integer;
      function GetTypeChamp: string;
      function getNbDecimales: integer;
      function GetFormat: string;
    public
      property TheParent : Tfichier read fparent Write fparent;
      property description : TdescField read fDescription write fDescription;
      property TypeChamp : string read GetTypeChamp;
      property nbDecimales : integer read getNbDecimales;
      property Longueur : integer read Getlongueur;
      property NomChamp : string read fNomChamp;
      property format : string read GetFormat;
      //
      constructor create;
      destructor destroy ; override;
      procedure SetDescription(const Value: string);
  end;

  // les champs d'un fichier
  TFileFields = class (Tlist)
    private
      fparent : Tfichier;
      function GetItems(Indice : integer): TFileField;
      procedure SetItems(Indice : integer; const Value: TFileField);
      function Add(AObject: TFileField): Integer;
      function Remove(AObject: TFileField): Integer;
      function ConvertitDate(LaValeur,SFormat : string) : TDateTime;
      function RecupLaValeur(var Chaine: string; longueur: integer;FormatFichier, Separateur: string): string;
    public
      property TheParent : Tfichier read fparent Write fparent;
      property Items [Indice : integer] : TFileField read GetItems write SetItems;
      procedure addChamps (UneTOB : TOB);
      procedure putvalues(TOBdest: TOB;Enregistrement:string);
      procedure clear; override;
  end;

  // un fichier
  TFichier = class
    private
      fFileName : string;
      fFields : TFileFields;
      FFormatFic : string; // (F)ixe (V)ariable  = utilisé pour la longueur des zones
      Fseparator : string ; // separateur utilisé
    public
      property Format : string read FFormatFic;
      property Separator : string read Fseparator;
      property FileName : string read fFileName write fFileName;
      function AddField (NomChamp : String; description : string) : boolean;
      function Setdescription (description : string) : boolean;
      procedure AjouteChamps(UneTOB : TOB);
      procedure putvalues(TOBdest: TOB;Enregistrement:string);
      //
      constructor create;
      destructor destroy; override;
  end;

  Tfichiers = class (TList)
    private
      function Add(AObject: Tfichier): Integer;
      function Remove(AObject: Tfichier): Integer;
      function GetItems(Indice: integer): Tfichier;
      procedure SetItems(Indice: integer; const Value: Tfichier);
    public
      property Items [Indice : integer] : Tfichier read GetItems write SetItems;
      function findfichier (nomfichier : string ): TFichier;
      procedure clear; override;
  end;

  TCOntainerFichier = class
    private
      FIniFile: TIniFile;
      fFichiers : Tfichiers;
      FFile : TextFile;
      fFileOpen : boolean;
      fresult : integer;
      ffileEof : boolean;
      fCurrentFile : Tfichier;
      ffiledescriptor : string;
      procedure OpenDescriptionFile;
      procedure SetDescriptions;
      procedure SetFileDefinition (NomFile : string);
      function FindFile (nomfichier : string ): TFichier;
      function GetTOBEnreg (Enregistrement : AnsiString) : TOB;
      procedure SetfileDescriptor(const Value: string);
    public
      property EofFile : boolean read fFileEof;
      property FileDescriptor : string write SetfileDescriptor;
      constructor create; overload;
      constructor create (filedescriptor : string) ; overload;
      destructor destroy; override;
      procedure FermeFile;
      procedure OpenFile (NomFichier : string);
      function ReadFile : TOB;
  end;

  TintegrationError = class
    private
      fErrorFile : TextFile;
      fFileOpen : boolean;
    public
      //
      constructor create;
      procedure Open (repertoire : string);
      procedure close;
      procedure SetError (codeErreur : integer; chaine1 : string=''; chaine2:string=''); overload;
      procedure SetError (codeErreur : integer; chaine1 : string=''; valeur:integer=0); overload;
  end;
implementation
  const
  TexteErreurs: array[1..54] of string = (
  						 {1}  '------ Début traitement %s ------------------------------->'
               {2},	'------ Fin traitement %s ----------------------------------'
							 {3}, 'Erreur de lecture : le fichier séquentiel %s est introuvable'
						{4..9}, '','','','','',''
              {10}, 'Erreur Intégration Famille de Materiaux : la famille %s de niveau %d ---- n'' a pas de père de niveau %d'
     			{11..19}, '','','','','','','','',''
					  	{20}, 'Erreur Intégration Famille d''Ouvrage : la famille %s de niveau %d ---- n'' a pas de père de niveau %d'
          {21..29}, '','','','','','','','',''
							{30}, 'Erreur Intégration des Matériaux: il n'' y a pas de famille de niveau 5 pour l'' article %s'
              {31},	'Erreur Intégration des Matériaux: Le genre %s n''existe pas'
          {32..39}, '','','','','','','',''
							{40}, 'Erreur Intégration des Ouvrages: il n'' y a pas de famille de niveau 4 pour l'' ouvrage %s'
          {41..49}, '','','','','','','','',''
              {50}, 'Erreur Intégration des Détails d''ouvrages : Aucun ID d''ouvrage ne correspond à l''ouvrage composé %s'
              {51}, 'Erreur Intégration des Détails d''ouvrages : Aucun ID d''ouvrage ne correspond à l''ouvrage %s'
              {52}, 'Erreur Intégration des Détails d''ouvrages : Aucun ID d''article ne correspond à l''article %s'
              {53}, '====> Impossible d''enregistrer l''ouvrage composé %s avec l''ouvrage %s'
              {54}, '====> Impossible d''enregistrer l''ouvrage %s avec l''article %s'
							 );

{ TdescField }

procedure TdescField.SetDescription(const Value: string);
var LesDonnees,UneDonnee : string;
begin
  LesDonnees := Value;
  UneDonnee := READTOKENST (LesDonnees);
  fTypeChamp := UneDonnee;
  //
  UneDonnee := READTOKENST (LesDonnees);
  //
  longueur := StrToInt(READTOKENPipe (UneDonnee,','));
  if (UneDonnee <> '') and (longueur>=0) then NbDecimales  := StrToInt(UneDonnee) else NbDecimales := 0;
  //
  if LesDonnees <> '' then
  begin
    format := LesDonnees;
  end;
end;

{ TFileField }

constructor TFileField.create;
begin
  fDescription := TdescField.create;
  fDescription.TypeChamp := '';
  fDescription.Longueur  := 0;
  fDescription.NbDecimales := 0;
  fDescription.Separateur :='';
end;

destructor TFileField.destroy;
begin
  fDescription.free;
  inherited;
end;


function TFileField.GetFormat: string;
begin
  result := fDescription.format;
end;

function TFileField.Getlongueur: integer;
begin
  result := fDescription.Longueur;
end;

function TFileField.getNbDecimales: integer;
begin
  result := fDescription.nbdecimales;
end;

function TFileField.GetTypeChamp: string;
begin
  result := fDescription.fTypeChamp;
end;

procedure TFileField.SetDescription(const Value: string);
begin
  fDescription.SetDescription (Value);
end;

{ TFileFields }

function TFileFields.Add(AObject: TFileField): Integer;
begin
  result := Inherited Add(AObject);
end;

procedure TFileFields.addChamps(UneTOB: TOB);
var Indice : integer;
    unChamp : TFileField;
begin
  for Indice := 0 to Count -1 do
  begin
    UnChamp := items[Indice];
    if TdescField(UnChamp.description).TypeChamp = 'A' then
    begin
      uneTOB.AddChampSupValeur (unChamp.fNomChamp,'');
    end else if TdescField(UnChamp.description).TypeChamp = 'N' then
    begin
      uneTOB.AddChampSupValeur (unChamp.fNomChamp,0);
    end else if TdescField(UnChamp.description).TypeChamp = 'D' then
    begin
      uneTOB.AddChampSupValeur (unChamp.fNomChamp,iDate1900);
    end;
  end;
end;

procedure TFileFields.clear;
var indice : integer;
begin
  Indice := 0;
  for Indice := 0 to Count -1 do
  begin
    TFileField(Items [Indice]).free;
  end;
  inherited;
end;

function TFileFields.ConvertitDate(LaValeur, SFormat: string): TDateTime;
var Iposyear,iPosMonth,iPosDay : integer;
    ByearOk, BMonthOk, BDayOk : boolean;
    Ilongyear,ILongMonth,IlongDay : integer;
    Indice : integer;
    TheCar : string;
    Day : word;
    Month : Word;
    Year : word;
begin
  Indice := 1;
  ByearOk := false; BMonthOk:=false; BDayOk:= false;

  IPosYear := 0; IPosMonth := 0; IPosDay := 0;
  Ilongyear := 0;ILongMonth:=0;IlongDay:=0;
  for Indice := 1 to Length (SFormat) do
  begin
    TheCar := SFormat[Indice];
    if TheCar = 'Y' then
    begin
      if not BYearOk then
      begin
        IPosYear := Indice;
        ByearOk := true;
      end;
      inc(Ilongyear);
    end else if TheCar = 'M' then
    begin
      if not BMonthOk then
      begin
        IPosMonth := Indice;
        BMonthOk := true;
      end;
      inc(IlongMonth);
    end else if TheCar = 'D' then
    begin
      if not BDayOk then
      begin
        IPosDay := Indice;
        BDayOk := true;
      end;
      inc(IlongDay);
    end;
  end;
  //
  Day:=StrtoInt(Copy(LaValeur,IPosDay,ILongDay));
  Month:=StrtoInt(Copy(LaValeur,IPosMonth,ILongMonth));
  year:=StrtoInt(Copy(LaValeur,IPosyear,ILongyear));
  result := EncodeDate (Year,Month,Day);
end;

function TFileFields.GetItems(Indice : integer): TFileField;
begin
 result := TFileField (Inherited Items[Indice]);
end;

procedure TFileFields.putvalues(TOBdest: TOB; Enregistrement: string);
var Indice,decimales : integer;
    unChamp : TFileField;
    Chaine,LaValeur : string;
    Dvaleur : double;
    First : boolean;
    FormatFichier : string; // F/V
    Separateur : string;
    Mess : string;
begin
  Chaine := Enregistrement;
  First := true;
  for Indice := 0 to Count -1 do
  begin
    UnChamp := items[Indice];
    if First then
    begin
      FormatFichier := fparent.Format;
      Separateur := fparent.Separator;
      first := false;
    end;
    decimales := UnChamp.nbDecimales;
    LaValeur := RecupLaValeur (Chaine,UnChamp.longueur,FormatFichier,Separateur);
    if UnChamp.TypeChamp = 'A' then
    begin
      TOBdest.putvalue (UnChamp.NomChamp,LaValeur);
    end else if TdescField(UnChamp.description).TypeChamp = 'N' then
    begin
      Dvaleur := VALEUR (LaValeur);
      if decimales <> 0 then Dvaleur := Dvaleur / power (10,decimales);
      TOBdest.PutValue (unChamp.fNomChamp,DValeur);
    end else if TdescField(UnChamp.description).TypeChamp = 'D' then
    begin
       if lavaleur <> '' then TOBdest.PutValue (unChamp.fNomChamp,ConvertitDate(LaValeur,Unchamp.format));
    end;
  end;
end;

function TFileFields.RecupLaValeur(var Chaine: string; longueur: integer; FormatFichier,Separateur : string): string;
begin
  if FormatFichier = 'F' then
  begin
    result := copy(Chaine,1,Longueur);
    Chaine := Copy(Chaine,Longueur+1,length(chaine));
  end else
  begin
    result := READTOKENPipe (Chaine,Separateur);
  end;
end;

function TFileFields.Remove(AObject: TFileField): Integer;
begin
 result := Inherited Remove(Aobject);
end;

procedure TFileFields.SetItems(Indice : integer; const Value: TFileField);
begin
  Inherited Items[Indice]:= Value;
end;

{ TFichier }

function TFichier.Setdescription(description: string): boolean;
var LesDonnees,UneDonnee : string;
begin
  LesDonnees := description;
  if lesDonnees <> '' then
  begin
    FFormatFic := READTOKENST (LesDonnees); // recup du format
  end;
  //
  if LesDonnees <> '' then
  begin
    Fseparator := READTOKENST (LesDonnees); // separateur
  end;
  //
end;

function TFichier.AddField(NomChamp, description: string): boolean;
var TheField : TFileField;
begin
  TheField := TFileField.create;
  TheField.TheParent := self;
  TheField.fNomChamp := NomChamp;
  TheField.SetDescription (description);
  fFields.Add (TheField);
end;

procedure TFichier.AjouteChamps(UneTOB: TOB);
begin
  fFields.addChamps (UneTOB);
end;

constructor TFichier.create;
begin
  ffields := TFileFields.create;
  ffields.theparent := Self;
  FFormatFic := 'F';
  Fseparator := '';
end;

destructor TFichier.destroy;
begin
  ffIelds.clear;
  fFields.Free;
  inherited;
end;

procedure TFichier.putvalues(TOBdest: TOB; Enregistrement: string);
begin
  fFields.putvalues (TOBDest,Enregistrement);
end;

{ Tfichiers }

function Tfichiers.Add(AObject: Tfichier): Integer;
begin
  result := Inherited Add(AObject);
end;

procedure Tfichiers.clear;
var indice : integer;
begin
  if count > 0 then
  begin
    for Indice := 0 to count -1 do
    begin
      TFichier(Items [Indice]).free;
    end;
  end;
  inherited;
end;

function Tfichiers.findfichier(nomfichier: string): TFichier;
var Indice : integer;
begin
  result := nil;
  for Indice := 0 to Count -1 do
  begin
    if Items[Indice].FileName = NomFichier then
    begin
      result:=Items[Indice];
      break;
    end;
  end;
end;

function Tfichiers.GetItems(Indice: integer): Tfichier;
begin
  result := TFichier (Inherited Items[Indice]);
end;

function Tfichiers.Remove(AObject: Tfichier): Integer;
begin

end;

procedure Tfichiers.SetItems(Indice: integer; const Value: Tfichier);
begin
  Inherited Items[Indice]:= Value;
end;

{ TCOntainerFichier }

procedure TCOntainerFichier.FermeFile;
begin
  if FfileOpen then
  begin
    closeFile(Ffile);
    FfileOpen := false;
    fCurrentFile := nil;
  end;
end;

constructor TCOntainerFichier.create;
begin
  fFichiers := Tfichiers.Create;
  ffileOpen := false;
  ffiledescriptor := '';
end;

destructor TCOntainerFichier.destroy;
begin
  inherited;
  fFichiers.clear;
  fFichiers.free;
end;

function TCOntainerFichier.FindFile(nomfichier: string): TFichier;
begin
  result := fFichiers.findfichier (NomFichier);
end;

function TCOntainerFichier.GetTOBEnreg(Enregistrement: AnsiString): TOB;
var UneTOB : TOB;
    NomFichier : string;
begin
  if ffileOpen then
  begin
    NomFichier := fCurrentFile.FileName;
    UneTOB := TOB.Create ( '+'+NomFichier+'+',nil,-1);
    fCurrentFile.AjouteChamps(UneTOB);
    fCurrentFile.putvalues(UneTOB,Enregistrement);
    result := UneTOB;
  end;
end;

procedure TCOntainerFichier.OpenDescriptionFile;
begin
{
  Rep := ExtractFilePath(Application.ExeName);
  Chemin := Copy (Rep,1,Pos('\APP',UpperCase(Rep))-1)+ '\BATIPRIX\';
}

  FIniFile := TIniFile.Create (ffiledescriptor);
end;

procedure TCOntainerFichier.OpenFile(NomFichier: string);
var Etiquette : string;
begin
  Etiquette := ExtractFileName (NomFichier);
  if Pos('.',Etiquette) > 0 then Etiquette := Copy(Etiquette,1,Pos('.',Etiquette)-1);
  if FFileopen then
  begin
    PgiBox('Un fichier est déjà ouvert');
    exit;
  end;
  FfileOpen := false;
  if FileExists (nomFichier) then
  begin
    fCurrentFile := FindFile (Etiquette);
    if fCurrentFile = nil then exit;
    AssignFile (FFile,NomFichier);
    reset (FFile);
    FfileOpen := true;
  end;
end;

function TCOntainerFichier.ReadFile: TOB;
var enregistrement : AnsiString;
begin
  result := nil;
  enregistrement := '';
  if FfileOpen then
  begin
  	if not eof(Ffile) then
    begin
    	Readln (FFile,enregistrement);
    	ffileEof := eof(Ffile);
    	result := GetTOBEnreg (enregistrement);
    end;
  end;
end;

procedure TCOntainerFichier.SetDescriptions;
var Indice : integer;
    Sections : TStringList;
begin
  Sections := TStringList.Create;
  // Lecture de al definitions du fichier CE.Seq
  FiniFile.ReadSections (Sections);
  for Indice := 0 to Sections.Count -1 do
  begin
    SetFileDefinition (Uppercase(Sections.strings[Indice]));
  end;
  Sections.Clear;
  Sections.free;
end;

procedure TCOntainerFichier.SetFileDefinition(NomFile: string);
var LesChamps : TStringList;
    Indice : integer;
    Ligne,Champ,parametres : string;
    UnFIchier : TFichier;
    LeNom : string;
begin
  LeNom := UpperCase (NomFile);
  lesChamps := TStringList.Create;
  //
  FiniFile.ReadSectionValues (LeNom,LesChamps);
  if lesChamps.Count > 0 then
  begin
    UnFichier := TFichier.create;
    UnFIchier.FileName := LeNom;
    // Ajout d'un descripteur de fichier
    for Indice :=0 to lesChamps.Count -1 do
    begin
      // Recup des infos
      Ligne := lesChamps.Strings [Indice];
      Champ := Copy(Ligne,1,Pos('=',Ligne)-1);
      Parametres := Copy(Ligne,Pos('=',Ligne)+1,Length(Ligne));
      // desciption du fichier
      if Champ = 'DESCRIPTOR' then UnFichier.SetDescription (Parametres)
                              else UnFichier.AddField (Champ,Parametres);
    end;
    fFichiers.Add (unFichier);
  end;
  //
 LesChamps.Clear;
 LesChamps.free;
end;

constructor TCOntainerFichier.create(filedescriptor: string);
begin
  fFichiers := Tfichiers.Create;
  ffileOpen := false;
  ffiledescriptor := filedescriptor;
  OpenDescriptionFile;
  SetDescriptions;
end;

procedure TCOntainerFichier.SetfileDescriptor(const Value: string);
begin
  ffiledescriptor := Value;
  OpenDescriptionFile;
  SetDescriptions;
end;

{ TintegrationError }

procedure TintegrationError.close;
begin
  if fFileOpen then
  begin
    WriteLn(FerrorFile,'------------------------------------');
    CloseFile(ferrorFile);
  end;
end;

constructor TintegrationError.create;
var repertoire,NomFic : string;
begin
	repertoire := IncludeTrailingBackslash (TCBPPAth.GetCegidDataDistri)+IncludeTrailingBackslash('BATIPRIX');
//  repertoire := ExtractFileDrive (Application.exename)+'\PGI00\BATIPRIX\';
  if not DirectoryExists (repertoire) then
  begin
    if not Creationdir (repertoire) then exit;
  end;
  NomFic := repertoire+'ErrorLog.log';
  if not FileExists (nomFic) then CreateFichier (NomFic);
end;

procedure TintegrationError.Open(repertoire: string);
var NomFic : string;
    Rep : string;
begin
  fFileOpen := false;
  if repertoire='' then
  begin
//    Rep := ExtractFilePath(Application.ExeName);
(*    if Pos('\APP',UpperCase(Rep))=0 then
    begin
      // Mode developpement
      repertoire := ExtractFileDrive (Application.exename)+'\PGI00\BATIPRIX\';
    end else
    begin
      repertoire := Copy (Rep,1,Pos('\APP',UpperCase(Rep))-1)+ '\BATIPRIX\';
    end;
*)
  end;
	repertoire := IncludeTrailingBackslash (TCBPPAth.GetCegidDataDistri)+IncludeTrailingBackslash('BATIPRIX');
  if not DirectoryExists (repertoire) then
  begin
    if not Creationdir (repertoire) then exit;
  end;
  NomFic := repertoire+'ErrorLog.log';
  if not FileExists (nomFic) then CreateFichier (NomFic);
  AssignFile(fErrorFile, NomFic);
  {$I-}
  Reset(fErrorFile);
  {$I+}
  if IoResult = 0 then
  begin
    if not fileExists(Nomfic) then rewrite(fErrorFile) else Append(fErrorFile);
    WriteLn(FerrorFile,'------------------------------------');
    WriteLn(FerrorFile,'Messages d''erreur du '+formatDateTime('dd/mm/yyyy hh-nn-ss',Now));
    WriteLn(FerrorFile,'------------------------------------');
    fFileOpen := true;
  end;
end;


procedure TintegrationError.SetError(codeErreur : integer; chaine1 : string='';chaine2:string='');
var Libelle : string;
		ChaineTemp : string;
begin
  if not FfileOpen then exit;
  //
	chaineTemp := TexteErreurs[CodeErreur];
  if (chaine1<>'') then
  begin
  	if chaine2 <> '' then
    begin
    	Libelle := Format(ChaineTemp,[chaine1,chaine2]);
    end else
    begin
    	Libelle := Format(ChaineTemp,[chaine1]);
    end;
  end else
  begin
  	Libelle := ChaineTemp;
  end;
  if (Libelle <> '') then
  begin
    WriteLn (fErrorFile,Libelle);
    Flush(fErrorFile);
  end;
end;

procedure TintegrationError.SetError(codeErreur : integer; chaine1 : string='';valeur: integer=0);
var Libelle : string;
		ChaineTemp : string;
begin
  if not fFileOpen then exit;
	chaineTemp := TexteErreurs[CodeErreur];
  if (chaine1<>'') then
  begin
  	if valeur <> 0 then
    begin
			Libelle := Format(ChaineTemp,[chaine1,valeur,valeur-1]);
    end else
    begin
    	Libelle := Format(ChaineTemp,[chaine1]);
    end;
  end else
  begin
  	Libelle := ChaineTemp;
  end;
  if (Libelle <> '') then
  begin
    WriteLn (fErrorFile,Libelle);
    Flush(fErrorFile);
  end;
end;

end.
