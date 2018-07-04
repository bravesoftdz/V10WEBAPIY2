unit UgestionBobs;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Hctrls, HTB97, ExtCtrls, HPanel,QDialogs,UTOB, Grids,
  HSysMenu,Hent1,HMsgBox;

type
  TForm2 = class(TForm)
    HPanel1: THPanel;
    HPanel2: THPanel;
    EMPLACE: TEdit97;
    Label1: TLabel;
    ToolbarButton972: TToolbarButton97;
    LanceTrait: TToolbarButton97;
    Label2: TLabel;
    FILENAME: TEdit97;
    BfindFile: TToolbarButton97;
    OfindFile: TOpenDialog;
    GS: THGrid;
    HmTrad: THSystemMenu;
    GroupBox1: TGroupBox;
    RBBOB: TRadioButton;
    RBMDL: TRadioButton;
    SaveMDL: TSaveDialog;
    HPanel3: THPanel;
    TBSelect: TToolbarButton97;
    BSaveMDL: TToolbarButton97;
    ToolbarButton971: TToolbarButton97;
    BGENERBOB: TToolbarButton97;
    SauveBOB: TSaveDialog;
    procedure ToolbarButton971Click(Sender: TObject);
    procedure ToolbarButton972Click(Sender: TObject);
    procedure LanceTraitClick(Sender: TObject);
    procedure BfindFileClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BSaveMDLClick(Sender: TObject);
    procedure TBSelectClick(Sender: TObject);
    procedure BGENERBOBClick(Sender: TObject);
  private
    { Déclarations privées }
    fBOBCharge : TOB;
    fLastrepert : string;
  	stcols :string;
  	FTitre :string;
  	FLargeur :string;
  	Falignement :string;
  	title :string;
  	NC :string;
    NBCOLS : integer;
    procedure ConstitueLagrille;
    procedure Affichegrille;
    function DecodeType (TypeData : string) : string;
    function ExtractName (Name : string) : string;
    procedure AddBob(ABOB: TOB; TypeMDL: boolean);
    procedure  EnregistreLigneMdl (TOBECR : TOB; Ligne : integer);
    procedure  EnregistreLigneBOB (TOBECR : TOB; Ligne : integer);

  public
    { Déclarations publiques }
  end;

procedure LancegestionBobs;

implementation

{$R *.dfm}

procedure LancegestionBobs;
var XX : Tform2;
begin
  XX := Tform2.Create(application);
  TRY
  	XX.ShowModal;
  FINALLY
  	XX.free;
  END;
end;

procedure TForm2.ToolbarButton971Click(Sender: TObject);
begin
	close;
end;

procedure TForm2.ToolbarButton972Click(Sender: TObject);
var TheRepert : Widestring;
begin
	if SelectDirectory ('Selectionner un répertoire','',TheRepert) then
  begin
    EMPLACE.Text := TheRepert;
    fLastrepert := TheRepert;
  end;
end;

procedure TForm2.AddBob (ABOB: TOB;TypeMDL : boolean);
var OneBob : TOB;
begin
  OneBOB := TOB.Create ('UNE BOB',fBOBCharge,-1);
  if TypeMDL then
  begin
  	OneBOB.Dupliquer(ABOB,false,true);
  end else
  begin
  	OneBOB.Dupliquer(ABOB,true,true);
  end;
  OneBOB.AddChampSupValeur('ZZPOS',' ');
  OneBOB.AddChampSupValeur('ZZTYPE','0');
  OneBOB.AddChampSupValeur('ZZNAME','');
  OneBOB.AddChampSupValeur('ZZFROM',0);
  OneBOB.AddChampSupValeur('ZZWITHDATA','-');
  if TypeMDL then
  begin
    OneBOB.putvalue('ZZNAME',ExtractName(OneBOB.GetValue('_OBJECTNAME')));
    OneBOB.putvalue('ZZTYPE',DecodeType(OneBOB.GetValue('_OBJECTTYPE')));
    OneBOB.putvalue('ZZFROM',1);
    if (OneBOB.getString ('_OBJECTWITHDATA')='') or (OneBOB.GetInteger('_OBJECTWITHDATA') = -1) then
    begin
    	OneBOB.putvalue('ZZWITHDATA','X');
    	OneBOB.SetString('_OBJECTWITHDATA','-1');
    end else
    begin
    	OneBOB.SetString('_OBJECTWITHDATA','0');
    end;
  end else
  begin
    OneBOB.putvalue('ZZNAME',ExtractName(OneBOB.GetValue('NAME')));
    OneBOB.putvalue('ZZTYPE',DecodeType(OneBOB.GetValue('TYPE')));
    if (OneBOB.GetString('DATA')='') or (OneBOB.GetInteger('DATA') = -1) then
    begin
			OneBob.SetString('DATA','-1');
    	OneBOB.putvalue('ZZWITHDATA','X');
    end else
    begin
			OneBob.SetString('DATA','0');
    end;
  end;
end;

procedure TForm2.LanceTraitClick(Sender: TObject);
var TheBOB : TOB;
		indice : integer;
    TypeMDL : boolean;
    NomFic : string;
    Rec : TSearchRec;
begin
	//
  fBOBCharge.ClearDetail;
  TypeMdl := false;
  if FileName.Text <> '' then
  begin
    if FileExists(FileName.text) then
    begin
    	if ExtractFileExt(FileName.text) = '.MDL' then TypeMDl := true;
			TheBOB := TOB.Create ('UNE BOB',nil,-1);
      TRY
        TOBLoadFromBinFile(FileName.text,nil,TheBOB);
        For Indice := 0 to TheBOB.detail.count -1 do
        begin
          AddBob (TheBOB.detail[Indice],TypeMDL);
        end;
      FINALLY
        Affichegrille;
      	TheBOB.free;
      END;
    end;
  end else if  EMPLACE.Text <> '' then
  begin
  	// traitement via répertoire
    if RBBOB.Checked then
    begin
    	Nomfic:= IncludeTrailingBackslash (EMPLACE.Text)+'*.BOB';
    end else
    begin
      TypeMDl := true;
    	Nomfic:= IncludeTrailingBackslash (EMPLACE.Text)+'*.MDL';
    end;
    TheBOB := TOB.Create ('UNE BOB',nil,-1);
    TRY
      if FindFirst (Nomfic,faAnyFile,Rec) = 0 then
      begin
        if (rec.name <> '.') and (rec.name <> '..') then
        begin
          TRY
            TOBLoadFromBinFile(IncludeTrailingBackslash (EMPLACE.Text)+rec.name,nil,TheBOB);
            For Indice := 0 to TheBOB.detail.count -1 do
            begin
              AddBob (TheBOB.detail[Indice],TypeMDL);
            end;
          FINALLY
            TheBOB.clearDetail;
          END;
        end;
        while FindNext (Rec) = 0 do
        begin
          if (rec.name <> '.') and (rec.name <> '..') then
          begin
            TRY
              TOBLoadFromBinFile(IncludeTrailingBackslash (EMPLACE.Text)+rec.name,nil,TheBOB);
              For Indice := 0 to TheBOB.detail.count -1 do
              begin
                AddBob (TheBOB.detail[Indice],TypeMDL);
              end;
            FINALLY
            	TheBOB.clearDetail;
            END;
          end;
        end;
      end;
    FINALLY
    	TheBOB.free;
    END;
    FindClose (Rec);
    Affichegrille;
  end;
end;

procedure TForm2.BfindFileClick(Sender: TObject);
begin
	 if OfindFile.Execute then
   begin
   	FILENAME.Text := OfindFile.FileName;
   end;
end;

procedure TForm2.FormCreate(Sender: TObject);
begin
	fBOBCharge := TOB.Create ('LE DETAILS DES BOBS',nil,-1);
  ConstitueLagrille;
end;

procedure TForm2.FormClose(Sender: TObject; var Action: TCloseAction);
begin
	fBOBCharge.free;
end;

procedure TForm2.Affichegrille;
begin
  GS.RowCount := fBOBCharge.detail.count+1;
	fBOBCharge.PutGridDetail(GS,false,false,stcols) ;
  HmTrad.ResizeGridColumns(GS); 
end;

procedure TForm2.ConstitueLagrille;
var indice : integer;
		st,lesalignements,leslargeurs,lestitres,lalargeur,letitre : string;
    alignement,FF,Nam : string;
    Dec : integer;
    Sep,Obli,Oklib,OkVisu,OkNulle,OkCumul : boolean;
begin
  NBCOLS := 4;   GS.ColCount := NBCOLS;
  stcols := 'ZZPOS;ZZTYPE;ZZNAME;ZZWITHDATA';
  FTitre := ' ;Objet;Définition;Donnée;';
  FLargeur := '5;15;75;10';
  Falignement := 'G.0  ---;C.0  ---;G.0  ---;C.0  -X-;';
  title := 'Définition des BOB/MDL';
  NC := '1;1;1;1;';
  st := stcols;
  lesalignements := Falignement;
  leslargeurs := FLargeur;
  lesTitres := Ftitre;

  for indice := 0 to NBCOLS -1 do
  begin
    Nam := ReadTokenSt (St); // nom
    alignement := ReadTokenSt(lesalignements);
    lalargeur := readtokenst(leslargeurs);
    letitre := readtokenst(lestitres);
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;

    GS.cells[Indice,0] := leTitre;

    if copy(Alignement,1,1)='G' then GS.ColAligns[indice] := taLeftJustify
    else if copy(Alignement,1,1)='D' then GS.ColAligns[indice] := taRightJustify
    else if copy(Alignement,1,1)='C' then GS.ColAligns[indice] := taCenter;

    GS.ColWidths[indice] := strtoint(lalargeur);
    if Indice = 3 then
    begin
      GS.ColTypes[Indice]:='B' ;
      GS.colformats[Indice]:= inttostr(Integer(csCoche));
    end;
    if OkLib then GS.ColFormats[indice] := 'CB=' + Get_Join(Nam)
    else if (Dec<>0) or (Sep) then GS.ColFormats[indice] := FF ;
  end ;
end;

function TForm2.DecodeType(TypeData: string): string;
begin
  result := '';
	if TypeData = '0' then result := 'Table'
  else if TypeData = '1' then result := 'Tablette'
  else if TypeData = '2' then result := 'Liste'
  else if TypeData = '3' then result := 'Fiche'
  else if TypeData = '4' then result := 'Etat'
  else if TypeData = '5' then result := 'Vue'
  else if TypeData = '6' then result := 'Menu'
  else if TypeData = '7' then result := 'ParamSoc'
  else if TypeData = '8' then result := 'Graph'
  else if TypeData = '9' then result := 'Requete SQL'
  else if TypeData = '17' then result := 'Filtre'
  else if TypeData = '18' then result := 'Présentation'
  else if TypeData = '19' then result := 'Fichier Std'
  else if TypeData = '20' then result := 'Vignette';
end;

function TForm2.ExtractName(Name: string): string;
var Ipos : integer;
begin
	Ipos := pos (':',name);
  if Ipos > 0 then
  begin
    result := copy (Name,Ipos+2,255);
  end else result := Name;

end;

procedure TForm2.BSaveMDLClick(Sender: TObject);
var TOBEcr : TOB;
		Ligne  : integer;
begin
	if GS.nbSelected = 0 then BEGIN PGiInfo('Aucune ligne sélectionnée');exit; END;

	if SaveMdl.Execute then
  begin
    if FileExists(SaveMdl.FileName ) then
    begin
      if PGIAsk('Attention : le fichier MDl existe déjà #13#10 Confirmez vous son écrasement ?')<>Mryes then exit;
    end;
    TOBEcr := TOB.Create ('',nil,-1);
    TRY
      (*
      TOBEcr.AddChampSupValeur('BOBNAME',ExtractFileName(SaveMdl.FileName));
      TOBEcr.AddChampSupValeur('BOBVERSION',1);
      TOBEcr.AddChampSupValeur('BOBNUMSOCREF',0);
      TOBEcr.AddChampSupValeur('BOBCOMMENT',0);
      TOBEcr.AddChampSupValeur('BOBDATEGEN',DateTimeToStr(date()) );
      TOBEcr.AddChampSupValeur('BOBSOCREFDEPART',0 );
      TOBEcr.AddChampSupValeur('BOBSOCREFFINALE',999 );
      TOBEcr.AddChampSupValeur('BOBTYPE',0 ); // --> bob non versionné
      *)
      Ligne := 1;
      repeat
        if (GS.IsSelected (Ligne) ) then
        begin
          EnregistreLigneMdl (TOBECR,Ligne);
        end;
        Inc(Ligne);
      until Ligne >= GS.rowCount;
      TOBECR.SaveToBinFile(SaveMdl.FileName,false,true,true,true);
    FINALLY
      TOBECR.free;
      EMPLACE.Text := '';
      FILENAME.text := '';
      GS.ClearSelected;
      GS.VidePile(false); 
    END;
  end;
end;

procedure TForm2.EnregistreLigneMdl(TOBECR: TOB; Ligne: integer);
var Indice : integer;
		LaTOB,OneBOB : TOB;
    FromMdl : boolean;
begin
	Indice:= Ligne -1;
	laTOB := fBOBCharge.detail[indice];
  FromMdl := (laTOB.getValue('ZZFROM')=1);
  laTOB.DelChampSup ('ZZPOS',false);
  laTOB.DelChampSup('ZZTYPE',false);
  laTOB.DelChampSup('ZZNAME',false);
  laTOB.DelChampSup('ZZFROM',false);
  laTOB.DelChampSup('ZZWITHDATA',false);

  if FromMdl then
  begin
    // provenance d'un MDL -- > Copie directe dans le TOBECR
    OneBOB := TOB.Create ('',TOBECR,-1);
    OneBOB.Dupliquer(laTOB,false,true);
  end else
  begin
    // provenance d'un BOB  --> transformation en MDL
    OneBOB := TOB.Create ('',TOBECR,-1);
    OneBOB.AddChampSupvaleur('_OBJECTCOMMENT','');
    OneBOB.AddChampSupvaleur('_OBJECTTYPE',laTOB.getValue('TYPE'));
    OneBOB.AddChampSupvaleur('_OBJECTNAME',ExtractName(laTOB.getValue('NAME')));
    OneBOB.AddChampSupvaleur('_OBJECTWITHDATA',laTOB.getValue('DATA'));
    OneBOB.AddChampSupvaleur('_OBJECTTYPEETAT',laTOB.getValue('TYPEETAT'));
    OneBOB.AddChampSupvaleur('_OBJECTNATUREETAT',laTOB.getValue('NATUREETAT'));
    OneBOB.AddChampSupvaleur('_OBJECTCODEETAT',laTOB.getValue('CODEETAT'));
    OneBOB.AddChampSupvaleur('_OBJECTLANGUEETAT',laTOB.getValue('LANGUEETAT'));
    OneBOB.AddChampSupvaleur('_OBJECTMN1',laTOB.getValue('MN1'));
    OneBOB.AddChampSupvaleur('_OBJECTMN2',laTOB.getValue('MN2'));
    OneBOB.AddChampSupvaleur('_OBJECTMN3',laTOB.getValue('MN3'));
    OneBOB.AddChampSupvaleur('_OBJECTMN4',laTOB.getValue('MN4'));
    OneBOB.AddChampSupvaleur('_OBJECTNATUREFICHE',laTOB.getValue('NATUREFICHE'));
    OneBOB.AddChampSupvaleur('_OBJECTCODEFICHE',laTOB.getValue('CODEFICHE'));
    OneBOB.AddChampSupvaleur('_OBJECTPARAMSOC',laTOB.getValue('CODEPARAMSOC'));
    OneBOB.AddChampSupvaleur('_OBJECTDOMAINE','');
    OneBOB.AddChampSupvaleur('_OBJECTTABLE1',laTOB.getValue('CODETABLE1'));
    OneBOB.AddChampSupvaleur('_OBJECTTABLE2',laTOB.getValue('CODETABLE2'));
    OneBOB.AddChampSupvaleur('_OBJECTPREVIUSDELETE',laTOB.getValue('PREVIUSDELETE'));
    OneBOB.AddChampSupvaleur('_OBJECTLIBELLE','');
    OneBOB.AddChampSupvaleur('_OBJECTCODEPRODUIT',laTOB.getValue('CODEPRODUIT'));
    OneBOB.AddChampSupvaleur('_OBJECTNOMPRODUIT',laTOB.getValue('NOMPRODUIT'));
    OneBOB.AddChampSupvaleur('_OBJECTCRITERE1',laTOB.getValue('CRITERE1'));
    OneBOB.AddChampSupvaleur('_OBJECTCRITERE2',laTOB.getValue('CRITERE2'));
    OneBOB.AddChampSupvaleur('_OBJECTCRITERE3',laTOB.getValue('CRITERE3'));
    OneBOB.AddChampSupvaleur('_OBJECTCRITERE4',laTOB.getValue('CRITERE4'));
    OneBOB.AddChampSupvaleur('_OBJECTCRITERE5',laTOB.getValue('CRITERE5'));
    OneBOB.AddChampSupvaleur('_OBJECTLANGUESTD',laTOB.getValue('LANGUESTD'));
    OneBOB.AddChampSupvaleur('_OBJECTPREDEFINISTD',laTOB.getValue('PREDEFINISTD'));
  end;
end;

procedure TForm2.TBSelectClick(Sender: TObject);
var Ligne : integer;
begin
	for Ligne := 1 to GS.rowCount do
  begin
    GS.FlipSelection  (Ligne);
  end;
end;

procedure TForm2.BGENERBOBClick(Sender: TObject);
var TOBEcr : TOB;
		Ligne  : integer;
begin
	if GS.nbSelected = 0 then BEGIN PGiInfo('Aucune ligne sélectionnée');exit; END;

	if SauveBOB.Execute then
  begin
    if FileExists(SauveBOB.FileName ) then
    begin
      if PGIAsk('Attention : le fichier BOB existe déjà #13#10 Confirmez vous son écrasement ?')<>Mryes then exit;
    end;
    TOBEcr := TOB.Create ('',nil,-1);
    TRY
      TOBEcr.AddChampSupValeur('BOBNAME',ExtractFileName(SauveBOB.FileName));
      TOBEcr.AddChampSupValeur('BOBVERSION',1);
      TOBEcr.AddChampSupValeur('BOBNUMSOCREF',0);
      TOBEcr.AddChampSupValeur('BOBCOMMENT',0);
      TOBEcr.AddChampSupValeur('BOBDATEGEN',DateTimeToStr(date()) );
      TOBEcr.AddChampSupValeur('BOBSOCREFDEPART',0 );
      TOBEcr.AddChampSupValeur('BOBSOCREFFINALE',9999 );
      TOBEcr.AddChampSupValeur('BOBTYPE',0 ); // --> bob non versionné
      Ligne := 1;
      repeat
        if (GS.IsSelected (Ligne) ) then
        begin
          EnregistreLigneBOB (TOBECR,Ligne);
        end;
        Inc(Ligne);
      until Ligne >= GS.rowCount;
      TOBECR.SaveToBinFile(SauveBOB.FileName,false,true,true,true);
    FINALLY
      TOBECR.free;
      EMPLACE.Text := '';
      FILENAME.text := '';
      GS.ClearSelected;
      GS.VidePile(false);
    END;
  end;
end;

procedure TForm2.EnregistreLigneBOB(TOBECR: TOB; Ligne: integer);
var Indice : integer;
		LaTOB,OneBOB : TOB;
    FromMdl : boolean;
begin
	Indice:= Ligne -1;
	laTOB := fBOBCharge.detail[indice];
  laTOB.DelChampSup ('ZZPOS',false);
  laTOB.DelChampSup('ZZTYPE',false);
  laTOB.DelChampSup('ZZNAME',false);
  laTOB.DelChampSup('ZZFROM',false);
  laTOB.DelChampSup('ZZWITHDATA',false);

  OneBOB := TOB.Create ('',TOBECR,-1);
  OneBOB.Dupliquer(laTOB,true,true);
end;

end.
