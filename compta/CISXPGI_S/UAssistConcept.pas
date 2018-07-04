unit UAssistConcept;

interface

uses
  Windows, SysUtils, Classes, Controls, Forms, Dialogs,
  assist, hmsgbox, StdCtrls, ComCtrls, ExtCtrls, HTB97, Hctrls,
  Mask, Uscript, Db, HDB, HQry,
  HRichOLE, ASCIIV, Grids, Buttons, Spin, UConcept, HEnt1, UTOB,
  UControlParam, licUtil, ImptSql,
  bde,  FileCtrl, Clipbrd, UScriptDelim,
{$IFDEF CISXPGI}
  uYFILESTD,
  ULibCpContexte,
{$ENDIF}

{$IFDEF EAGLCLIENT}
  UScriptTob,
  UscriptCwas,
  ImgList,
  HRichEdt, HSysMenu, HPanel,
{$ELSE}
{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet, Variants, ADODB,  {$ENDIF}
{$ENDIF}
   DBGrids;


type
  TFAssistExt = class(TFAssist)
    TabNom: TTabSheet;
    HLabel1: THLabel;
    GroupBox2: TGroupBox;
    Label25: TLabel;
    Label19: TLabel;
    Code: THCritMaskEdit;
    Label22: TLabel;
    edcomment: TEdit;
    TedNature: TLabel;
    EdNature: THCritMaskEdit;
    RTypefichier: TRadioGroup;
    RadioMode: TRadioGroup;
    LFichier: TLabel;
    FILENAME: THCritMaskEdit;
    TabDelimite: TTabSheet;
    Label28: TLabel;
    cbDelimChamp: TComboBox;
    ChargementDelim: TBitBtn;
    Label29: TLabel;
    GD: THGrid;
    TabResultat: TTabSheet;
    Memo1: TASCIIView;
    Label1: TLabel;
    FicheSortie: THCritMaskEdit;
    TTypetransfert: TLabel;
    Typetransfert: TComboBox;
    TabSelect: TTabSheet;
    FZBase: TGroupBox;
    ListBox1: TListBox;
    BAjouter: TToolbarButton97;
    BEnlever: TToolbarButton97;
    FModif: TGroupBox;
    ListBox2: TListBox;
    TabRequete: TTabSheet;
    GroupBox3: TGroupBox;
    Label3: TLabel;
    HLabel5: THLabel;
    EdReqSuivante: THCritMaskEdit;
    TabResultR: TTabSheet;
    SQL: TGroupBox;
    HM: THMsgBox;
    SQ: TDataSource;
    HLabel6: THLabel;
    GroupBox4: TGroupBox;
    HLabel7: THLabel;
    GroupBox5: TGroupBox;
    GroupBox6: TGroupBox;
    HLabel2: THLabel;
    TFUser: THLabel;
    cbOrigine: TComboBox;
    lblAliasODBC: THLabel;
    cbAliasODBC: TComboBox;
    EdUtilisateur: TEdit;
    TFPassword: THLabel;
    EdMotdepasse: TEdit;
    HLabel4: THLabel;
    cbTypeSortie: TComboBox;
    ImageList1: TImageList;
    HLabel8: THLabel;
    Chemin: THCritMaskEdit;
    CheckBoxTitre: TCheckBox;
    lblDebut: TLabel;
    ChampDebut: TSpinEdit;
    GRILVISU: THGrid;
    Label2: TLabel;
    Champfin: TSpinEdit;
    ModeParam: TRadioGroup;
    TabParamsoc: TTabSheet;
    GroupBox7: TGroupBox;
    Label4: TLabel;
    ComboDevise: TComboBox;
    Label5: TLabel;
    Nbdecimale: TSpinEdit;
    HLabel9: THLabel;
    GroupBox8: TGroupBox;
    Label6: TLabel;
    PariteEuro: THCritMaskEdit;
    GroupBox9: TGroupBox;
    Label7: TLabel;
    LGgene: TSpinEdit;
    Label8: TLabel;
    LGauxi: TSpinEdit;
    Label9: TLabel;
    LGaxe1: TSpinEdit;
    GroupBox10: TGroupBox;
    Analytique: TCheckBox;
    COMPAUX: TCheckBox;
    Label10: TLabel;
    Champref: TSpinEdit;
    Label11: TLabel;
    CARELIM: TEdit;
    Label12: TLabel;
    Chaineref: TEdit;
    Label13: TLabel;
    Nbexercice: TSpinEdit;
    Statistique: TCheckBox;
    CheckSauve: TCheckBox;
    BVoir: TBitBtn;
    Label15: TLabel;
    cbTypeCar: TComboBox;
    Z_SQL: THRichEditOLE;
    HInfoFichier: THLabel;
    GroupBox11: TGroupBox;
    Label14: TLabel;
    BValide: TToolbarButton97;
    EdShellExecute: THCritMaskEdit;
    GroupBox12: TGroupBox;
    Label16: TLabel;
    Editeur: THCritMaskEdit;
    Car: TEdit;
    LEtap: THLabel;
    LblMAJLongueur: TLabel;
    AutoLigneRef: TSpeedButton;
    EditAppliquerLong: TSpinEdit;
    bMAJLongueur: TToolbarButton97;
    HLabel3: THLabel;
    Domaine: THValComboBox;
    Label17: TLabel;
    DELIMLIB: TEdit;
    FListe: THGrid;
    procedure FormCreate(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure ChargementDelimClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure GDDblClick(Sender: TObject);
    procedure GDKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure GDSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure RTypefichierClick(Sender: TObject);
    procedure DOMAINEExit(Sender: TObject);
    procedure BAjouterClick(Sender: TObject);
    procedure BEnleverClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure cbTypeSortieChange(Sender: TObject);
    procedure cbOrigineChange(Sender: TObject);
    function  DetermineTypeODBC(Typ : TFieldType;sDsn : string):TFieldType;
    procedure EdUtilisateurExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure CheminElipsisClick(Sender: TObject);
    procedure ChampfinChange(Sender: TObject);
    procedure ChampDebutChange(Sender: TObject);
    procedure ListBox2DragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure ListBox2DragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure RenseigneTobParam(TPControle : TSVControle);
    procedure DOMAINEChange(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure BVoirClick(Sender: TObject);
    procedure cbTypeCarChange(Sender: TObject);
    procedure cbDelimChampChange(Sender: TObject);
    procedure CodeExit(Sender: TObject);
    procedure bMAJLongueurClick(Sender: TObject);
    procedure AutoLigneRefClick(Sender: TObject);

  private
    { Déclarations privées }
    FScript                : TScript;
    SavRow,SavCol          : Integer;
{$IFNDEF EAGLCLIENT}
    DBODBC                 : TDatabase;
{$ENDIF}
    SL_Origin, SL_ODBC     : TStringList;
    ParametrageCegid       : boolean;
    FTypeSortie            : TTypeSortie;
    FDestSortie            : TDestSortie;
    FUtilisateur           : string;
    OkRequete              : Boolean;
    CommentC               : TActionFiche;
    CComplement, LePays    : string;
    OkFin,ModifInterface   : Boolean;
    DomaineParam,NtTrans   : string;
    Indexcb                : integer;
    OkDelim                : Boolean;
    LongFich               : integer;
    CarDelim               : string;
    PSuiv                  : integer;
    FScriptRequete         : TScriptRequete;
{$IFNDEF CISXPGI}
    Q          : TQuery;
{$ENDIF}
    procedure BValiderDelimite(OkShellexcute : Boolean= TRUE);
    function  TrouveLongueur (ind : integer): string;
    procedure BChercheClick;
{$IFNDEF CISXPGI}
    procedure ChargeSQL;
    procedure BValiderClickRequete;
{$ENDIF}
    procedure ChargementFichierDeTraitement;
    procedure EnableTabparamsoc;
    procedure SauvegardeFiltre;
    function  ChargeDelim(Codedelim : string; FichEnt: string=''; FichS : string='') : Boolean;
    procedure SuppressionScriptDelim;
    procedure LibereDelim;
    function  CompteDelim(sLigne : string; sDelim : string) : integer; // compte le nombre de délimiteurs d'une chaine
    procedure MajColonneLongueur;  // Mettre à jour la colonne longueur dans le paramétrage délimité
    procedure SelectRequete;
  public
    { Déclarations publiques }
    constructor create (OkR : Boolean; Comment : TActionFiche; Complement : string; py : string); reintroduce; overload ; virtual ;

  end;

var
  FAssistExt: TFAssistExt;
Const ColonneCoche : Integer = 2;


procedure CreateScript(OkR : Boolean; Comment : TActionFiche; Table : string=''; Shw:Boolean=TRUE; Complement : string=''; Pays :string='');
procedure ModifScriptDelim(Table : string=''; Nat: string=''; Nt: string=''; Edt : string=''; Dm : string='');
function TraiteScriptDelim(Table,FileEntree,FileSortie : string; OkShellexcute : Boolean= TRUE): Boolean;
{$IFDEF CISXPGI}
procedure LanceCorrespCisx (Cmd : string);
{$ENDIF}


implementation

{$R *.DFM}

uses UPDomaine,
{$IFNDEF EAGLCLIENT}
UVoir,
{$ENDIF}
UDMIMP;


const lgdevise = 14;
type

Tabledevise = record
             code     : string;
             libelle     : string;
end;
var
Correspdevise : array [0..lgdevise] of Tabledevise =
   (
   (code : 'EUR' ; libelle : 'euros'),
   (code : 'ATS' ; libelle : 'schillings autrichiens'),
   (code : 'BEF' ; libelle : 'francs belges'),
   (code : 'CHF' ; libelle : 'francs suisses'),
   (code : 'DEM' ; libelle : 'deutschemark'),
   (code : 'ESP' ; libelle : 'pesetas'),
   (code : 'FIN' ; libelle : 'marks finlandais'),
   (code : 'GBP' ; libelle : 'livres'),
   (code : 'ITL' ; libelle : 'lires'),
   (code : 'JPY' ; libelle : 'yen'),
   (code : 'LUF' ; libelle : 'francs luxembourgeois'),
   (code : 'NLG' ; libelle : 'florins'),
   (code : 'PTE' ; libelle : 'escudos'),
   (code : 'USD' ; libelle : 'dollars'),
   (code : 'FRF' ; libelle : 'Franc')
   );

procedure CreateScript(OkR : Boolean; Comment : TActionFiche; Table : string=''; Shw:Boolean=TRUE; Complement : string=''; Pays :string='');
var
TPControle : TSVControle;
begin
  FAssistExt             := TFAssistExt.create(OkR, Comment, Complement, Pays);
  with FAssistExt do
  begin
        if OkR then
        begin
             if CommentC = taCreat then
                  FScriptRequete := TScriptRequete.Create(nil)
             else
                  FScriptRequete := LoadScriptRequete (Table);
        end
        else
           FScript := TScript.Create(nil);

           if Shw then ShowModal
           else
           begin
                with FScriptRequete do
                begin
                  cbAliasODBC.ItemIndex := cbAliasODBC.Items.IndexOf(AliasName);
                  if cbAliasODBC.ItemIndex = -1 then cbAliasODBC.Text := ALiasName;
                  FTypeSortie := TypeSortie;
                  Code.Text   := NomFichier;
                  FDestSortie := DestSortie;
                  EdReqSuivante.Text := ScriptSuivant;
                  EdUtilisateur.text := Utilisateur;
                  EdMotdepasse.text := Motdepasse;
                  EdComment.Text := FComment;
                  Code.text := Name;
                  Chemin.text := RChemin;
                  cbOrigine.Text := DriverName;
                  //Z_SQL.Text := ExtraireMemo(UPPERCASE(DropComment(Requete.Text)));
                  AssignSqlMemo (Z_SQL, Requete);
                end;
{$IFNDEF CISXPGI}
                  Q.DatabaseName     := cbAliasODBC.text;
                 if DBODBC <> nil then DBODBC.destroy;
                 DBODBC              := TDatabase.Create(nil);
                 DBODBC.DriverName   := cbOrigine.Text;
                 if EdUtilisateur.text <> '' then
                    DBODBC.Params.Add('USER NAME='+EdUtilisateur.text);
                 if Chemin.Text <> '' then
                    DBODBC.Params.Add('PATH='+Chemin.Text)
                 else
                    DBODBC.Params.Add('PATH='+CurrentDonnee);

                 if EdMotdepasse.text <> '' then
                    DBODBC.Params.Add('PASSWORD='+EdMotdepasse.text);
                 DBODBC.Params.Add('LANGDRIVER=Paradox intl');
                 DBODBC.DatabaseName := cbAliasODBC.text;
                      DBODBC.LoginPrompt  := false;
                 if cbOrigine.Items[cbOrigine.ItemIndex ] <> 'STANDARD' then
                 begin
                  {$IFNDEF DBXPRESS}
                      DBODBC.AliasName    := Q.DatabaseName;
                  {$ENDIF}
                      DBODBC.LoginPrompt  := false;
                 end;
                 DBODBC.Open;
                 BChercheClick;
                 if DBODBC <> nil then DBODBC.destroy;
{$ENDIF}
           end;
  end;
  if not OkR then
  begin
    With FAssistExt do
    begin
            TPControle := TSVControle.create;
            if OkFin then
            begin
               TPControle.LePays := Lepays;
               TPControle.LeMode := CComplement;
               if ListBox2.Items.count = 0 then
                   TPControle.ChargeTobParam(DomaineParam,TRUE,nil)
               else
                   TPControle.ChargeTobParam(DomaineParam,TRUE, ListBox2);
               RenseigneTobParam(TPControle);
               PExtConcept(FScript, Domaine.Text, Code.Text, EdNature.Text, Typetransfert.Text, Editeur.text, taCreat, TPControle, COMPAUX.checked, Nbexercice.value, Statistique.Checked, EdShellExecute.Text, Analytique.Checked);
               if (Comment = taCreat) and (not ExisteScript (Code.Text)) and (RTypefichier.ItemIndex = 1) then
                  SuppressionScriptDelim;
            end
            else
            begin
                if TPControle <> nil then
                begin
                     TPControle.TOBParamFree;
                     TPControle.free;
                end;
            end;
            FScript.Destroy;
    end;
  end;
  if FAssistExt.OkRequete then
     FAssistExt.FScriptRequete.Destroy;
  FAssistExt.ListBox2.clear;
  FAssistExt.free;

end;

procedure ModifScriptDelim(Table : string=''; Nat: string=''; Nt: string=''; Edt : string=''; Dm : string='');
begin
  FAssistExt                     := TFAssistExt.create(Application);
  FAssistExt.CommentC            := taModif;
  FAssistExt.OkRequete           := FALSE;
  FAssistExt.OkDelim             := TRUE;
  FAssistExt.LongFich            := 1;
  FAssistExt.EdNature.Text       := Nat;
  FAssistExt.NtTrans             := Nt;
  FAssistExt.Editeur.text        := Edt;
  FAssistExt.ModifInterface      := TRUE;
  FAssistExt.DomaineParam        := Dm;
  if FAssistExt.ChargeDelim(Table) then
  FAssistExt.ShowModal
  else
   FAssistExt.LibereDelim;
  FAssistExt.free;

end;

function TraiteScriptDelim(Table,FileEntree,FileSortie : string; OkShellexcute : Boolean= TRUE) : Boolean;
begin
  Result := FALSE;
  FAssistExt           := TFAssistExt.create(Application);
  With  FAssistExt do
  begin
       CommentC  := taModif;
       OkRequete := FALSE;
       OkDelim   := TRUE;
       LongFich  := 99;
       ModifInterface := TRUE;
       if ChargeDelim(Table,FileEntree,FileSortie) then
       begin
         BValiderDelimite(OkShellexcute);
         Result := TRUE;
       end;
       LibereDelim;
       free;
  end;
end;

procedure TFAssistExt.FormCreate(Sender: TObject);
begin
  inherited;
  FScript := nil;
  FScriptRequete := nil;
{$IFNDEF CISXPGI}
         ChargementComboDomaine (Domaine, FALSE);
         ChargeSQL;
{$ELSE}
         Domaine.dataType  := 'CPZIMPDOMAINE';
{$ENDIF}
end;

procedure TFAssistExt.bFinClick(Sender: TObject);
var
aPreTrt    : TPreTrt;
begin
  inherited;
  if RadioMode.ItemIndex = 0 then // mode script
  begin
        if FScript = nil  then
        begin
           if OkDelim then  // traitement delimite
           begin
               if CheckSauve.checked then SauvegardeFiltre;
           end;
           ModalResult := MrOk; exit;
        end;
        FScript.Name     := Copy (Code.text,1, 12);
        FScript.FComment := Edcomment.Text;
        if (RTypefichier.ItemIndex = 0) then
           FScript.Shellexec  :=  EdShellExecute.Text
        else
           FScript.Shellexec  :=  '';
        FScript.FNouveau := False;
        FScript.FileType := 1; // pour ascii à voir
        FScript.ParName  := FScript.Name;
        Fscript.options.asciiFilename := FScript.Name;
        FScript.FileDest              := 0;
        //FScript.ASCIIMODE             := RTypefichier.ItemIndex;
        FScript.ASCIIMODE             := 0;
        if (RTypefichier.ItemIndex = 1)   then // traitement delimite
        begin
           FScript.Options.FileName   := FicheSortie.Text;
           if CheckSauve.checked then
           begin
                SauvegardeFiltre;
                aPreTrt := TPreTrt.Create(FScript.PreTrt);
                aPreTrt.Nom := 'TRAITEMENT DELIMITE';
                aPreTrt.Param.Add('DELIMITE');
           end;
        end
        else
           FScript.Options.FileName := FileName.Text;
        FScript.Modeparam := Modeparam.Itemindex;
  end;
{$IFNDEF CISXPGI}
  if RadioMode.ItemIndex = 1 then // mode SQL
  begin
         BValiderClickRequete;
  end;
{$ENDIF}
  OkFin := TRUE;
  ModalResult := MrOk;
end;

procedure TFAssistExt.ChargementDelimClick(Sender: TObject);
var
F                   : TextFile;
Tb,TF               : TOB;
col,i,ii            : integer;
SLect,tmp           : string;
carDel              : char;
OkChargement        : Boolean;
begin
  inherited;

    if cbDelimChamp.ItemIndex <> -1 then
    begin
      carDel := CarSepField[cbDelimChamp.ItemIndex];
      if (cbDelimChamp.ItemIndex = 5) then
         carDel := Car.Text[1];
    end
    else carDel := cbDelimChamp.Text[1];

    OkChargement := FALSE;
    AssignFile(F, FileName.Text);
    Reset(F);
    col :=1;
    Tb := TOB.Create('Enregistrement', nil, -1);
    if Chaineref.Text <> '' then
    begin
              for ii:= 1 to Champfin.value do
              begin
                  Readln(F, SLect);
                  if (pos(Chaineref.Text, SLect) <> 0) then
                  begin
                    for i:= 0 to length(SLect) do
                    begin
                      tmp := ReadTokenPipe (SLect, carDel);
                      if (tmp = '') and (SLect = '') then break;
                      TF := TOB.Create('fille'+IntToStr(col), Tb, -1);
                      if not CheckBoxTitre.checked then
                        TF.AddChampSupValeur('Colonne', 'C'+IntToStr(col))
                      else
                        TF.AddChampSupValeur('Colonne', tmp) ;
                        TF.AddChampSupValeur('Longueur', IntToStr(length(tmp)));
                        TF.AddChampSupValeur('Visible', 'X');
                        TF.AddChampSupValeur('Separateur', 'X');
                        TF.AddChampSupValeur('ALin', 'X');
                        inc (col);
                      end;
                      Champdebut.value := ii;
                      OkChargement := TRUE;
                      break;
                    end;
                  end;
    end
    else
      for ii:= 1 to Champref.value do
      begin
          Readln(F, SLect);
          if ii <> Champref.value then continue;
          for i:= 0 to length(SLect) do
          begin
            tmp := ReadTokenPipe (SLect, carDel);
            if (tmp = '') and (SLect = '') then break;
            TF := TOB.Create('fille'+IntToStr(col), Tb, -1);
            if not CheckBoxTitre.checked then
               TF.AddChampSupValeur('Colonne', 'C'+IntToStr(col))
            else
               TF.AddChampSupValeur('Colonne', tmp) ;
            TF.AddChampSupValeur('Longueur', IntToStr(length(tmp)));
            TF.AddChampSupValeur('Visible', 'X');
            TF.AddChampSupValeur('Separateur', 'X');
            TF.AddChampSupValeur('ALin', 'X');
            inc (col);
            OkChargement := TRUE;
          end;
        end;
    CloseFile(F);
    Tb.PutGridDetail(GD, False, False, 'Colonne;Longueur;Visible;Separateur;ALin');
    Tb.Free;
    if OkChargement then
    begin
          GD.ColTypes[ColonneCoche]:='B';
          GD.ColFormats[ColonneCoche]:=IntToStr(Ord(csCheckbox));
          GD.ColEditables[ColonneCoche]:=FALSE;

          GD.ColTypes[ColonneCoche+1]:='B';
          GD.ColFormats[ColonneCoche+1]:=IntToStr(Ord(csCheckbox));
          GD.ColEditables[ColonneCoche+1]:=FALSE;

          GD.ColTypes[ColonneCoche+2]:='B';
          GD.ColFormats[ColonneCoche+2]:=IntToStr(Ord(csCheckbox));
          GD.ColEditables[ColonneCoche+2]:=FALSE;

          GD.RowCount := col;
    end;
    ChargementFichierDeTraitement;
end;

procedure TFAssistExt.ChargementFichierDeTraitement;
var
Fe                  : TextFile;
ii                  : integer;
SLect               : string;
TbVisu,TFV          : TOB;
begin
  inherited;
    if (CommentC  = taModif) and (not FileExists(FileName.Text)) then
    begin
        ii :=1;
        TbVisu := TOB.Create('Visu', nil, -1);
        TFV := TOB.Create('fille'+IntToStr(ii), TbVisu, -1);
        TFV.AddChampSupValeur('Ligne', IntToStr(ii));
        TFV.AddChampSupValeur('Contenu', '');
        GRILVISU.RowCount := ii;
        ChampFin.MaxValue := ii+1;
        ChampFin.Value := ii-1;
        Champref.MaxValue := ii+1;
        TbVisu.PutGridDetail(GRILVISU, False, False, 'Ligne;Contenu');
        TbVisu.Free;
        exit;
    end;

    AssignFile(Fe, FileName.Text);
    {$I-} Reset (Fe) ; {$I+}
    ii :=1;
    TbVisu := TOB.Create('Visu', nil, -1);
     while not EOF(Fe) do
     begin
        Readln(Fe, SLect);
        if ii < ChampDebut.value then begin inc(ii); continue; end;
        TFV := TOB.Create('fille'+IntToStr(ii), TbVisu, -1);
        TFV.AddChampSupValeur('Ligne', IntToStr(ii));
        TFV.AddChampSupValeur('Contenu', SLect);
        inc(ii);
     end;
     CloseFile(Fe);
     GRILVISU.RowCount := ii;
     ChampFin.MaxValue := ii+1;
     ChampFin.Value := ii-1;
     Champref.MaxValue := ii+1;
     TbVisu.PutGridDetail(GRILVISU, False, False, 'Ligne;Contenu');
     TbVisu.Free;
end;

procedure TFAssistExt.BValiderDelimite(OkShellexcute : Boolean= TRUE);
var
SLect,tmp           : string;
F,FS                : TextFile;
i, li,Lg,Lg1        : integer;
carDel                 : char;
FormatLigne,Ligne,ll: string;
P, pch              : PChar;
Rep                 : string;
          procedure CarElimine;
          begin
                    if Pos(CARELIM.Text, tmp) <> 0 then
                    begin
                         P := PChar(tmp);
                         repeat
                               pch := StrScan(P, CARELIM.text[1]);
                               if pch <> nil then
                               begin
                                    pch^ := ' ';
                                    inc(pch);
                               end;
                               P := pch;
                        until pch = nil;
                    end;
          end;
        function ReadTokenPipeSaveText(var SInput: string; Car : Char): string;
        var
          P           : Pchar;
          S2          : string;
          FMode, ii   : Integer;
        begin
          if SInput = '' then begin Result := ''; exit; end;
          P := PChar(SInput);
          S2 := '';
          Fmode := 0;  ii := 1;
          while P^ <> #0 do
          begin
            if Fmode = 1 then
            begin
              if P^ = DELIMLIB.Text[1] then
                Fmode := 0;
              S2 := S2 + P^
            end
            else if P^= DELIMLIB.Text[1] then
            begin
              Fmode := 1;
              S2 := P^;
            end
            else
            begin
              if P^ <> Car then
                 S2 := S2 + P^
              else break;
            end;
            Inc(P); inc (ii);
          end;
          Result := S2;
          SInput := Copy(SInput, ii+1, Length(Sinput)-ii);
        end;
begin
  inherited;
    Memo1.FileName :=FILENAME.Text;
    Memo1.Ignored := 0;
    Memo1.Hint := FILENAME.Text;
    memo1.UpdateLines;

    ModalResult := MrNone; CarDel := '0';
    if Filename.Text = '' then
    begin
         Pgiinfo ('Selectionnez le fichier à transformer', 'Traitement ASCII Délimité');
         exit;
    end;
    if FicheSortie.Text = '' then
    begin
         Pgiinfo ('Selectionnez le fichier de Sortie', 'Traitement ASCII Délimité');
         exit;
    end;
    if cbDelimChamp.Text = '' then
    begin
              Pgiinfo ('Selectionnez un Délimiteur de champ', 'Traitement ASCII Délimité');
              exit;
    end;
    if GD.Cells[1, 1] = '' then
    begin
              Pgiinfo ('Charger les longueurs', 'Traitement ASCII Délimité');
              exit;
    end;
    if RTypefichier.ItemIndex = 1 then
    begin
         if (cbDelimChamp.ItemIndex <> -1)  then
         begin
                  carDel := CarSepField[cbDelimChamp.ItemIndex];
                  if (cbDelimChamp.Text = 'Autres') and (cbDelimChamp.ItemIndex = 5) then
                  begin
                    cbDelimChamp.Text := Car.Text[1];  carDel := Car.Text[1];
                  end;
         end
         else carDel := cbDelimChamp.Text[1];
    end;
    if (EdShellExecute.Text <> '') and OkShellexcute then
    begin
             Memo1.CloseMemo;
             ExecuteShell(EdShellExecute.Text);
    end;
    Rep := ExtractFileDir (FicheSortie.Text);
    if Rep = '' then
    begin
        PgiBox('Renseignez le répertoire du fichier #10'+FicheSortie.Text,'Traitement ASCII Délimité') ;
        Exit ;
    end;

    AssignFile(FS, FicheSortie.Text) ;
    Rewrite(FS) ;
    if IoResult<>0 then
    begin
        PgiBox('Impossible d''écrire dans le fichier #10'+FicheSortie.Text,'Traitement ASCII Délimité') ;
        CloseFile(FS);
        Exit ;
    end ;

    if not FileExists(FileName.Text) then
    begin
        PgiBox('Impossible de lire le fichier #10'+FileName.Text,'Traitement ASCII Délimité') ;
        CloseFile(FS);
        Exit ;
    end;

    AssignFile(F, FileName.Text);
    {$I-} Reset (F) ; {$I+}
    li  := 1;
    while not EOF(F) do
    begin
      Readln(F, SLect);
      if li < ChampDebut.value then begin inc(li); continue; end;
      Ligne := '';
      for i:= 1 to length(SLect) do
      begin
          ll := TrouveLongueur (i);
          if ll = '' then break;
          if (GD.Cells[3, i]= '-') or (GD.Cells[2, i]= '-') then
          begin
            FormatLigne := Format('%%-%s.%ss',[ll,ll]);
            if DELIMLIB.Text <> '' then
                    tmp := ReadTokenPipeSaveText (SLect, CarDel)
            else
                    tmp := ReadTokenPipe (SLect, CarDel);
            CarElimine;
            if (GD.Cells[3, i]= '-') then
            begin
              if i = 0 then
                Ligne := Format(FormatLigne,[tmp])+carDel
              else
                Ligne := Ligne + Format(FormatLigne,[tmp])+carDel;
            end
            else
            begin
              if i = 0 then
                Ligne := Format(FormatLigne,[tmp])
              else
                Ligne := Ligne + Format(FormatLigne,[tmp]);
            end;
          end
          else
          begin
              if ll <> '0' then
              begin
                   if DELIMLIB.Text <> '' then
                    tmp := ReadTokenPipeSaveText (SLect, CarDel)
                   else
                    tmp := ReadTokenPipe (SLect, CarDel);
                    Lg := Length(tmp);
                    Lg1:= StrToInt(ll);
                    if GD.Cells[4, i] = '-' then
                    begin
                       if Lg = Lg1 then
                       begin
                           FormatLigne := Format('%%-%s.%ss',[ll,ll]);
                           CarElimine;
                           if i = 0 then Ligne := Format(FormatLigne,[tmp])
                           else Ligne := Ligne + Format(FormatLigne,[tmp]);
                       end
                       else
                       begin
                         if Lg1 > Lg then
                         begin
                               FormatLigne := Format('%%-%d.%ds',[Lg1-Lg,Lg1-Lg])+Format('%%-%d.%ds',[Lg,Lg]);
                               CarElimine;
                               if i = 0 then
                                 Ligne := Format(FormatLigne,['', tmp])
                               else
                                 Ligne := Ligne + Format(FormatLigne,['', tmp]);
                          end
                          else
                          begin
                               FormatLigne := Format('%%-%d.%ds',[Lg1,Lg1]);
                               CarElimine;
                               if i = 0 then
                                 Ligne := Format(FormatLigne,[tmp])
                               else
                                 Ligne := Ligne + Format(FormatLigne,[tmp]);
                          end;

                       end;
                    end
                    else
                    begin
                           FormatLigne := Format('%%-%s.%ss',[ll,ll]);
                           CarElimine;
                           if i = 0 then Ligne := Format(FormatLigne,[tmp])
                           else Ligne := Ligne + Format(FormatLigne,[tmp]);
                    end;
              end;
          end;
      end;
      writeln(FS, Ligne);
      inc(li);
      if (ChampFin.value = 1) and (LongFich > 1) and  (CommentC = taModif) then
      begin
           if (li > LongFich) then break;
      end
      else
         if (li > ChampFin.value) and (LongFich <> 99) then break;
    end;
    CloseFile(F);
    CloseFile(FS);
    Memo1.FileName :=FicheSortie.Text;
    Memo1.Ignored := 0;
    Memo1.Hint := FicheSortie.Text;
    memo1.UpdateLines;
end;


function TFAssistExt.TrouveLongueur (ind : integer): string;
var
tmp      : string;
begin
     tmp  := (Trim(GD.Cells[1, ind]));
     if GD.Cells[2, ind]= '-' then
        Result := '0'
     else
        Result := (GD.Cells[1, ind]);
        //AVOIR Result := (Trim(GD.Cells[1, ind]));
end;
procedure TFAssistExt.bSuivantClick(Sender: TObject);
      Procedure     AfficheEtape;
      begin
        inc(Psuiv);
        lEtap.Caption := Msg.Mess[0] + ' ' + IntToStr(PSuiv);
      end;
begin
  inherited;
  if RadioMode.Itemindex = 0 then
  begin
      if (P.ActivePage = TabResultat) and (RTypefichier.ItemIndex = 1) then
      begin
        BValiderDelimite(FALSE);
        exit;
      end;
      if (P.ActivePage = TabSelect) and (RTypefichier.ItemIndex = 1) and (CommentC  = taModif) then
      begin
        bFinClick(Sender);
        exit;
      end;
      if (P.ActivePage = TabParamsoc) and (OkDelim or (DomaineParam = 'O') or (DomaineParam = 'X'))then
      begin
          if Indexcb <> -1 then cbDelimChamp.ItemIndex := Indexcb
          else
          begin
               cbDelimChamp.ItemIndex := 5;
               cbDelimChamp.Text      := CarDelim;
               Car.Text               := CarDelim;
               Car.visible            := TRUE;
          end;
           bSuivantClick(Sender);  exit;
      end;
      if (P.ActivePage = TabSelect) and (DomaineParam = 'O') then
      begin
          bFinClick(Sender);  exit;
      end;
      if (P.ActivePageIndex <> 1) and (P.ActivePageIndex < 4) and (RTypefichier.ItemIndex = 0) then
         begin bSuivantClick(Sender); exit; end;
      if (P.ActivePage = TabSelect) and (ModeParam.ItemIndex = 1) then
       begin bSuivantClick(Sender);  exit; end;

      if (P.ActivePage = TabSelect) and ((RTypefichier.ItemIndex = 0) or (Domaine.Enabled = FALSE))
      and (ListBox1.Items.Count = 0) then DOMAINEExit(Sender);
      if (P.ActivePageIndex >= 5)  then bFinClick(Sender);
      if (P.ActivePage = TabParamsoc) then EnableTabparamsoc;
  end;
// Cela sert-il vraiment à quelque chose ? (charger l'affichage dans la page que l'on veut quitter)
  //if (P.ActivePage = TabDelimite) and (RTypefichier.ItemIndex = 1) then
          //ChargementFichierDeTraitement;
  if (RadioMode.ItemIndex = 1) then
  begin
       if (P.ActivePageIndex < 5) then bSuivantClick(Sender);
       if (P.ActivePageIndex = 6) then
       begin
{$IFNDEF CISXPGI}
            Q.DatabaseName     := cbAliasODBC.text;
           if DBODBC <> nil then DBODBC.destroy;
           DBODBC              := TDatabase.Create(nil);
           DBODBC.DriverName   := cbOrigine.Text;
           if EdUtilisateur.text <> '' then
              DBODBC.Params.Add('USER NAME='+EdUtilisateur.text);
           if Chemin.Text <> '' then
              DBODBC.Params.Add('PATH='+Chemin.Text)
           else
              DBODBC.Params.Add('PATH='+CurrentDonnee);

           if EdMotdepasse.text <> '' then
              DBODBC.Params.Add('PASSWORD='+EdMotdepasse.text);
           DBODBC.Params.Add('LANGDRIVER=Paradox intl');
           DBODBC.DatabaseName := cbAliasODBC.text;
                DBODBC.LoginPrompt  := false;
           if cbOrigine.Items[cbOrigine.ItemIndex ] <> 'STANDARD' then
           begin
           {$IFNDEF DBXPRESS}
                DBODBC.AliasName    := Q.DatabaseName;
           {$ENDIF}
                DBODBC.LoginPrompt  := false;
           end;
           DBODBC.Open;
           BChercheClick;
{$ENDIF}
       end;
  end;
       AfficheEtape;
end;

procedure TFAssistExt.GDDblClick(Sender: TObject);
begin
  inherited;
  if (SavCol = 2) or  (SavCol = 3) or (SavCol = 4)then
  begin
        if GD.Cells[SavCol, SavRow] = 'X' then
          GD.Cells[SavCol, SavRow] := '-'
        else
          GD.Cells[SavCol, SavRow] := 'X';
  end;
end;

procedure TFAssistExt.GDKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
  case Key of
    32 :
      begin
        if GD.Cells[SavCol, SavRow] = 'X' then
          GD.Cells[SavCol, SavRow] := '-'
        else
          GD.Cells[SavCol, SavRow] := 'X';
      end;
  end;

end;

procedure TFAssistExt.GDSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
  inherited;
  SavCol := ACol;
  SavRow := ARow;
end;

procedure TFAssistExt.RTypefichierClick(Sender: TObject);
begin
  inherited;
   if (RTypefichier.ItemIndex = 1) then
   begin
      TabResultat.TabVisible := True;
      TabDelimite.TabVisible := True;
      Label14.visible        := True;
      BValide.visible        := True;
   end
   else
   begin
      TabResultat.TabVisible := False;
      TabDelimite.TabVisible := False;
      Label14.visible        := False;
      BValide.Visible        := False;
   end;

end;


procedure TFAssistExt.DOMAINEExit(Sender: TObject);
var
{$IFDEF CISXPGI}
CodeRetour       : integer;
TobTemp          : TOB;
i                : integer;
Q                : TQuery;
{$ELSE}
TPar             : TQuery;
{$ENDIF}
ConnectionString : string;
begin
  inherited;
ListBox1.clear;
ListBox2.clear;
DomaineParam := RendDomaine(Domaine.Text);
{$IFDEF CISXPGI}
        if (DomaineParam = 'X') and (CComplement ='IMPORT') then
        begin
{$IFNDEF EAGLCLIENT}
             Q := OpenSQL ('SELECT * FROM DETABLES WHERE DT_DOMAINE="C" AND DT_NOMTABLE  like "%EEXBQ%" or '+
             ' DT_NOMTABLE like "%ETEBAC%" or  DT_NOMTABLE="BANQUECP"', TRUE);

             while not Q.EOF do
             begin
                 if ListBox1.items.IndexOf(Q.FindField ('DT_NOMTABLE').asstring) = -1 then
                      ListBox1.Items.Add(Q.FindField ('DT_NOMTABLE').asstring);
                 Q.Next;
             end;
             Ferme(Q);
{$ELSE}
             ChargementListeTable (ListBox1, DomaineParam);
{$ENDIF}
        end
        else
        begin
                if (CComplement ='EXPORT') then
                   CodeRetour := RendTobparametre (DomaineParam, TobTemp, 'COMPTA', 'PARAM', LePays)
                else
                CodeRetour := RendTobparametre (DomaineParam, TobTemp);
                if CodeRetour = -1 then
                begin
                    For i:= 0 to TobTemp.detail.count-1 do
                    begin
                                    if ListBox1.items.IndexOf(TobTemp.detail[i].Getvalue('TableName')) = -1 then
                                       ListBox1.Items.Add(TobTemp.detail[i].Getvalue('TableName'));
                    end;
                end;
                TobTemp.free;
        end;
{$ELSE}
(*TPar := TQuery.create(nil);
TPar.DatabaseName := DMImport.DBGlobal.Name;
TPar.SQL.Add ('Select * from '+DMImport.GzImpPar.TableName+
' Where Domaine="'+DomaineParam+'"');
TPar.Open;
*)
{$IFDEF DBXPRESS}
              ConnectionString := DMImport.DBGlobal.ConnectionString;
{$ELSE}
              ConnectionString := DMImport.DBGlobal.Name;
{$ENDIF}

TPar := OpenSQLADO ('Select * from '+DMImport.GzImpPar.TableName+
' Where Domaine="'+DomaineParam+'"', ConnectionString);

While not TPar.EOF do
begin
   if ListBox1.items.IndexOf(TPar.FieldByName('TableName').AsString) = -1 then
        ListBox1.Items.Add(TPar.FieldByName('TableName').AsString);
   TPar.next;
end;
TPar.free;
{$ENDIF}
end;

procedure TFAssistExt.BAjouterClick(Sender: TObject);
var
N         : integer;
begin
  inherited;
    for N:=ListBox1.Items.Count-1 downto 0 do
    begin
      if ListBox1.Selected[N] then
      begin
        ListBox2.Items.AddObject(ListBox1.Items[N], ListBox1.Items.Objects[N]);
        if (Nbexercice.Value = 1) and ((ListBox1.Items[N]='Exercice1') or  (ListBox1.Items[N]='Exercice2')) then
          Nbexercice.Value := Nbexercice.Value+1;
        ListBox1.Items.Delete(N);
      end;
    end;
end;


procedure TFAssistExt.BEnleverClick(Sender: TObject);
var
N  : integer;
begin
  inherited;
	for N:=ListBox2.Items.Count-1 downto 0 do begin
		if ListBox2.Selected[N] then begin
			ListBox1.Items.AddObject(ListBox2.Items[N], ListBox2.Items.Objects[N]);
			ListBox2.Items.Delete(N);
		end;
	end;

end;

procedure TFAssistExt.bPrecedentClick(Sender: TObject);
      Procedure     AfficheEtape;
      begin
        dec(PSuiv);
        if PSuiv <= 0 then PSuiv := 1;
        lEtap.Caption := Msg.Mess[0] + ' ' + IntToStr(PSuiv);
      end;
begin
  inherited;
(*  case P.ActivePageIndex of
  1 : Image.ImageIndex := 6;
  2 : Image.ImageIndex := P.ActivePageIndex;
  end;
*)
  if (RadioMode.ItemIndex = 0) then
  begin
     if ((P.ActivePageIndex = 3) or (P.ActivePageIndex = 2)) and (RTypefichier.ItemIndex = 0) then
        bPrecedentClick(Sender);
     if (P.ActivePage = TabParamsoc) and (OkDelim or (DomaineParam = 'O') or (DomaineParam = 'X'))then
     begin
           bPrecedentClick(Sender);  exit;
     end;
     if (P.ActivePage = Tabselect) and (DomaineParam = 'O') then
     begin
           bPrecedentClick(Sender);  exit;
     end;
  end;
  if ((P.ActivePageIndex = 1) or (P.ActivePageIndex = 2) or (P.ActivePageIndex = 3) or (P.ActivePageIndex = 4)) and (RadioMode.ItemIndex = 1) then
         bPrecedentClick(Sender);

  AfficheEtape;
end;

{$IFNDEF CISXPGI}
procedure TFAssistExt.ChargeSQL;
var
    SL, SLP : TStringList;
	  N       : Integer;
    Pass    : string;
    NN      : string;
begin
  inherited;
	cbOrigine.ItemIndex := 0;
	cbTypeSortie.ItemIndex := 2;
	SL := TStringList.Create;
	SLP := TStringList.Create;
	SL_origin := TStringList.Create;
	SL_ODBC := TStringList.Create;
{$IFNDEF DBXPRESS}
        Session.GetAliasNames(SL);
	for N:=0 to SL.Count-1 do
	begin
        NN := Session.GetAliasDriverName (SL[N]);
		Session.GetAliasParams(SL[N], SLP);
		// XXXXX if (SLP.Values['ODBC DSN'] <> '') and (NN = 'SQL Server') then
			cbAliasODBC.Items.Add(SL[N]);
	end;
{$ENDIF}
	SL_Origin.add('Parametre');
	SL_Origin.add('Donnees');
	SL_ODBC.Assign(cbAliasODBC.Items);
	SLP.Free;
	SL.Free;
	if cbAliasODBC.Items.Count > 0 then cbAliasODBC.ItemIndex := 0;
    ChargementComboDomaine (Domaine, FALSE);
    Pass := DayPass(Date);
    if V_PGI.PassWord=Pass then ParametrageCegid := TRUE
    else ParametrageCegid := FALSE;
end;
{$ENDIF}

procedure TFAssistExt.cbTypeSortieChange(Sender: TObject);
begin
  inherited;
	FTypeSortie := TTypeSortie(cbTypeSortie.ItemIndex);
	if FTypeSortie = tsODBC then
	begin
		FDestSortie := tsAliasODBC;
	end;
end;

procedure TFAssistExt.cbOrigineChange(Sender: TObject);
begin
  inherited;
	cbAliasODBC.Items.clear;
	if not (cbOrigine.ItemIndex <> 0) then
    begin
         edUtilisateur.Enabled := FALSE;
         edMotdepasse.Enabled := FALSE;
//	     cbaliasODBC.Items.Assign(SL_Origin);
    end
	else
    begin
         edUtilisateur.Enabled := TRUE;
         edMotdepasse.Enabled := TRUE;
         cbAliasODBC.Items.assign(SL_ODBC);
    end;
    cbAliasODBC.Itemindex := 0;
end;

{$IFDEF CISXPGI}
procedure TFAssistExt.BChercheClick;
begin
end;
{$ELSE}
procedure TFAssistExt.BChercheClick;
Var stSQL : String ;
    OkExec : boolean ;
    nb : integer ;
begin
SourisSablier ;
If Trim(uppercase(Copy(Z_SQL.text,1,6)))<>'SELECT' Then
if (Not V_PGI.SAV) or (v_pgi.Laserie<=S2) then
   BEGIN
   HM.Execute(0,'','') ; exit ;
   END ;
OkExec :=(Trim(uppercase(Copy(Z_SQL.text,1,6)))='UPDATE') or
         (Trim(uppercase(Copy(Z_SQL.text,1,6)))='DELETE') or
         (Trim(uppercase(Copy(Z_SQL.text,1,6)))='INSERT') ;
Q.Close ;
{$IFNDEF EAGLCLIENT}
StSQL:=Z_SQL.Lines.text ;
StSQL:=InsertWhereDPStd(StSQL) ;
{$ENDIF}
Q.SQL.Clear ;
Q.SQL.Add(StSQL) ;
Q.Requestlive:=FALSE ;

ChangeSQL(Q) ;
Try
if OkExec then
   BEGIN
   Q.ExecSQL ;
   nb:=Q.RowsAffected ;
   Q.Close ;
   HShowMessage('0;Moniteur SQL;Requête éxecutée : $$ enregistrements traités;I;O;O;O;','',IntToStr(Nb)) ;
   END else Q.OPEN ;
finally
end ;
SourisNormale ;
SelectRequete;
end;
{$ENDIF}

{$IFDEF EAGLCLIENT}
procedure TFAssistExt.SelectRequete;
var B                             : TBookmark;
	N                         : Integer;
	aFMTNumber, aNewFMTNumber : FMTNumber;
	FAliasName                : String;

begin
   if  FScriptRequete = nil then
       FScriptRequete := TScriptRequete.Create(Application);
   if DomaineParam = '' then DomaineParam := RendDomaine(Domaine.Text);
   FScriptRequete.Domaine := DomaineParam;
   with FScriptRequete do begin
        AliasName      := cbAliasODBC.Text;
        TypeSortie     := FTypeSortie;
        NomFichier     := Code.Text;
        DestSortie     := FDestSortie;
        ModeASCII      := TModeASCII(0); // ascii fixe
        ScriptSuivant  := EdReqSuivante.Text;
        AliasSortieODBC:= '';
        Utilisateur    := EdUtilisateur.text;
        Motdepasse     := EdMotdepasse.text;
        FComment       := EdComment.Text;
        RChemin        := Chemin.Text;
        DriverName     := cbOrigine.Text;
        ConstruitVariable(true);
   end;
	    AjouteVariableAlias(FScriptRequete.Variable);

    FScriptRequete.TypeSortie := tsASCII;

    if FScriptRequete.TypeSortie <> tsAucune then
    begin
         SelectRequeteCwas (FScriptRequete, SQ);
    end;
end;
{$ELSE}

procedure TFAssistExt.SelectRequete;
var
  B                         : TBookmark;
  {$IFDEF DBXPRESS}
	T                         : TADOTable;
  {$ELSE}
	T                         : TTable;
  {$ENDIF}
	N                         : Integer;
	aFMTNumber, aNewFMTNumber : FMTNumber;
	FAliasName                : String;
  Q                         : TQuery;
	procedure InitTable;
	begin
  {$IFDEF DBXPRESS}
                T.Connection := DMImport.DBGLOBALD;
  {$ELSE}
                FAliasName :=DMImport.DBGlobalD.DatabaseName ;
                T.DatabaseName :=FAliasName ;
   {$ENDIF}
		if FScriptRequete.TypeSortie = tsParadox then
		begin // Sortie Paradox
			T.TableName := ExtractFileName(FScriptRequete.NomFichier);
 {$IFNDEF DBXPRESS}
			T.TableType := ttParadox;
 {$ENDIF}
		end
		else if FScriptRequete.TypeSortie = tsODBC then
		begin // sortie ODBC
			T.TableName := ExtractFileName(FScriptRequete.NomFichier);
			FScriptRequete.DestSortie := tsAliasODBC;
 			FScriptRequete.BdeToODBC(Q, T);
			exit;
		end
		else
		begin  // sortie Ascii
                        if (ParamCount > 0) and (GetInfoVHCX.NomFichier <> '') then
                           T.TableName := GetInfoVHCX.NomFichier
                        else
			    T.TableName := ExtractFileName(FScriptRequete.NomFichier);
 {$IFNDEF DBXPRESS}
			T.TableType := ttASCII;
 {$ENDIF}
		end;
	end;
begin
   B := nil;
   if  FScriptRequete = nil then
       FScriptRequete := TScriptRequete.Create(Application);
   if DomaineParam = '' then DomaineParam := RendDomaine(Domaine.Text);
   FScriptRequete.Domaine := DomaineParam;
   with FScriptRequete do begin
        AliasName      := cbAliasODBC.Text;
        TypeSortie     := FTypeSortie;
        NomFichier     := Code.Text;
        DestSortie     := FDestSortie;
        ModeASCII      := TModeASCII(0); // ascii fixe
        ScriptSuivant  := EdReqSuivante.Text;
        AliasSortieODBC:= '';
        Utilisateur    := EdUtilisateur.text;
        Motdepasse     := EdMotdepasse.text;
        FComment       := EdComment.Text;
        RChemin        := Chemin.Text;
        DriverName     := cbOrigine.Text;
        ConstruitVariable(true);
   end;
	    AjouteVariableAlias(FScriptRequete.Variable);

    FScriptRequete.TypeSortie := tsASCII;

    if FScriptRequete.TypeSortie <> tsAucune then
    begin
		  DbiGetNumberFormat(aFMTNumber);	// modifiaction du separateur de decimal
		  aNewFMTNumber := aFMTNumber;		//  DecimalSeparator ne marche pas
		  aNewFMTNumber.cDecimalSeparator := '.';	// pour les tables ASCII
		  DbiSetNumberFormat(aNewFMTNumber);
		  T := TTable.Create(nil);
		  try
                          B := Q.GetBookmark;
			  InitTable;
                          if not T.Active then
				  with Q.FieldDefs do
					for N:=0 to Count-1 do
						with Items[N] do
						begin
						  if DataType = ftMemo then
							  T.FieldDefs.Add(Name, ftString, 1024, false)
						  else if (FScriptRequete.TypeSortie = tsODBC) and (DataType = ftCurrency) then
							  T.FieldDefs.Add(Name, DetermineTypeODBC(ftCurrency, FScriptRequete.AliasSortieODBC), Size, false)
						  else
							  T.FieldDefs.Add(Name, DataType, Size, false);
						end;

 {$IFNDEF DBXPRESS}
			  if not T.Active then
			  	T.CreateTable;
 {$ENDIF}
			  Sleep(2000); // suite au pb avec Le fichier .SCH
			  if FScriptRequete.TypeSortie <> tsODBC then
			  begin
 {$IFNDEF DBXPRESS}
 				if FScriptRequete.TypeSortie = tsParadox then
				   T.EmptyTable
				else
{$ENDIF}
        if FScriptRequete.ModeAscii = maDelimite then   // TCurrencyField
				begin
					ModifierSchema(FAliasName, ExtractFileName(FScriptRequete.NomFichier));
				end;
			  end;
  			  try
				T.Open;
				SQ.dataset := nil ;
				while not Q.Eof do
				begin
				  T.Append;
				  for N:=0 to Q.FieldCount-1 do
						T.Fields[N].Value := Q.Fields[N].Value;
				  T.Post;
				  Q.Next;
				end;
			  finally Q.GotoBookmark(B);  end;
			  SQ.dataset := Q;
		  finally
			  Q.FreeBookmark(B);
			  T.Free;
			  DbiSetNumberFormat(aFMTNumber);
		  end;
	  end;
end;
{$ENDIF}

function TFAssistExt.DetermineTypeODBC(Typ : TFieldType;sDsn : string):TFieldType;
begin
	Result := Typ;
	if typ = ftCurrency then
		result := Typ;
end;

{$IFNDEF CISXPGI}
procedure TFAssistExt.BValiderClickRequete;
var
    AStream, AStreamTable    : TStream;
    ABlobField, ATblCorField : TBlobField;
{$IFDEF  DBXPRESS}
    tbSauver                 : TADOTable;
{$ELSE}
    tbSauver                 : TTable;
{$ENDIF}
    bTrouve                  : Boolean;
    S                        : String;
begin
  inherited;
      tbSauver := DMImport.GzImpReq;
{$IFDEF  DBXPRESS}
      DMImport.Cmd.CommandText := 'SELECT * FROM PGZIMPREQ WHERE CLE="'+ FScriptRequete.NomFichier +'"';
      DMImport.Cmd.Execute;
{$ENDIF}

      if tbSauver.active then tbSauver.close;
      tbSauver.Open;
      if  FScriptRequete = nil then
       FScriptRequete := TScriptRequete.Create(Application);
      if DomaineParam = '' then DomaineParam := RendDomaine(Domaine.Text);
      FScriptRequete.Domaine := DomaineParam;

      with FScriptRequete do begin
           AliasName  := cbAliasODBC.Text;
           TypeSortie := FTypeSortie;
           NomFichier := Code.Text;
           DestSortie := FDestSortie;
           ModeASCII  := TModeASCII(0); // ascii fixe
           ScriptSuivant := EdReqSuivante.Text;
           AliasSortieODBC:= '';
           Utilisateur := EdUtilisateur.text;
           Motdepasse := EdMotdepasse.text;
           FComment   := EdComment.Text;
           Name       := Code.text;
           Rchemin    := Chemin.Text;
           DriverName     := cbOrigine.Text;
           AssignSql(Z_SQL);
{$IFNDEF  DBXPRESS}
           bTrouve := tbSauver.FindKey([NomFichier]);
           if bTrouve then
{$ELSE}
           if not tbSauver.Eof then
{$ENDIF}
           begin
              if tbSauver.Fields[5].value  <> 'SQL'  then
              begin
              PGIBox('Au niveau Conception, un Script '+ NomFichier + ' existe déjà, changez le nom du code de votre requête', 'Conception') ;
              tbSauver.Close;
              exit;
              end;
			{ La requête %s existe déjà  Voulez-vous la remplacer ? }
              if CommentC = taCreat then
              begin
                   if (PGIAsk('La Requête Existe déjà, voulez-vous néanmoins enregistrer la requête ?',
                   'Conception')<> mryes) then  begin tbSauver.Close; exit; end;
              end
              else
              begin
                   S := Format('Confirmez-vous les modifications de la requête %s ?', [Name]);
                   if (PGIAsk(S, 'Requête')<> mryes) then  begin tbSauver.Close; exit; end;
              end;
            end;
           if ParametrageCegid then
               NiveauAcces := 0
           else
               NiveauAcces := 1;

           Name := NomFichier;
           Destination := poGlobal;

           with tbSauver do begin
		   if Active then
		      Close;
{$IFDEF  DBXPRESS}
          DMImport.Cmd.CommandText := 'SELECT * FROM PGZIMPREQ WHERE CLE="'+ Code.text +'"';
          DMImport.Cmd.Execute;
{$ENDIF}
		      Open;
		      IndexName := '';	{ selection de l'index primaire }
{$IFNDEF  DBXPRESS}
		      if FindKey([Code.text]) then Edit
		      else Insert;
{$ELSE}
          if not EOF then Edit
          else Insert;
{$ENDIF}
		      ABlobField := FieldByName('PARAMETRES') as TBlobField;
		      ATblCorField := FieldByName('TBLCOR') as TBlobField;
		      FieldByName('CLE').AsString := NomFichier;
		      FieldByName('COMMENT').AsString := FComment;
		      FieldByName('CLEPAR').AsString := 'SQL';
		      FieldByName('MODIFIABLE').AsInteger := NiveauAcces;
                      FieldByName('DOMAINE').asstring := Domaine;
                      FieldByName('DATEDEMODIF').asdatetime := now;
		      AStream := TBlobStream.Create(ABlobField, bmWrite);
		      AStreamTable := TBlobStream.Create(ATblCorField, bmWrite);
		      FScriptRequete.SaveTo(AStream, AStreamTable);
		      AStreamTable.Free;
		      AStream.Free;
		      Post;
		      Close;
           end;
	  end;
end;
{$ENDIF}

procedure TFAssistExt.EdUtilisateurExit(Sender: TObject);
begin
  inherited;
  FUtilisateur := EdUtilisateur.text;
end;

procedure TFAssistExt.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
LibereDelim;
end;

procedure TFAssistExt.FormShow(Sender: TObject);
var
val,Lib : string;
begin
  inherited;
    OkFin := FALSE;
    PSuiv := 1;
    if Indexcb <> -1 then cbDelimChamp.ItemIndex := Indexcb
    else cbDelimChamp.Text := CarDelim;
    if OkRequete then RadioMode.ItemIndex := 1
    else RadioMode.ItemIndex := 0;
    RadioMode.Enabled := FALSE;
    ComboDevise.itemindex := 0;
    EdNature.Text := CComplement;
    if ParametrageCegid and (CommentC = tacreat) and (RadioMode.ItemIndex = 0) then
    begin
         EdNature.Text := 'STD CEGID';
         EdNature.Enabled := FALSE;
    end;
    if RadioMode.ItemIndex = 0  then  HLabel3.caption := 'de script'
    else
    begin
      Caption := 'Assistant création d''extraction SGBD';
//    if CommentC = taModif then
      HLabel1.caption := 'Bienvenue dans l''assistant de création ';
      HLabel3.caption := 'd''extraction SGBD';
    end;
    if OkRequete then
    begin
      edNature.Visible       := FALSE;
      TedNature.visible      := FALSE;
      TypeTransfert.visible  := FALSE;
      TTypeTransfert.visible := FALSE;
      RTypefichier.visible   := FALSE;
      ModeParam.visible      := FALSE;
      Filename.visible       := FALSE;
      LFichier.visible       := FALSE;
      cbTypeCar.visible      := FALSE;
      Label15.visible        := FALSE;
      BVoir.visible          := FALSE;
      GroupBox11.visible     := FALSE;
      if CommentC = taModif then
      begin
        Domaine.itemindex := Domaine.items.IndexOf(RendLibelleDomaine(FScriptRequete.Domaine));
        with FScriptRequete do
        begin
          cbAliasODBC.ItemIndex := cbAliasODBC.Items.IndexOf(AliasName);
          if cbAliasODBC.ItemIndex = -1 then cbAliasODBC.Text := ALiasName;
          FTypeSortie := TypeSortie;
          Code.Text   := NomFichier;
          FDestSortie := DestSortie;
          EdReqSuivante.Text := ScriptSuivant;
          EdUtilisateur.text := Utilisateur;
          EdMotdepasse.text := Motdepasse;
          EdComment.Text := FComment;
          Code.text := Name;
          Chemin.text := RChemin;
          cbOrigine.Text := DriverName;
          //Z_SQL.Text := ExtraireMemo(UPPERCASE(DropComment(Requete.Text)));
          AssignSqlMemo (Z_SQL, Requete);
        end;
	cbOrigine.ItemIndex := 0;
        val := FScriptRequete.AliasName;

        if (FScriptRequete.AliasName <> '') then
        begin	// Alias
          if SL_origin.indexof(val) <> -1 then // non paradox donc ODBC
          begin
	    cbOrigine.itemIndex := 0;
	    cbAliasODBC.Items.clear;
	    cbAliasODBC.Items.Assign(SL_origin);
          end
          else
          begin
            //cbOrigine.ItemIndex := 1;
            cbOrigine.ItemIndex := cbOrigine.Items.IndexOf(FScriptRequete.DriverName);
            if cbOrigine.ItemIndex = -1 then
              cbOrigine.Text := FScriptRequete.DriverName;

            cbAliasODBC.Items.clear;
            cbAliasODBC.Items.Assign(SL_ODBC);
          end;

          cbAliasODBC.ItemIndex := cbAliasODBC.Items.IndexOf(val);
	  cbAliasODBC.Enabled   := true;
	  lblAliasODBC.Enabled  := true;
        end;
      end;
    end;

    if (GetInfoVHCX.Mode = 'I') then
    begin
      Lib := RendLibelleDomaine(GetInfoVHCX.Domaine);
      Domaine.itemindex := Domaine.items.IndexOf(Lib);
      DOMAINEChange(Sender);
      Domaine.Enabled := FALSE;
    end
    else
           if CommentC = taCreat then   Domaine.itemindex := 2;

    if ModifInterface then
    begin
      Domaine.itemindex       := Domaine.items.IndexOf(RendLibelleDomaine(DomaineParam));
      DOMAINEChange (Sender);
      Typetransfert.itemindex := Typetransfert.items.IndexOf(NtTrans);
      Code.Enabled            := FALSE;
      edComment.Enabled       := FALSE;
      EdNature.Enabled        := FALSE;
      Typetransfert.Enabled   := FALSE;
      Editeur.Enabled         := FALSE;
      ModeParam.Enabled       := FALSE;
      CbTypeCar.Enabled       := FALSE;
      RTypefichier.ItemIndex  := 1;
    end;
end;

procedure TFAssistExt.CheminElipsisClick(Sender: TObject);
begin
  inherited;
    with Topendialog.create(Self) do
    begin
         Execute;
         Chemin.Text :=  FileName;
         Free;
    end;
    SetCurrentDirectory(PChar(Chemin.Text));

end;

procedure TFAssistExt.ChampfinChange(Sender: TObject);
begin
  inherited;
  if not ModifInterface then
     GRILVISU.Row := Champfin.value;
end;

procedure TFAssistExt.ChampDebutChange(Sender: TObject);
begin
  inherited;
  if not ModifInterface then
     GRILVISU.Row := ChampDebut.value;
end;

procedure TFAssistExt.ListBox2DragDrop(Sender, Source: TObject; X,
  Y: Integer);
var
  index, index2 : integer;
  APoint        : TPoint;
begin
  inherited;
  APoint.X := X;
  APoint.Y := Y;
	if Source is TListBox then
	begin
	  Index := TListBox(Source).ItemIndex;
          Index2 := ListBox2.ItemAtPos(APoint, True);
	  ClipBoard.SetTextBuf(PChar(TListBox(Source).Items[index]));
          ListBox2.Items.Move(index, index2);
	end

end;

procedure TFAssistExt.ListBox2DragOver(Sender, Source: TObject; X,
  Y: Integer; State: TDragState; var Accept: Boolean);
begin
  inherited;
  Accept := TRUE;
end;

procedure TFAssistExt.RenseigneTobParam(TPControle : TSVControle);
var
T1         : TOB;
Codedevise : string;
Typet      : string;
iet        : integer;
            procedure returnCodedevise;
            var
            ie         : integer;
            begin
                 Codedevise := 'EUR';
                 for ie := 0 to lgdevise do
                 begin
                         if Correspdevise[ie].libelle = ComboDevise.Text then
                         begin Codedevise := Correspdevise[ie].code; break; end;
                 end;
            end;
begin
            if TPControle.TOBParam = nil then exit;
            if DomaineParam = 'C' then // domaine compta  PGI
            begin
                 returnCodedevise;
                 T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Param5', 'devise'], FALSE);
                 if T1 <> nil then
                 begin
                     T1.Putvalue ('Contenu', Codedevise);
                     T1.Putvalue ('Fige', Modeparam.Itemindex);
                 end;
                 T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Ecriture', 'codedevise'], FALSE);
                 if T1 <> nil then
                 begin
                     T1.Putvalue ('Contenu', Codedevise);
                     T1.Putvalue ('Fige', Modeparam.Itemindex);
                 end;
                 T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Ecriture', 'codemontant'], FALSE);
                 if T1 <> nil then
                 begin
                        if Codedevise = 'EUR' then T1.Putvalue ('Contenu', 'E--')
                        else
                        if Codedevise = 'FRF' then T1.Putvalue ('Contenu', 'F--')
                        else
                        T1.Putvalue ('Contenu', 'D--');
                 end;

                 T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Param5', 'tenueeuro'], FALSE);
                 if T1 <> nil then
                 begin
                     if Codedevise = 'EUR' then T1.Putvalue ('Contenu', 'X')
                     else T1.Putvalue ('Contenu', '-');
                     T1.Putvalue ('Fige', Modeparam.Itemindex);
                 end;
                 T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Param5', 'decimale'], FALSE);
                 if T1 <> nil then
                 begin
                     T1.Putvalue ('Contenu', Nbdecimale.value);
                     T1.Putvalue ('Fige', Modeparam.Itemindex);
                 end;
                 T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Param5', 'pariteeuro'], FALSE);
                 if T1 <> nil then
                 begin
                     T1.Putvalue ('Contenu', PariteEuro.text);
                     T1.Putvalue ('Fige', Modeparam.Itemindex);
                 end;
                 T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Ecriture', 'tauxdevise'], FALSE);
                 if T1 <> nil then
                 begin
                     T1.Putvalue ('Contenu', PariteEuro.text);
                     T1.Putvalue ('Fige', Modeparam.Itemindex);
                 end;
                 T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Param5', 'analytique'], FALSE);
                 if T1 <> nil then
                 begin
                     T1.Putvalue ('Contenu', Analytique.checked);
                     T1.Putvalue ('Fige', Modeparam.Itemindex);
                 end;
                 if (not COMPAUX.checked) then
                 begin
                        T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Ecriture', 'auxiliaire'], FALSE);
                        if T1 <> nil then
                           T1.Putvalue ('Obligatoire', 'N');
                        T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Ecriture', 'natureligne'], FALSE);
                        if T1 <> nil then
                           T1.Putvalue ('Obligatoire', 'N');
                 end;

//                 if Modeparam.Itemindex = 1 then
                 begin
                     T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Param2', 'lgcomptegene'], FALSE);
                     if T1 <> nil then
                        T1.Putvalue ('Contenu', LGgene.value);
                     T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Param2', 'lgcptetiers'], FALSE);
                     if T1 <> nil then
                        T1.Putvalue ('Contenu', LGauxi.value);
                     T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Param2', 'lgaxe1'], FALSE);
                     if T1 <> nil then
                        T1.Putvalue ('Contenu', LGaxe1.value);

                     for iet := 0 to TPControle.TOBParam.Detail.count-1 do
                     begin
                          T1 := TPControle.TOBParam.detail[iet];
                          if ((T1.GetValue ('TableName')= 'Entete')and (T1.GetValue ('Nom')<> 'dossier'))
                          or (T1.GetValue ('TableName')= 'Param2')
                          or (T1.GetValue ('TableName')= 'Param3')
                          or (T1.GetValue ('TableName')= 'Param5')
                          or (T1.GetValue ('TableName')= 'Etabliss')
                          or ((T1.GetValue ('TableName')= 'Generaux') and (T1.GetValue ('Nom')= 'code'))
                          or ((T1.GetValue ('TableName')= 'Ecriture') and (T1.GetValue ('Nom')= 'etablissement'))
                          or ((T1.GetValue ('TableName')= 'Ecriture') and (T1.GetValue ('Nom')= 'naturemouvement'))
                          or ((T1.GetValue ('TableName')= 'Generaux') and (T1.GetValue ('Nom')= 'axe2'))
                          or ((T1.GetValue ('TableName')= 'Generaux') and (T1.GetValue ('Nom')= 'axe3'))
                          or ((T1.GetValue ('TableName')= 'Generaux') and (T1.GetValue ('Nom')= 'axe4'))
                          or ((T1.GetValue ('TableName')= 'Generaux') and (T1.GetValue ('Nom')= 'axe5'))
                          then T1.Putvalue ('Fige', Modeparam.Itemindex);
                     end;

                 end;
            end;
            if DomaineParam = 'S' then // compta  SISCOII
            begin
                 returnCodedevise;

                 T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Ardossier', 'monnaie'], FALSE);
                 if T1 <> nil then
                 begin
                     if (Codedevise <> 'EUR') and (Codedevise <> 'FRF') then Codedevise := 'EUR';
                     T1.Putvalue ('Contenu', Codedevise);
                     T1.Putvalue ('Fige', Modeparam.Itemindex);
                 end;
                 T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Ardossier', 'lgcpt'], FALSE);
                 if T1 <> nil then
                 begin
                        T1.Putvalue ('Contenu', FormatFloat('00',LGgene.value));
                        T1.Putvalue ('Fige', Modeparam.Itemindex);
                 end;
                 if Typetransfert.Text = 'Dossier' then Typet := 'E' else
                 if Typetransfert.Text = 'Journal' then Typet := 'J' else
                 if Typetransfert.Text = 'Balance' then Typet := 'B' else
                 Typet := 'E';

                 T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Ardossier', 'archivage'], FALSE);
                 if T1 <> nil then
                 begin
                     T1.Putvalue ('Contenu', Typet);
                     T1.Putvalue ('Fige', Modeparam.Itemindex);
                 end;
                 T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Ardossier', 'premtp'], FALSE);
                 if T1 <> nil then
                 begin
                     T1.Putvalue ('Contenu', Nbdecimale.value);
                     T1.Putvalue ('Fige', Modeparam.Itemindex);
                 end;

            end;
            if DomaineParam = 'M' then // domaine IMOII
            begin
                 returnCodedevise;
                 T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Param', 'monnaie'], FALSE);
                 if T1 <> nil then
                 begin
                     if (Codedevise <> 'EUR') and (Codedevise <> 'FRF') then Codedevise := 'E';
                     T1.Putvalue ('Contenu', Codedevise[1]);
                     T1.Putvalue ('Fige', Modeparam.Itemindex);
                 end;
                 T1 := TPControle.TOBParam.FindFirst(['Domaine','TableName','Nom'], [DomaineParam, 'Param', 'lgcompte'], FALSE);
                 if T1 <> nil then
                 begin
                        T1.Putvalue ('Contenu', FormatFloat('00000',LGgene.value));
                        T1.Putvalue ('Fige', Modeparam.Itemindex);
                 end;
            end;

end;

procedure TFAssistExt.EnableTabparamsoc;
begin
if (DomaineParam <> 'S') and  (DomaineParam <> 'C')then
begin
     GroupBox10.Enabled := FALSE; GroupBox7.Enabled := FALSE; GroupBox7.Enabled := FALSE; exit;
end
else
begin
     GroupBox10.Enabled := TRUE; GroupBox7.Enabled := TRUE; GroupBox7.Enabled := TRUE;
end;

if DomaineParam = 'S' then
begin
     Label8.Enabled := FALSE;
     LGAuxi.Enabled := FALSE;
     Label9.Enabled := FALSE;
     LGAxe1.Enabled := FALSE;
     Analytique.Enabled  := FALSE;
     Label6.Enabled      := FALSE;
     PariteEuro.Enabled  := FALSE;
     Statistique.Enabled := TRUE;
     NbExercice.Enabled  := FALSE;
     Label13.Enabled     := FALSE;
end
else
if DomaineParam = 'C' then
begin
     Label8.Enabled := TRUE;
     LGAuxi.Enabled := TRUE;
     Label9.Enabled := TRUE;
     LGAxe1.Enabled := TRUE;
     Analytique.Enabled := TRUE;
     Label6.Enabled := TRUE;
     PariteEuro.Enabled := TRUE;
     Statistique.Enabled := FALSE;
     NbExercice.Enabled  := TRUE;
     Label13.Enabled     := TRUE;
end;

end;

procedure TFAssistExt.DOMAINEChange(Sender: TObject);
begin
  inherited;
DomaineParam := RendDomaine(Domaine.Text);
// pour domaine conso
if (RadioMode.ItemIndex = 0) then
begin
    if (DomaineParam = 'O') then   // domaine Etafi Conso
    begin
      Typetransfert.clear;
      Typetransfert.Items.add ('<Tous>');
      Typetransfert.Items.add ('Balance');
      Typetransfert.Items.add ('Intra-Groupe');
      Typetransfert.Items.add ('Ecriture');
      Modeparam.Itemindex     := 1;
    end
    else
    begin
      Modeparam.Itemindex     := 0;
       Typetransfert.clear;
       Typetransfert.Items.add ('Dossier');
       Typetransfert.Items.add ('Journal');
       Typetransfert.Items.add ('Balance');
       Typetransfert.Items.add ('Synchronisation');
    end;
end;

end;


procedure TFAssistExt.BValideClick(Sender: TObject);
//var
//Command,Rep    : string;
//Executable : string;
begin
  inherited;
(*  if pos('.exe', EdShellExecute.Text) <> 0 then
  begin
       Executable := Copy(EdShellExecute.Text, 0, pos('.exe', EdShellExecute.Text)+3);
       Command    := Copy(EdShellExecute.Text, pos('.exe', EdShellExecute.Text)+4, length(EdShellExecute.Text));
       ShellExecute(0, PCHAR('open'),PCHAR(Executable),PCHAR(Command),Nil,SW_RESTORE);
  end
  else
  if pos('.bat', EdShellExecute.Text) <> 0 then
  begin
       Executable := Copy(EdShellExecute.Text, 0, pos('.bat', EdShellExecute.Text)+3);
       Command    := Copy(EdShellExecute.Text, pos('.bat', EdShellExecute.Text)+4, length(EdShellExecute.Text));
       ShellExecute(0, PCHAR('open'),PCHAR(Executable),PCHAR(Command),Nil,SW_RESTORE);
  end
  else
*)
//ShellExecute(0, PCHAR('open'),PCHAR('C:\tmp\00002\AppelDos.bat'),PCHAR('C:\tmp\00002\Fich1.txt C:\tmp\00002\fich2.txt C:\tmp\00002\ici.txt'),Nil,SW_RESTORE);
  ExecuteShell(EdShellExecute.Text);
end;

{$IFDEF EAGLCLIENT}
procedure TFAssistExt.SauvegardeFiltre;
begin
end;
{$ELSE}
{$IFDEF DBXPRESS}
procedure TFAssistExt.SauvegardeFiltre;
begin
end;
{$ELSE}
procedure TFAssistExt.SauvegardeFiltre;
var
FScriptDel             : TScriptDelimite;
{$IFNDEF CISXPGI}
tbSauver               : TTable;
{$ELSE}
tbSauver               : TQuery;
{$ENDIF}
AChamp                 : TChampParam;
N,NiveauAcces          : integer;
St                     : string;
AStream                : TStream;
ABlobField             : TBlobField;
begin
     FScriptDel := TScriptDelimite.Create(nil);
     FScriptDel.Name := Code.Text;
{$IFNDEF CISXPGI}
     tbSauver := DMImport.GzImpDelim;
{$ELSE}
        tbSauver := OpenSQl ('SELECT * FROM CPGZIMPREQ Where CIS_CLE="'+FScriptDel.Name+'" AND CIS_NATURE="X"', FALSE);
{$ENDIF}
     if ParametrageCegid then NiveauAcces := 0
     else NiveauAcces := 1;

	with tbSauver do
	begin

{$IFNDEF CISXPGI}
		if Active then
                Close;
		Open;
                if FindKey([FScriptDel.Name]) then
                begin
                        if (not ParametrageCegid) and (FieldByName('MODIFIABLE').AsInteger = 0)
                        and ModifInterface then
                        begin
                             PGIBox ('Il est impossible de modifier un stdandard CEGID', 'Conception');
                             Close; exit;
                        end;
                        St := Format('Le Script %s existe déjà, confirmez-vous sa modification  ?', [FScriptDel.Name]);
                        if (PGIAsk(St, 'Conception')<> mryes) then begin Close; exit; end;
                        Edit;
                end
                else Insert;
{$ELSE}
                if not tbsauver.EOF then
                    Edit
		else Insert;
{$ENDIF}
                FScriptDel.NomFichier :=  FileName.Text;
                FScriptDel.NomSortie  :=  Fichesortie.Text;
                FScriptDel.Domaine    :=  DomaineParam;
                FScriptDel.Libelle    :=  edcomment.Text;
                FScriptDel.Shellexec  :=  EdShellExecute.Text;
                FScriptDel.Chainref   :=  Chaineref.Text;
                FScriptDel.Longref    :=  Champref.value;
                FScriptDel.Londeb     :=  Champdebut.value;
                FScriptDel.Lonfin     :=  Champfin.value;
                FScriptDel.Chainelim  :=  CARELIM.Text;
                FScriptDel.Delimlib   :=  DELIMLIB.Text;
                if (cbDelimChamp.ItemIndex <> -1) and (cbDelimChamp.ItemIndex < 5) then
                FScriptDel.Cbdelim  := CarSepField[cbDelimChamp.ItemIndex]
                else
                begin
                     if cbDelimChamp.ItemIndex = 5 then
                        FScriptDel.Cbdelim := Car.text
                     else
                        FScriptDel.Cbdelim  := cbDelimChamp.Text[1];
                end;
	        for N:=1 to GD.RowCount-1 do
	        begin
                     if GD.Cells[1, N] = '' then exit;
                     AChamp               := TChampParam.Create(FScriptDel.Champ);
                     AChamp.Name         := GD.Cells[0, N];
                     AChamp.Longueur     := StrToint(GD.Cells[1, N]);
                     AChamp.Visible      := (GD.Cells[2, N]='X');
                     AChamp.Separateur   := (GD.Cells[3, N]='X');
                     AChamp.Aligne       := (GD.Cells[4, N]='X');

                end;
{$IFNDEF CISXPGI}
                ABlobField := FieldByName('PARAMETRES') as TBlobField;
                FieldByName('CLE').AsString := FScriptDel.Name;
                FieldByName('COMMENT').AsString := FScriptDel.Libelle;
                FieldByName('CLEPAR').AsString  := edNature.Text;
                FieldByName('MODIFIABLE').AsInteger := NiveauAcces;
                FieldByName('DOMAINE').asstring := DomaineParam;
                FieldByName('DATEDEMODIF').asdatetime := now;
{$ELSE}
                ABlobField := FieldByName('CIS_PARAMETRES') as TBlobField;
                FieldByName('CIS_CLE').AsString := FScriptDel.Name;
                FieldByName('CIS_COMMENT').AsString := FScriptDel.Libelle;
                FieldByName('CIS_CLEPAR').AsString  := edNature.Text;
                FieldByName('CIS_DOMAINE').asstring := DomaineParam;
                FieldByName('CIS_DATEMODIF').asdatetime := now;
                FieldByName ('CIS_Table1').Asstring := Editeur.text;
                FieldByName ('CIS_Table2').Asstring := Typetransfert.text;
                FieldByName('CIS_Table3').asstring := LePays;
                FieldByName('CIS_Table4').asstring := '';
                FieldByName('CIS_Table5').asstring := '';
                FieldByName('CIS_Table6').asstring := '';
                FieldByName('CIS_Table7').asstring := '';
                FieldByName('CIS_Table8').asstring := '';
                FieldByName('CIS_Table9').asstring := '';
                FieldByName('CIS_NATURE').asstring      := 'X';
{$ENDIF}
		AStream := TBlobStream.Create(ABlobField, bmWrite);
		FScriptDel.SaveDelimTo(AStream);
		AStream.Free;
		Post;
{$IFNDEF CISXPGI}
		Close;
{$ELSE}
                Ferme (tbSauver);
{$ENDIF}
	end;

     FScriptDel.free;
end;
{$ENDIF}
{$ENDIF}

function TFAssistExt.ChargeDelim(Codedelim : string; FichEnt: string=''; FichS : string='') : Boolean;
var
{$IFDEF CISXPGI}
  tbcharger              : TQuery;
{$ELSE}
  {$IFNDEF DBXPRESS}
  tbcharger              : TTable;
  {$ELSE}
  tbcharger              : TADOTable;
  {$ENDIF}
{$ENDIF}
  FScriptDel             : TScriptDelimite;
  ABlobField             : TField;
  Tb,TF                  : TOB;
  Stream1                : TmemoryStream;
  N                      : integer;
            procedure EnabChamp (ENB : Boolean);
            begin
                 EdShellExecute.Enabled := ENB;
                 FILENAME.Enabled       := ENB;
                 Rtypefichier.Enabled   := ENB;
                 ModeParam.Enabled      := ENB;
                 FicheSortie.Enabled    := ENB;
                 ChampRef.Enabled       := ENB;
                 ChaineRef.Enabled      := ENB;
                 ChampDebut.Enabled     := ENB;
                 Champfin.Enabled       := ENB;
                 CheckBoxTitre.Enabled  := ENB;
                 GD.Enabled             := ENB;
                 ChargementDelim.Enabled:= ENB;
                 Chaineref.Enabled      := ENB;
            end;
begin
        Result := FALSE; FScriptDel := nil; Stream1 := nil;
{$IFDEF CISXPGI}
        tbCharger := OpenSQl ('SELECT * FROM CPGZIMPREQ Where CIS_CLE="'+Codedelim+'"' +
        ' AND CIS_NATURE="X"' , TRUE);
        if tbCharger.EOF then
        begin
                     Ferme (tbCharger);
                     exit;
        end;
{$ELSE}
        tbcharger := DMImport.GzImpDelim;
        {$IFDEF  DBXPRESS}
              DMImport.Cmd.CommandText := 'SELECT * FROM PGZIMPREQ WHERE CLE="'+ Codedelim +'"';
              DMImport.Cmd.Execute;
        {$ENDIF}

{$ENDIF}

        RTypefichier.ItemIndex := 1;
	with tbcharger do
	begin
{$IFNDEF CISXPGI}
            if not Active then begin try Open; except Close; Exit; end; end;
{$IFDEF  DBXPRESS}
            if not EOF
{$ELSE}
                if (not FindKey([Codedelim])) and
                (not FindKey([UpperCase(Codedelim)]))  and
                (not FindKey([LowerCase(Codedelim)]))
{$ENDIF}
                 then
                 begin
                   if FichS = '' then
                     PGIBOX('Le Script '+Codedelim+' n''a pas de requête délimitée', 'Modification Filtre');
                   Close;
                   Exit;
                 end;
{$ENDIF}
                if FichEnt = '' then
                begin
{$IFNDEF CISXPGI}
                     if (not ParametrageCegid) and (FieldByName('MODIFIABLE').AsInteger = 0)
                     and ModifInterface then EnabChamp (FALSE)
                     else EnabChamp (TRUE);
{$ELSE}
                     EnabChamp (TRUE);
{$ENDIF}
                end;
            try
{$IFDEF CISXPGI}
                ABlobField := FindField('CIS_PARAMETRES');
{$ELSE}
                ABlobField := FieldByName('PARAMETRES');
{$ENDIF}
                Stream1 := TmemoryStream.create;
                TBlobField(ABlobField).SaveToStream (Stream1);
                Stream1.Seek (0,0);
                FScriptDel := LoadScriptFromStreamDelim(Stream1);
{$IFNDEF COMPTA}
                if not DirectoryExists(CurrentDossier+'\sauvegarde') then
                        CreateDir(CurrentDossier+'\sauvegarde');
                Stream1.SaveToFile (CurrentDossier+'\sauvegarde\'+Codedelim+'Delimite.txt');
{$ENDIF}
            except

                if (PGIAsk('Des incohérences dans le Script délimité, voulez-vous remonter la sauvegarde : ?',
                'Conception')= mryes) then
                begin
                      Stream1.free;
                      Stream1 := TmemoryStream.create;
                      Stream1.LoadFromFile (CurrentDossier+'\sauvegarde\'+Codedelim+'Delimite.txt');
                      Stream1.Seek (0,0);
                      FScriptDel := LoadScriptFromStreamDelim(Stream1);
                      Stream1.Free;  Stream1 := nil;
                end
                else
                begin
                      Stream1.Free;  FScriptDel.Destroy;
                      Close;
                      exit;
                end;
            end;
                Code.Text             := Codedelim;
                if FichEnt = '' then
                   FileName.Text      := FScriptDel.NomFichier
                else
                   FileName.Text      := FichEnt;
                if FichS = '' then
                   Fichesortie.Text   := FScriptDel.NomSortie
                else
                   Fichesortie.Text   := FichS;
                if FScriptDel.Domaine = '' then FScriptDel.Domaine := DomaineParam;
                Domaine.itemindex     := Domaine.items.IndexOf(RendLibelleDomaine(FScriptDel.Domaine));
                edcomment.Text        := FScriptDel.Libelle;
                EdShellExecute.Text   := FScriptDel.Shellexec;
                Chaineref.Text        := FScriptDel.Chainref;
                Champref.value        := FScriptDel.Longref;
                Carelim.Text          := FScriptDel.Chainelim;
                DELIMLIB.Text         := FScriptDel.Delimlib;
                cbDelimChamp.Text     := FScriptDel.Cbdelim;
                Indexcb               := -1;
                CarDelim              := cbDelimChamp.Text;
                for N:=0 to sizeof (CarSepField)-1 do
                begin
                     if CarSepField[N]= FScriptDel.Cbdelim  then begin Indexcb:= N; break; end;
                end;
                Tb := TOB.Create('Enregistrement', nil, -1);

	        for N:=0 to FScriptDel.Champ.Count-1 do
	        begin
                     TF := TOB.Create('fille'+IntToStr(N), Tb, -1);
                     TF.AddChampSupValeur('Colonne', FScriptDel.Champ[N].Name) ;
                     TF.AddChampSupValeur('Longueur', FScriptDel.Champ[N].Longueur);
                     if FScriptDel.Champ[N].Visible then TF.AddChampSupValeur('Visible', 'X')
                     else TF.AddChampSupValeur('Visible', '-');
                     if FScriptDel.Champ[N].Separateur then TF.AddChampSupValeur('Separateur', 'X')
                     else TF.AddChampSupValeur('Separateur', '-');
                     if  FScriptDel.Champ[N].Aligne then TF.AddChampSupValeur('ALin', 'X')
                     else TF.AddChampSupValeur('ALin', '-');
                end;
                Tb.PutGridDetail(GD, False, False, 'Colonne;Longueur;Visible;Separateur;ALin');
                Tb.Free;
                GD.ColTypes[ColonneCoche]:='B';
                GD.ColFormats[ColonneCoche]:=IntToStr(Ord(csCheckbox));
                GD.ColEditables[ColonneCoche]:=FALSE;

                GD.ColTypes[ColonneCoche+1]:='B';
                GD.ColFormats[ColonneCoche+1]:=IntToStr(Ord(csCheckbox));
                GD.ColEditables[ColonneCoche+1]:=FALSE;

                GD.ColTypes[ColonneCoche+2]:='B';
                GD.ColFormats[ColonneCoche+2]:=IntToStr(Ord(csCheckbox));
                GD.ColEditables[ColonneCoche+2]:=FALSE;

                ChargementFichierDeTraitement;
                GD.RowCount              := FScriptDel.Champ.Count+1;
                Champdebut.MaxValue      := FScriptDel.Londeb+10;
                Champdebut.value         := FScriptDel.Londeb;
                ChampFin.MaxValue        := FScriptDel.Lonfin+1;
                ChampFin.value           := FScriptDel.Lonfin;
                Champref.value           := FScriptDel.Longref;
                Champref.MaxValue        := FScriptDel.Longref+1;
                Close;
                if Stream1 <> nil then Stream1.Free;
	end;

     FScriptDel.free;
     Result := TRUE;
end;

procedure TFAssistExt.SuppressionScriptDelim;
//var
//  tbcharger: TTable;
begin
(* AVOIR
  tbcharger := TTable.create(Application);
  tbcharger.DatabaseName := DMImport.DBGlobal.Name;
  tbcharger.Tablename := DMImport.GzImpDelim.TableName;
  tbCharger.open;
  try
    if tbCharger.findKey([Code.Text]) then
    begin
      tbCharger.Delete;
    end;
  except
    PGIBox(Exception(ExceptObject).message, '');
    exit;
  end;
  tbcharger.free;

  tbcharger := TTable.create(Application);
  tbcharger.DatabaseName := DMImport.DBGlobal.Name;
  tbcharger.Tablename := DMImport.GzImpDelim.TableName;
  tbCharger.open;
  try
    if tbCharger.findKey([Code.text]) then
      tbCharger.Delete;
  except
  end;
  tbcharger.free;
*)
end;


procedure TFAssistExt.BVoirClick(Sender: TObject);
begin
  inherited;
{$IFNDEF EAGLCLIENT}
	if not Assigned(VoirDlg) then
		VoirDlg := TVoirDlg.Create(Self);
        VoirDlg.TypeCar := TTypeCar(cbTypeCar.ItemIndex);
	VoirDlg.ASCIIView1.FileName := FileName.Text;
	VoirDlg.Show;
{$ENDIF}
end;

procedure TFAssistExt.cbTypeCarChange(Sender: TObject);
begin
  inherited;
if Fscript <> nil then
    FScript.Options.TypeCar := TTypeCar(cbTypeCar.ItemIndex);
end;

procedure TFAssistExt.cbDelimChampChange(Sender: TObject);
begin
  inherited;
          if cbDelimChamp.ItemIndex < 5 then Car.visible := (cbDelimChamp.ItemIndex >= 5)
end;

procedure  TFAssistExt.LibereDelim;
begin
   SL_Origin.free;
   SL_ODBC.free;
{$IFNDEF EAGLCLIENT}
   if  DBODBC <> nil then  DBODBC.destroy;
{$ENDIF}
   cbAliasODBC.Items.clear;
   ListBox1.clear;
   Memo1.CloseMemo;
end;

procedure TFAssistExt.CodeExit(Sender: TObject);
begin
  inherited;
  {$IFNDEF DBXPRESS}
  If (Code.Text <> '') and (Code.Text[1] in ['0'..'9']) Then
  begin
    PGIBox('Code incorrect, il ne doit pas contenir des champs numériques', 'Conception');
    Code.SetFocus;
  end;
  {$ENDIF}
end;

// Met a jour la colonne Longueur
procedure TFAssistExt.MajColonneLongueur;
var
  compt : integer;
begin
  for compt := 1 to GD.rowcount-1 do
  begin
    GD.cells[1,compt] := EditAppliquerLong.Text;
  end;
end;

procedure TFAssistExt.bMAJLongueurClick(Sender: TObject);
begin
  inherited;
  If (EditAppliquerLong.Text <> '') then
  begin
    if StrToInt(EditAppliquerLong.Text) < 1 then EditAppliquerLong.Text := '1';
    MAJColonneLongueur();
  end;
end;

procedure TFAssistExt.AutoLigneRefClick(Sender: TObject);
var
  F             : textFile;
  CarDelim      : string;
  NbDelimTrouve : integer;
  NbDelimMax    : integer;
  LigneCour     : integer;
  LigneMaxDelim : integer;
  s             : string;
begin
  inherited;
  LigneCour := 1;
  LigneMaxDelim := 1;
  NbDelimMax := 0;
  AssignFile(F, FileName.Text);
  Reset(F);

  // On récupère le délimiteur choisi par l'utilisateur
  if cbDelimChamp.ItemIndex <> -1 then
  begin
    CarDelim := CarSepField[cbDelimChamp.ItemIndex];
    if (cbDelimChamp.ItemIndex = 5) then
         CarDelim := Car.Text[1];
  end
  else CarDelim := cbDelimChamp.Text[1];

  while not EOF(F) do
  begin
    Readln(F, s);
    NbDelimTrouve := CompteDelim(s, CarDelim);
    if (NbDelimTrouve > NbDelimMax) then
    begin
      NbDelimMax := NbDelimTrouve;
      LigneMaxDelim := LigneCour;
    end;
    LigneCour := LigneCour + 1;
  end;
  ChampRef.Value := LigneMaxDelim;
  CloseFile(F);
end;

function TFAssistExt.CompteDelim(sLigne : string; sDelim : string) : integer;
var
  compt : integer;
begin
  Result := 0;
  for compt := 1 to strlen(Pchar(sLigne)) do
  begin
    if sLigne[compt] = sDelim then Result := Result + 1;
  end;
end;

{------------------------- Appel dans compta PGI------------------------}
{$IFDEF CISXPGI}
procedure LanceCorrespCisx (Cmd : string);
var
Filetraitement : string;
OkDelim        : Boolean;
FVarCisx        : TZVarCisx ;    // les variables Cisx
begin
//    FVarCisx        := TZVarCisx.create(cmd);
    TCPContexte.GetCurrent.VarCisx.CHARGECIX(Cmd);

    Filetraitement := ExtractFileName(GetInfoVHCX.Listefichier);
    // Traitement délimité
    OkDelim := TraiteScriptDelim(GetInfoVHCX.Script, GetInfoVHCX.Listefichier, CurrentDonnee+'\'+Filetraitement+'FIXE', FALSE);
    if OkDelim then
       GetInfoVHCX.Listefichier := CurrentDonnee+'\'+Filetraitement+'FIXE';

    AppelCorresp(GetInfoVHCX.Domaine, GetInfoVHCX.Script, GetInfoVHCX.Complement, GetInfoVHCX.Nature, '', taModif, nil, FALSE, 1, FALSE, '', FALSE);
    if OkDelim then
       DeleteFile (CurrentDonnee+'\'+Filetraitement+'FIXE');
//    FreeAndNil (FVarCisx);
end;
{$ENDIF}

constructor TFAssistExt.create(OkR : Boolean; Comment : TActionFiche; Complement : string; py : string);
begin
CommentC    := Comment;
OkRequete   := OkR;
OkDelim     := FALSE;
Indexcb     := -1;
CComplement := Complement;
LePays      := py;
ModifInterface := FALSE;
inherited Create(Application);

end;

end.
