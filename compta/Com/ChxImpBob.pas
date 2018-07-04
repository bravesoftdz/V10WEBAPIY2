unit ChxImpBob;

interface

uses
  Windows, Classes, Controls, Forms,
  UIUtil, Hent1, Vierge, StdCtrls, Mask, Hctrls, HPanel, HSysMenu, uYFILESTD,
  HTB97, Sysutils, dialogs;

type
  TFChxDoss = class(TFVierge)
    Label1: TLabel;
    FListefichiers: THCritMaskEdit;
    CODEPRODUIT: THCritMaskEdit;
    HLabel1: THLabel;
    HLabel2: THLabel;
    EXTENTION: THCritMaskEdit;
    HLabel3: THLabel;
    CRIT1: THCritMaskEdit;
    TCRIT2: THLabel;
    CRIT2: THCritMaskEdit;
    HLabel4: THLabel;
    CRIT3: THCritMaskEdit;
    HLabel5: THLabel;
    CRIT4: THCritMaskEdit;
    HLabel6: THLabel;
    CRIT5: THCritMaskEdit;
    procedure BValiderClick(Sender: TObject);
    procedure FListefichiersElipsisClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Déclarations privées }
    StArg : string;
  public
    { Déclarations publiques }
    constructor create (stArgument : string); reintroduce; overload ; virtual ;

  end;

procedure ImportFichierBob(StArgument : string) ;

implementation



{$R *.DFM}


procedure ImportFichierBob(StArgument : string) ;
var
    X : TFChxDoss ;
begin
  X := TFChxDoss.Create(StArgument);
    try
      X.ShowModal;
    finally
      X.Free;
    end;
end;


procedure TFChxDoss.BValiderClick(Sender: TObject);
begin
  inherited;
        AGL_YFILESTD_IMPORT(FListefichiers.Text,Codeproduit.Text,ExtractFileName(FListefichiers.Text), Extention.Text,
                Crit1.Text, Crit2.Text, Crit3.Text,
                Crit4.Text, Crit5.Text,'-','-','-','-','-',
                V_PGI.LanguePrinc,
                '',
                '');
       ModalResult := MrOk;
end;

procedure TFChxDoss.FListefichiersElipsisClick(Sender: TObject);
var
  I   : integer;
  NN  : string;
  CurrentDir : string;
begin
  inherited;
    FListefichiers.Text := '';
    with Topendialog.create(Self) do
    begin
         Title := ' Choisir la liste des fichiers à traiter';
         Options := [ofFileMustExist, ofHideReadOnly, ofAllowMultiSelect ];
         Filter := 'Fichiers texte (*.PDF)|*.txt|Tous les fichiers (*.*)|*.*';
         FilterIndex := 2;
         if Execute then
         begin
         with Files do
              for I := 0 to Count - 1 do
              begin
              if I = 0 then FListefichiers.Text := Strings[I]
              else FListefichiers.Text := FListefichiers.Text+';'+Strings[I];
              end;
         end;
         NN := FileName;
         CurrentDir := ExtractFileDir (NN);
         SetCurrentDirectory(PChar(CurrentDir));
         Free;
    end;

end;

constructor TFChxDoss.create(stArgument : string);
begin
StArg := StArgument;
inherited Create(Application);
end;


procedure TFChxDoss.FormShow(Sender: TObject);
Var
S : string;
begin
  inherited;
  S := StArg;
  CODEPRODUIT.Text := ReadTokenSt(S);
  FListefichiers.Text := ReadTokenSt(S);
  CRIT1.Text := ReadTokenSt(S);
  Extention.Text := ReadTokenSt(S);

end;

procedure TFChxDoss.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
ModalResult := MrOk;
end;

end.
