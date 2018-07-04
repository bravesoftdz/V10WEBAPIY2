unit ImCRImp;

interface

uses
 Classes,  Controls, Forms,
   Hctrls, Grids, (* Windows, Messages, SysUtils, Graphics, Dialogs,StdCtrls,ExtCtrls, *)
   HPanel, HTB97, HEnt1
  {$IFDEF eAGLCLient}
  , uTOB
  {$ELSE}
  {$IFNDEF DBXPRESS},dbtables{$ELSE},uDbxDataSet{$ENDIF}
  , StdCtrls, ExtCtrls
  {$ENDIF eAGLCLient}
  ;

type
  TFCRImpImmo = class(TForm)
    Dock: TDock97;
    PanelBouton: TToolWindow97;
    BImprimer: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    HPanel1: THPanel;
    ListeCompte: THGrid;
    HPanel2: THPanel;
    HLabel1: THLabel;
    NRECUP: THNumEdit;
    HLabel2: THLabel;
    NPRO: THNumEdit;
    HLabel3: THLabel;
    HLabel4: THLabel;
    NCB: THNumEdit;
    HLabel5: THLabel;
    procedure ListeCompteDblClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;
  TInfoImpImmo = class
    nFiche : integer;
    nPRO : integer;
    nCB : integer;
    TCpte : HTStrings;
    constructor Create;
    destructor Destroy; override;
  end;

procedure CompteRenduImportationImmo ( InfoImp : TInfoImpImmo);

implementation

{$IFDEF SERIE1}
uses AGLInit,
  {$IFDEF eAGLCLient}
     MaineAGL
  {$ELSE}
     FE_Main
  {$ENDIF eAGLCLient}
     ;
{$ELSE}
{$IFNDEF RECUPPCL}
uses  CPGENERAUX_TOM;
{$ENDIF}
{$ENDIF}

{$R *.DFM}

constructor TInfoImpImmo.Create;
begin
  TCpte:=HTStringList.Create;
end;

destructor TInfoImpImmo.Destroy;
begin
  TCpte.Free;
end;

procedure CompteRenduImportationImmo(InfoImp : TInfoImpImmo);
var  FCRImpImmo: TFCRImpImmo;
     i : integer;
BEGIN
  FCRImpImmo:=TFCRImpImmo.Create(Application) ;
  try
    FCRImpImmo.NRECUP.Value := InfoImp.nFiche;
    FCRImpImmo.NPRO.Value   := InfoImp.nPRO;
    FCRImpImmo.NCB.Value    := InfoImp.nCB;
    if InfoImp.TCpte <> nil then
    begin
      if InfoImp.TCpte.Count<>0 then FCRImpImmo.ListeCompte.RowCount:=InfoImp.TCpte.Count+1 ;
      for i:= 0 to InfoImp.TCpte.Count-1 do
      begin
        FCRImpImmo.ListeCompte.Cells[0,i+1]:=InfoImp.TCpte[i];
        FCRImpImmo.ListeCompte.Cells[1,i+1]:='-';
      end;
    end;
    FCRImpImmo.ShowModal ;
  finally
    FCRImpImmo.Free ;
  end;
  Screen.Cursor:=SyncrDefault ;
END ;

procedure TFCRImpImmo.ListeCompteDblClick(Sender: TObject);
var
  Q:TQuery;
  CompteG :string ;
begin
  CompteG:=ListeCompte.Cells[0,ListeCompte.Row];
  if CompteG='' then exit;
  if not Presence('GENERAUX','G_GENERAL',CompteG) then
  begin
    Q:=OpenSQL ('SELECT G_LIBELLE FROM GENERAUX WHERE G_GENERAL = "'+CompteG+'"',true);
    try
      {$IFDEF SERIE1}
      if Q.eof then AGLLanceFiche('C','LISTCPT',CompteG,'',actiontostring(taCreatOne)+';') ;
      {$ELSE}
      {$IFNDEF RECUPPCL}
      if Q.Eof then
         FicheGene (nil,'',CompteG,taCreatOne,0) ;
      {$ENDIF}
      {$ENDIF}
    finally
      ferme(Q) ;
      if Presence('GENERAUX','G_GENERAL',CompteG) then ListeCompte.Cells[1,ListeCompte.Row]:='X'
                                                  else ListeCompte.Cells[1,ListeCompte.Row]:='-';
    end ;
  end
  else
  begin
    ListeCompte.Cells[1,ListeCompte.Row]:='X'
  end ;
end;

end.
