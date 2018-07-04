unit RTViewPersp;

interface

uses

  SysUtils, Classes, Controls, Forms,
  UTobView, Utob, HCtrls, 
{$IFDEF EAGLCLIENT}
     MainEAGL,
{$ELSE}
     Fe_Main, 
{$ENDIF}
     TVProp, HTB97, ExtCtrls, HPanel;

type
  TTobViewPersp = class(TForm)
    TobViewer1: TTobViewer;
    PBoutons: THPanel;
    BProperty: TToolbarButton97;
    BPrint: TToolbarButton97;
    procedure TobViewer1DblClickCell(Sender: TTobViewer; ColName: String;
      ACol, ARow: Integer; DataType: TDataType);
    procedure BPropertyClick(Sender: TObject);
{$IFNDEF EAGLCLIENT}
    procedure BPrintClick(Sender: TObject);
{$ENDIF}
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;


procedure TobView( TobGen: TOB; var Params: HTStrings ); overload;
procedure TobView( SqlTxt: String; var Params: HTStrings ); overload;

implementation

{$R *.DFM}

procedure TobView( TobGen: TOB; var Params: HTStrings );
begin
	with TTobViewPersp.Create( Application ) do
		try
			TobViewer1.TobView( TobGen, Params );
			if Params <> nil then TobViewer1.SetGridView( Params, True );
			ShowModal;
			if Params <> nil then TobViewer1.GetGridView( Params );
		finally
			Free;
		end;
end;

procedure TobView( SqlTxt: String; var Params: HTStrings );
begin
	with TTobViewPersp.Create( Application ) do
		try
			TobViewer1.TobView( SqlTxt, Params );
			if Params <> nil then TobViewer1.SetGridView( Params, True );
			ShowModal;
			if Params <> nil then TobViewer1.GetGridView( Params );
		finally
			Free;
		end;
end;

procedure TTobViewPersp.TobViewer1DblClickCell(Sender: TTobViewer;
  ColName: String; ACol, ARow: Integer; DataType: TDataType);
 var s_param :string;
begin
//   ShowMessage( 'Acol :'+IntToStr(ACol) +ColName+': '+Sender.AsString[ACol, ARow] );
    if ColName = 'RPE_AUXILIAIRE' then
    begin
      AGLLanceFiche('GC','GCTIERS','',Sender.AsString[ACol, ARow],'T_NATUREAUXI='+Sender.AsString[ACol-1, ARow]);
    end else
    if ColName = 'T_LIBELLE' then
    begin
      AGLLanceFiche('GC','GCTIERS','',Sender.AsString[ACol-1, ARow],'T_NATUREAUXI='+Sender.AsString[ACol-2, ARow]);
    end else
    if ColName = 'GCL_LIBELLE' then
    begin
      AGLLanceFiche('GC','GCCOMMERCIAL','',Sender.AsString[ACol-2, ARow]+';'+Sender.AsString[ACol-1, ARow],'ACTION=MODIFICATION');
    end else
    if ColName = 'RPE_INTERVENANT' then
    begin
      AGLLanceFiche('GC','GCCOMMERCIAL','',Sender.AsString[ACol-1, ARow]+';'+Sender.AsString[ACol, ARow],'ACTION=MODIFICATION');
    end else
    begin
     s_param := ''+ Sender.AsString[2, ARow] + ';' + IntToStr(Sender.AsInteger[4, ARow]) + ';' + IntToStr(Sender.AsInteger[5, ARow]) +'';
     AGLLanceFiche('RT','RTPERSPECTIVES','',s_param ,'');
     //     AGLLanceFiche('RT','RTPERSPECTIVES','',Sender.AsString[2, ARow]+';'+IntToStr(Sender.AsInteger[4, ARow])+';'+ IntToStr(Sender.AsInteger[5, ARow]),'');
    //ouvrefiche('RT','RTPERSPECTIVES','',GetChamp('RPE_AUXILIAIRE')+';'+GetChamp('RPE_PERSPECTIVE')+';'+GetChamp('RPE_INDICE'),'') ;
    end;
end;

procedure TTobViewPersp.BPropertyClick(Sender: TObject);
begin
    TVProp.DlgProp( TobViewer1 );
end;
{$IFNDEF EAGLCLIENT}
procedure TTobViewPersp.BPrintClick(Sender: TObject);
begin
    TobViewer1.Print( True, True, 'Impression');
end;
{$ENDIF}
end.
