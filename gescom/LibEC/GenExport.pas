unit GenExport;

interface
uses Classes, UTOF, UTOB, ETransUtil, UxmlUtils;

type
    TOF_ExportGen = class(TOF)
    protected
      FTF : TTOBFormat;

    public
      FBaseName : String;
      FExp : TETransfert;

    published
      procedure ExportGenerik(NomTable : String; Keyz : Array of String; Fmt : TTOBFormat; TraitementSpecif : TUseTOB = nil);

    end;

implementation
uses sysutils, Forms, Controls, Dialogs, dbtables, HCtrls, Mul, HMsgBox, HStatus, M3FP;


procedure AGLExportGen(Parms : Array of Variant; Nb : Integer);
var F : TForm;
    TOTOF : TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFmul) then TOTOF := TFMul(F).LaTOF
                  else exit;
  if (TOTOF is TOF_ExportGen) then TOF_ExportGen(TOTOF).ExportGenerik(String(Parms[1]), [Parms[2]], DefaultFormat)
                              else exit;
end;


procedure TOF_ExportGen.ExportGenerik(NomTable : String; Keyz : Array of String; Fmt : TTOBFormat; TraitementSpecif : TUseTOB = nil);
var i,j : integer;
    Motha : TOB;
    FQ : TQuery;
    TokKey : String;
    KeyValuez : Array of Variant;
    b : boolean;
begin
  GetParams;
  FTF := Fmt;

  with TFMul(Ecran) do
  try
    if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
    begin
      PGIBox('Veuillez sélectionner les lignes à exporter', Caption);
      exit;
    end;

    if PGIAsk('Exporter les lignes sélectionnés ?', Caption) = mrNo then Exit;

    FExp := TETransfert.Create;
    FExp.CreateTT(Ecran.Caption, 'Export en cours...', 2, true, true);

    FBaseName := 'Export_'+NomTable+'_Sel';

    FQ := PrepareSQL('SELECT * FROM '+NomTable, true);
    RecupWhereSQL(Q, FQ);
    FQ.Open;

    Q.DisableControls;
    if FListe.AllSelected then
    begin
      FExp.QueryToFormat(FQ, FBaseName, true, Fmt, TraitementSpecif);

      FListe.AllSelected := false;
      bSelectAll.Down := false;
    end else
    begin
      Motha := TOB.Create('Mama_'+NomTable, nil, -1);

      InitMove(FListe.NbSelected,'');
      for i := 0 to FListe.NbSelected-1 do
      begin
        FListe.GotoLeBookMark(i);

        SetLength(KeyValuez, Length(Keyz));
        TokKey := '';
        for j := Low(Keyz) to High(Keyz) do
        begin
          KeyValuez[j] := Q.FindField(Keyz[j]).Value;
          TokKey := TokKey+Keyz[j];
          if j < High(Keyz) then TokKey := TokKey+';';
        end;
        if Length(KeyValuez) = 1 then b := FQ.Locate(TokKey, KeyValuez[0], [])
                                 else b := FQ.Locate(TokKey, VarArrayOf(KeyValuez), []);
        SetLength(KeyValuez, 0);
        if not b then Continue;

        TOB.Create(NomTable, Motha, -1).SelectDB('', FQ);

        MoveCur(False);
      end;
      FiniMove;

      FExp.TOBToFormat(Motha, FBaseName, true, Fmt);
      FExp.SetTTValue(1);
      if Assigned(TraitementSpecif) then TraitementSpecif(Motha);
      Motha.Free;

      FListe.ClearSelected;
    end;
    Q.First;
    Q.EnableControls;

    Ferme(FQ);

    if FExp.PostMQ then PGIInfo('Export effectué avec succès', Caption)
                   else PGIInfo('Une erreur est survenue pendant l''export :'#13+FExp.LastErrorMsg, Caption);
    FExp.FreeTT;
    FExp.Free;

  except
    on E : Exception do PGIInfo('Une erreur est survenue pendant l''export :'#13+E.Message, Caption);

  end;
end;


initialization
RegisterClasses([TOF_ExportGen]);
RegisterAGLProc('ExportGen', true, 2, AGLExportGen);

end.
