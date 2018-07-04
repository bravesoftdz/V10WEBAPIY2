unit AssistMajCompteurSouche;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
{$IFDEF EAGLCLIENT}
{$ELSE} // EAGLCLIENT
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF} // EAGLCLIENT
  Ed_Tools, HPanel, UTOB, EntGC, HEnt1, UtilUtilitaires;

procedure EntreeMajCompteurSouche;

type
  TFAssistMajCptSouche = class (TFAssist)
    TabSheet1: TTabSheet;
    PTITRE: THPanel;
    TabSheet2: TTabSheet;
    LExplication: THLabel;
    GBParam: TGroupBox;
    TNaturePieceG: THLabel;
    CBNaturePieceG: THValComboBox;
    PanelFin: TPanel;
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    TRecap: THLabel;
    LBRecap: TListBox;
    bImprimer: TToolbarButton97;
    procedure bFinClick (Sender: TObject);
    procedure bSuivantClick (Sender: TObject);
    procedure FormClose (Sender: TObject; var Action: TCloseAction);
    procedure FormCreate (Sender: TObject);
    procedure FormShow (Sender: TObject);
    procedure CBNaturePieceGChange(Sender: TObject);
    procedure bImprimerClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
    TobSouche : TOB;
    stDejaSouche : string;
    ioerr : TIOErr;
    bFinCaptionOrigine : string;
    procedure ChargeNumeroMax;
    procedure ChargeSouche (TobS : TOB;
                            stPrefixe, stChamp, stNaturePieceg : string;
                            bNatureSeule, bNaturePlus : boolean);
    function ConstruitSqlSouche (TobS : TOB; stSqlOrig : string) : string;
    procedure ListeRecap;
    procedure MiseAJourCompteur;
    procedure RenseigneJournalEvenement;
    procedure SoucheAModifier (stNaturePieceG : string);
  end;

var
  FAssistMajCptSouche: TFAssistMajCptSouche;

implementation

{$R *.DFM}

procedure EntreeMajCompteurSouche;
var Fo_MajCptSouche : TFAssistMajCptSouche;
begin
     Fo_MajCptSouche := TFAssistMajCptSouche.Create (Application);
     Fo_MajCptSouche.CBNaturePieceg.Value := '';
     Try
         Fo_MajCptSouche.ShowModal;
     Finally
         Fo_MajCptSouche.free;
     End;
end;

procedure TFAssistMajCptSouche.bFinClick (Sender: TObject);
begin
    inherited;
    if bFin.Caption <> 'Terminer' then
    begin
        ioErr := Transactions (MiseAJourCompteur, 1);
        if ioErr <> oeOk then
            PgiBox ('La Mise a jour ne s''est pas effectuée');
        bFinCaptionOrigine := bFin.Caption;
        bFin.Caption := 'Terminer';
        bImprimer.Enabled := true;
    end else Close;
end;

procedure TFAssistMajCptSouche.bSuivantClick (Sender: TObject);
var iInd : integer;
    TobPP, TobS, TobSPrime : TOB;
    TSql : TQuery;
    stSouche : string;
begin
    inherited;
    if bFin.Enabled then
    begin
        TobSouche.ClearDetail;
        stDejaSouche := '';
        if (CBNaturePieceG.Value = '') then
        begin
            for iInd := 0 to VH_GC.TOBParPiece.Detail.Count - 1 do
            begin
                TobPP := VH_GC.TobParPiece.Detail[iInd];
                SoucheAModifier (TobPP.GetValue ('GPP_NATUREPIECEG'));;
            end;
        end else SoucheAModifier (CBNaturePieceG.Value);

        stSouche := '';
        for iInd := 0 to TobSouche.Detail.Count - 1 do
        begin
            if stSouche <> '' then stSouche := stSouche + ', ';
            stSouche := stSouche + '"' + TobSouche.Detail[iInd].GetValue ('SOUCHE') + '"';
        end;
        TobS := Tob.Create ('', nil, -1);
        TSql := OpenSql ('SELECT SH_SOUCHE, SH_NUMDEPART FROM SOUCHE WHERE SH_SOUCHE in (' + stSouche +
                            ') AND SH_TYPE="GES"', True,-1,'',true);
        if not TSql.Eof then
            TobS.LoadDetailDB ('', '', '', TSql, False);
        Ferme (TSql);
        for iInd := 0 to TobS.Detail.Count - 1 do
        begin
            TobSPrime := TobSouche.FindFirst (['SOUCHE'], [TobS.Detail[iInd].GetValue ('SH_SOUCHE')], False);
            if TobSPrime <> nil then
                TobSPrime.AddChampSupValeur ('ANCIENNUMERO', TobS.Detail[iInd].GetValue ('SH_NUMDEPART'));
        end;
        TobS.Free;

        ChargeNumeroMax;        

        ListeRecap;
    end;
end;

procedure TFAssistMajCptSouche.bImprimerClick(Sender: TObject);
var TobToPrint : TOB;
    Cpt : integer;
begin
  inherited;
  TobToPrint := TOB.Create('',nil,-1);
  for Cpt := 0 to LBRecap.Items.Count -1 do
    UtilTobCreat(TobToPrint,'','',LBRecap.Items[Cpt],'');
  UtilTobPrint(TobToPrint,Caption,1);
  TobToPrint.free;
end;

procedure TFAssistMajCptSouche.FormClose (Sender: TObject;
  var Action: TCloseAction);
begin
    inherited;
    RenseigneJournalEvenement;
    TobSouche.Free;
end;

procedure TFAssistMajCptSouche.FormCreate (Sender: TObject);
begin
  inherited;
TobSouche := Tob.Create ('', nil, -1);
end;

procedure TFAssistMajCptSouche.FormShow (Sender: TObject);
begin
    inherited;
{    if Not BlocageMonoPoste (True) then
    begin
        Close;
    end; }
    bImprimer.Enabled := false;
    ioerr := oeOk;
end;

procedure TFAssistMajCptSouche.ListeRecap;
var iInd, iInd2, iInd3 : integer;
    TobS, TobS1, TobS2 : TOB;
begin
    LBRecap.Items.Clear;
    LBRecap.Items.Add (PTITRE.Caption);
    LBRecap.Items.Add ('');
    if CBNaturePieceg.Value = '' then
    begin
        LBRecap.Items.Add ('Mise à Jour des compteurs');
        LBRecap.Items.Add ('pour toutes les natures de pièce');
    end else
    begin
        LBRecap.Items.Add ('Mise à Jour des compteurs');
        LBRecap.Items.Add ('  Nature de pièce : ' + CBNaturePieceG.Text);
    end;

    for iInd := 0 to TobSouche.Detail.Count - 1 do
    begin
        LBRecap.Items.Add ('');
        TobS := TobSouche.Detail[iInd];
        LBRecap.Items.Add ('Souche : ' + RechDom ('GCSOUCHEGESCOM', TobS.GetValue ('SOUCHE'), False));
        if TobS.FieldExists ('ANCIENNUMERO') then
            LBRecap.Items.Add ('  Ancien Numéro : ' + IntToStr (TobS.GetValue ('ANCIENNUMERO')));
        LBRecap.Items.Add ('  Numéro mis à jour : ' + IntToStr (TobS.GetValue ('NUMERO')));
        for iInd2 := 0 to TobS.Detail.Count - 1 do
        begin
            LBRecap.Items.Add ('');
            TobS1 := TobS.Detail[iInd2];
            if TobS1.FieldExists ('SEULE') then
                LBRecap.Items.Add ('    Nature de pièce : ' +
                                    GetInfoParPiece (TobS1.GetValue ('NATURE'), 'GPP_LIBELLE') + ('  (Seule)'))
            else LBRecap.Items.Add ('    Nature de pièce : ' +
                                        GetInfoParPiece (TobS1.GetValue ('NATURE'), 'GPP_LIBELLE'));
            for iInd3 := 0 to TobS1.Detail.Count - 1 do
            begin
                TobS2 := TobS1.Detail[iInd3];
                if TobS2.FieldExists ('DOMAINE') then
                begin
                    LBRecap.Items.Add ('        Domaine : ' + RechDom ('YYDOMAINE',
                                       TobS2.GetValue ('DOMAINE'), False));
                end;
                if TobS2.FieldExists ('ETABLISSEMENT') then
                begin
                    LBRecap.Items.Add ('        Etablissement : ' + RechDom ('TTETABLISSEMENT',
                                       TobS2.GetValue ('ETABLISSEMENT'), False));
                end;
{$IFDEF NOMADE}
                if TobS2.FieldExists ('CODESITE') then
                begin
                    LBRecap.Items.Add ('        Site : ' + TobS2.GetValue ('CODESITE'));
                end;
{$ENDIF} // NOMADE
            end;
        end;
    end;
    LBRecap.Items.Add ('');
    LBRecap.Items.Add ('Nombre de souches à modifier : ' + IntToStr (TobSouche.Detail.count));
end;

procedure TFAssistMajCptSouche.ChargeNumeroMax;
var iInd : integer;
    TobS, TobP : TOB;
    TSql : TQuery;
    stSql : string;
begin
    TobP := Tob.Create ('', nil, -1);
    stSql := '';
    for iInd := 0 to TobSouche.Detail.Count - 1 do
    begin
        TobS := TobSouche.Detail[iInd];
        stSql := ConstruitSqlSouche (TobS, stSql);
    end;
    TSql := OpenSql ('SELECT GP_SOUCHE, MAX (GP_NUMERO) AS MAXNUM FROM PIECE WHERE ' + stSql + ' GROUP BY GP_SOUCHE',
                     True,-1,'',true);
    TobP.LoadDetailDB ('', '', '', TSql, False);
    Ferme (TSql);
    for iInd := 0 to TobP.Detail.Count - 1 do
    begin
        TobS := TobSouche.FindFirst (['SOUCHE'], [TobP.Detail[iInd].GetValue ('GP_SOUCHE')], False);
        if TobS <> nil then
        begin
            TobS.PutValue ('NUMERO', TobP.Detail[iInd].GetValue ('MAXNUM') + 1);
        end;
    end;
    TobP.Free;
end;

procedure TFAssistMajCptSouche.ChargeSouche (TobS : TOB;
                                             stPrefixe, stChamp, stNaturePieceg : string;
                                             bNatureSeule, bNaturePlus : boolean);
var iInd : integer;
    TobS1, TobS2, TobDetailLu : TOB;
    TobLue : TOB;
    stNature : string;
    bOkDetail : boolean;
{$IFDEF NOMADE}
    TSql : TQuery;
{$ENDIF} // NOMADE
begin
    if stPrefixe = 'GPC_' then TobLue := VH_GC.MTOBParPieceComp
    else if stPrefixe = 'GDP_' then TobLue := VH_GC.MTOBParPieceDomaine
{$IFDEF NOMADE}
    else if stPrefixe = 'GSP_' then
    begin
        TobLue := Tob.Create ('', nil, -1);
        TSql := OpenSQL('SELECT GSP_NATUREPIECEG, GSP_SOUCHE FROM SITEPIECE WHERE GSP_NATUREPIECEG="' +
                            stNaturePieceG + '" and GSP_CODESITE="' +
                            PCP_LesSites.LeSiteLocal.SSI_CODESITE + '"', True,-1,'',true);
        TobLue.LoadDetailDB ('', '', '', TSql, False);
        Ferme (TSql);
    end
{$ENDIF} // NOMADE
    else TobLue := VH_GC.TOBParPiece;

    for iInd := 0 to TobLue.Detail.Count - 1 do
    begin
        TobDetailLu := TobLue.Detail[iInd];
        bOkDetail := True;
        if (Pos (TobDetailLu.GetValue (stPrefixe + 'SOUCHE'), stDejaSouche) > 0) then bOkDetail := False
        else if bNaturePlus then
        begin
            if (TobDetailLu.GetValue (stPrefixe + 'SOUCHE') <> TobS.GetValue ('SOUCHE')) or
               (TobDetailLu.GetValue (stPrefixe + 'NATUREPIECEG') = stNaturePieceG) then bOkDetail := False;
        end else if TobDetailLu.GetValue (stPrefixe + 'NATUREPIECEG') <> stNaturePieceG then bOkDetail := False;

        if bOkDetail then
        begin
            if bNaturePlus then stNature := TobDetailLu.GetValue (stPrefixe + 'NATUREPIECEG')
            else stNature := stNaturePieceG;
            if (TobS = nil) or (iInd > 0) then
            begin
                TobS := TobSouche.FindFirst (['SOUCHE'], [TobDetailLu.GetValue (stPrefixe + 'SOUCHE')], False);
                if TobS = nil then
                begin
                    TobS := Tob.Create ('', TobSouche, -1);
                    TobS.AddChampSupValeur ('SOUCHE', TobDetailLu.GetValue (stPrefixe + 'SOUCHE'));
                    TobS.AddChampSupValeur ('NUMERO', 0);
                    TobS.AddChampSup ('NOUVELLE', False);
                end;
            end;
            TobS1 := TobS.FindFirst (['NATURE'], [stNature], False);
            if TobS1 = nil then
            begin
                TobS1 := Tob.Create ('', TobS, -1);
                TobS1.AddChampSupValeur ('NATURE', stNature);
            end;
            if bNatureSeule then TobS1.AddChampSup ('SEULE', False)
            else
            begin
                TobS2 := TobS1.FindFirst ([stChamp], [TobDetailLu.GetValue (stPrefixe + stChamp)], False);
                if TobS2 = nil then
                begin
                    TobS2 := Tob.Create ('', TobS1, -1);
                    TobS2.AddChampSupValeur (stChamp, TobDetailLu.GetValue (stPrefixe + stChamp));
                end;
            end;
        end;
    end;
{$IFDEF NOMADE}
    if stPrefixe = 'GSP_' then
    begin
        TobLue.Free;
    end;
{$ENDIF} // NOMADE
end;

function TFAssistMajCptSouche.ConstruitSqlSouche (TobS : TOB; stSqlOrig : string) : string;
var iInd : integer;
    stSql : string;
begin
    if stSqlOrig <> '' then stSql := stSqlOrig + ' OR '
    else stSql := '';
    stSql := stSql + ' (GP_SOUCHE="' + TobS.GetValue ('SOUCHE') + '"';
    if TobS.Detail.Count > 0 Then
    begin
        stSql := stSql + ' AND GP_NATUREPIECEG IN (';
        for iInd := 0 to TobS.Detail.Count - 1 do
        begin
            if iInd > 0 then stSql := stSql + ', ';
            stSql := stSql + '"' + TobS.Detail[iInd].GetValue ('NATURE') + '"';
        end;
        stSql := stSql + ') ';
    end;
    stSql := stSql + ')';
    Result := stSql;
end;

procedure TFAssistMajCptSouche.MiseAJourCompteur;
var iInd : integer;
    TobS : TOB;
begin
    InitMoveProgressForm (nil, 'Traitement en cours...', 'Mise à jour des compteurs',
                          TobSouche.Detail.Count, true, true) ;
    try
        for iInd := 0 to TobSouche.Detail.Count - 1 do
        begin
            TobS := TobSouche.Detail[iInd];
            ExecuteSQL ('UPDATE SOUCHE SET SH_NUMDEPART=' + IntToStr (TobS.GetValue ('NUMERO')) +
                        ' WHERE SH_SOUCHE="' + TobS.GetValue ('SOUCHE') + '" AND SH_TYPE="GES"');
            MoveCurProgressForm ('');
        end;
    except
        V_PGI.ioError := oeUnknown ;
    end;
    FiniMoveProgressForm ;
end;

procedure TFAssistMajCptSouche.RenseigneJournalEvenement;
var TOBJnal : TOB ;
    NumEvt : integer ;
    TSql : TQuery ;
begin
    NumEvt := 0 ;
    TOBJnal := TOB.Create ('JNALEVENT', Nil, -1) ;
    TOBJnal.PutValue ('GEV_TYPEEVENT', 'UTI');
    TOBJnal.PutValue ('GEV_LIBELLE', PTITRE.Caption);
    TOBJnal.PutValue ('GEV_DATEEVENT', Date);
    TOBJnal.PutValue ('GEV_UTILISATEUR', V_PGI.User);
    TSql := OpenSQL ('SELECT MAX (GEV_NUMEVENT) FROM JNALEVENT',True,-1,'',true) ;
    if Not TSql.EOF then NumEvt := TSql.Fields[0].AsInteger ;
    Ferme (TSql) ;
    Inc (NumEvt) ;
    TOBJnal.PutValue ('GEV_NUMEVENT', NumEvt);
    LBRecap.Items.Add ('');
    if ioerr = oeOk then
    begin
        if bFin.Caption = 'Terminer' then
        begin
            TOBJnal.PutValue ('GEV_ETATEVENT', 'OK');
            LBRecap.Items.Add ('Traitement terminé');
        end else
        begin
            TOBJnal.PutValue ('GEV_ETATEVENT', 'ANN');
            LBRecap.Items.Add ('Traitement Annulé');
        end;
    end else
    begin
        TOBJnal.PutValue ('GEV_ETATEVENT', 'ERR');
        LBRecap.Items.Add ('Traitement non terminé');
    end;
    TOBJnal.PutValue ('GEV_BLOCNOTE', LBRecap.Items.Text);
    TOBJnal.InsertDB (Nil) ;
    TOBJnal.Free ;
end;

procedure TFAssistMajCptSouche.SoucheAModifier (stNaturePieceg : string);
var iInd, iInd2, iNbNature : integer;
    TobS, TobS1 : TOB;
begin
{$IFDEF NOMADE}
    ChargeSouche (nil, 'GSP_', 'CODESITE', stNaturePieceG, False, False);
{$ENDIF} // NOMADE
    ChargeSouche (nil, 'GPC_', 'ETABLISSEMENT', stNaturePieceG, False, False);
    ChargeSouche (nil, 'GDP_', 'DOMAINE', stNaturePieceg, False, False);
    ChargeSouche (nil, 'GPP_', 'NATURE', stNaturePieceG, True, False);

    for iInd := 0 to TobSouche.Detail.Count - 1 do
    begin
        TobS := TobSouche.Detail[iInd];

        iNbNature := TobS.Detail.Count - 1;
        for iInd2 := 0 to iNbNature do
        begin
            if TobS.FieldExists ('NOUVELLE') then
            begin
                TobS1 := TobS.Detail[iInd2];
{$IFDEF NOMADE}
                ChargeSouche (TobS, 'GSP_', 'CODESITE', TobS1.GetValue ('NATURE'), False, True);
{$ENDIF} // NOMADE
                ChargeSouche (TobS, 'GPC_', 'ETABLISSEMENT', TobS1.GetValue ('NATURE'), False, True);
                ChargeSouche (TobS, 'GDP_', 'DOMAINE', TobS1.GetValue ('NATURE'), False, True);
                ChargeSouche (TobS, 'GPP_', 'NATURE', TobS1.GetValue ('NATURE'), True, True);
            end;
        end;
        TobS.DelChampSup ('NOUVELLE', False);
        if stDejaSouche <> '' then stDejaSouche := stDejaSouche + ', ';
        stDejaSouche := stDejaSouche + '"' + TobS.GetValue ('SOUCHE') + '"';
    end;
end;

procedure TFAssistMajCptSouche.CBNaturePieceGChange(Sender: TObject);
begin
    inherited;
    if bFin.Caption = 'Terminer' then
    begin
        RenseigneJournalEvenement;
        LBRecap.Clear;
        bFin.Caption := bFinCaptionOrigine;
    end;
end;

end.
