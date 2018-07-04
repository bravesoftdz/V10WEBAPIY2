{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 14/05/2014
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTIMPORTBSV ()
Mots clefs ... : TOF;BTIMPORTBSV
*****************************************************************}
Unit BTIMPORTBSV_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
		 HPanel,
		 Hctrls,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     Variants,
     XMLDoc,xmlintf,UXMLvalidation,CBPPath,Graphics, UImportFacXml,
     UTOF,
     AglInit ;

Type
  TOF_BTIMPORTBSV = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    fImpFacXml : TImportFachXml;
    fOkOk : boolean;
    BVALIDER : TToolbarButton97;
    NOMXML : THEdit;
    BVOIRCONTENU : TToolbarButton97;
    PINTERNE : THPanel;
    CODEFOU : THEdit;
    MONTANTHT,TVA,MONTANTTTC : THnumEdit;
    NOMFOU,INFOPIECE : THLabel;
    MDETAIL : Tmemo;
    TOBretour : TOB;
    ColorInit : TColor;
    //
    procedure GetControls;
    procedure InitEcran;
    procedure GetFichierXml (Sender : TObject);
    procedure BVOIRCONTENUClick (Sender : TObject);
    function ConstitueInfoPiece (TOBDatas : TOB): boolean;
    function ConstitueDetail (TOBDatas : TOB): boolean;
    function PieceExists(Naturepiece, Numero: string): boolean;
    procedure AfficheEcran;
    procedure AjouteEvent;
    //
  end ;

Implementation
uses Dialogs,BTPUtil,AffaireUtil,factGrp,ParamSoc;

procedure TOF_BTIMPORTBSV.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTBSV.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTBSV.OnUpdate ;
var LibPiece : string;
//		II : Integer;
    TOBGenere{,TOBP} : TOB;
//  Sql : string;
    DateFac : TDateTime;
//    QQ : TQuery;
    OkOk : Boolean;
//    SMontantHt,SMontantTTc : Double;
//    Mess : string;
    i_ret : Integer;
    TOBDatas : TOB;
begin
  Inherited ;
  TOBretour := nil;
  OkOk := True;
  TOBDatas := fImpFacXml.TOBDatas;
  //
  libPiece := TOBDatas.GetString('LIBDOC');
  // ---
  if PGIAsk('Vous allez lancer le traitement de génération de la '+LibPiece+'. Confirmez-vous ?')<>Mryes then Exit;
  TOBGenere := TOB.Create('LES PIECES',nil,-1);
  DateFac := TOBDatas.GetDateTime('DATEFAC');
  //
  (*
  SMontantHt := 0;
  SMontantTTC := 0;
  for II := 0 to TOBDatas.detail.Count -1 do
  begin
    Sql := 'SELECT * FROM PIECE WHERE GP_NATUREPIECEG="'+TOBDatas.detail[II].GetString('NATUREPIECE')+'" AND '+
    			 'GP_REFEXTERNE='+TOBDatas.detail[II].GetString('NUMEROPIECE');
    QQ := OpenSql (Sql,True,1,'',true);
    if Not QQ.eof then
    begin
	    TOBP := TOB.Create ('PIECE',TOBGenere,-1);
    	TOBP.SelectDB ('',QQ);
      SMontantHt := SMontantHt + TOBP.GetDouble('GP_TOTALHTDEV');
      SMontantTTC := SMontantTTC + TOBP.GetDouble('GP_TOTALTTCDEV');
    end;
    ferme (QQ);
  end;
  if (SMontantHt <> TOBDatas.GetDouble('MONTANTHT')) or (SMontantTTC <> TOBDatas.GetDouble('MONTANTTTC')) then
  begin
    mess := 'Des écarts ont été constatés entre le montant de la facture fournisseur et les montants des réceptions#13#10'+
    				'Désirez-vous quand même enregister la facture ?';
    if PGIAsk(Mess)<> Mryes then OkOk := false;
  end;
  *)
  //
  if OkOK then
  begin
    i_ret := RegroupeLesPieces (TOBGenere,TOBDatas.GetString('TYPEDOC'),False,false,True,0,
                        				DateFac,true,true,false,False,false,'',false,'','GENEREFF',TOBDatas.GetDouble('MONTANTHT'),
                        				TOBDatas.GetDouble('MONTANTTTC'),TOBDatas.GetString('NUMFACTUREFOU'));
		if TheTOB <> nil Then TOBretour := TheTOB;
    if i_ret = 0 then
    begin
      fImpFacXml.MoveToTreated (NOMXML.Text);
    end else
    begin
      fImpFacXml.MoveToError (NOMXML.Text);
    end;
    AjouteEvent;
    if TheTOB <> nil then
    begin
      FreeAndNil(TheTOB);
      TOBretour := nil;
    end;
  end;
  //
	InitEcran;
end ;



procedure TOF_BTIMPORTBSV.AjouteEvent;
var QQ: TQuery;
  MotifPiece: TStrings;
  NumEvent: integer;
  II : Integer;
begin

  MotifPiece := TStringList.Create;
  NumEvent := 0;
  QQ := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', True,-1, '', True);
  if not QQ.EOF then NumEvent := QQ.Fields[0].AsInteger;
  Ferme(QQ);
  Inc(NumEvent);
  QQ := OpenSQL('SELECT * FROM JNALEVENT WHERE GEV_TYPEEVENT="GEN" AND GEV_NUMEVENT=-1', False);
  QQ.Insert;
  InitNew(QQ);
  QQ.FindField('GEV_NUMEVENT').AsInteger := NumEvent;
  QQ.FindField('GEV_TYPEEVENT').AsString := 'BIX';
  QQ.FindField('GEV_LIBELLE').AsString := 'Intégration du fichier XML';
  MotifPiece.Add('Document : '+NOMXML.text);
  QQ.FindField('GEV_DATEEVENT').AsDateTime := Date;
  QQ.FindField('GEV_UTILISATEUR').AsString := V_PGI.User;
  if fImpFacXml.Statutimp = [TCsOk] then
  begin
    QQ.FindField('GEV_ETATEVENT').AsString := 'OK';
    if TOBRetour <> nil then
    begin
      for II := 0 to TOBretour.detail.count -1 do
      begin
        MotifPiece.Add(rechdom('GCNATUREPIECEG',TOBRetour.detail[ii].GetString('NATUREPIECE'),false)+
                       ' N° '+InttOstr(TOBRetour.detail[ii].GetInteger('NUMERO'))+' Généré');
      end;
    end;
  end else
  begin
    QQ.FindField('GEV_ETATEVENT').AsString := 'ERR';
    fImpFacXml.ContitueMotif (MotifPiece);
  end;
  TMemoField(QQ.FindField('GEV_BLOCNOTE')).Assign(MotifPiece);
  QQ.Post;
  Ferme(QQ);
  MotifPiece.Free;
end;

procedure TOF_BTIMPORTBSV.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTBSV.OnArgument (S : String ) ;
begin
  Inherited ;
  ColorInit := THPanel(GetControl('PCONSTITUANT')).Font.Color;
  fImpFacXml := TImportFachXml.create;
  GetControls;
  InitEcran;
  NOMXML.OnElipsisClick := GetFichierXml;
  BVOIRCONTENU.OnClick := BVOIRCONTENUClick;
  fOkOk := false;
end ;

procedure TOF_BTIMPORTBSV.OnClose ;
begin
	fImpFacXml.Free;
  Inherited ;
end ;

procedure TOF_BTIMPORTBSV.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTBSV.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTBSV.InitEcran;
begin
  BVALIDER.Enabled := false;
  PINTERNE.Visible := false;
  NOMXML.Text := '';
  BVOIRCONTENU.Enabled := false;
  CODEFOU.Text := '';
  NOMFOU.Caption := '';
  INFOPIECE.Caption := '';
  MONTANTHT.Value := 0;
  TVA.Value := 0;
  MONTANTTTC.Value := 0;
  MDETAIL.Lines.Clear;
end;

procedure TOF_BTIMPORTBSV.GeTControls;
begin
  BVALIDER := TToolbarButton97 (GetControl('BVALIDER'));
  NOMXML := THEdit (GetControl('NOMXML'));
  BVOIRCONTENU := TToolbarButton97 (GetControl('BVOIRCONTENU'));
  PINTERNE := THPanel(GetControl('PINTERNE'));
  CODEFOU := THEdit(GetControl('CODEFOU'));
  MONTANTHT := THNumEdit(GetControl('MONTANTHT'));
  TVA := THNumEdit(GetControl('TVA'));
  MONTANTTTC := THNumEdit(GetControl('MONTANTTTC'));
  NOMFOU := THLabel(GeTControl('NOMFOU'));
  INFOPIECE := THLabel(GeTControl('INFOPIECE'));
  MDETAIL := Tmemo (GetControl('MDETAIL'));
end;

procedure TOF_BTIMPORTBSV.GetFichierXml(Sender : TObject);
var BrowseXML : TOpenDialog;
begin

  BrowseXML := TOpenDialog.Create(ecran);
  BrowseXML.Title := 'Recherche fichier XML à intégrer';
  BrowseXML.Filter := 'Fichiers XML|*.XML';

  BrowseXML.InitialDir := fImpFacXml.RepertDef;

  if BrowseXML.Execute then
  begin
    NOMXML.Text := IncludeTrailingBackslash(BrowseXML.FileName);
    BVOIRCONTENU.Enabled := True;
  end else
  begin
    NOMXML.Text := '';
  end;

  FreeAndNil(BrowseXML);
  if nomXml.Text = '' then BVOIRCONTENU.Enabled := false;
  PINTERNE.Visible := false;

end;

procedure TOF_BTIMPORTBSV.BVOIRCONTENUClick(Sender: TObject);
begin
	// TODO lecture du xml et alimentation des datas d'écran
  MDETAIL.Lines.Clear;
  fOkOk := false;

  CODEFOU.Text := '';
  NOMFOU.Caption := '';
  INFOPIECE.Caption := '';
  MONTANTHT.Value := 0;
  TVA.Value := 0;
  MONTANTTTC.Value := 0;
  //
  fImpFacXml.TraiteXml(NOMXML.Text);
  //
  if TCsNconforme in fImpFacXml.Statutimp then
  begin
  	fImpFacXml.MoveToError (NOMXML.Text);
  	AjouteEvent;
  end else
  begin
    AfficheEcran;
    if fImpFacXml.Statutimp = [Tcsok]  then
    begin
      BVALIDER.Enabled := True
    end else
    begin
      fImpFacXml.MoveToError (NOMXML.Text);
      AjouteEvent;
    end;
  end;
end;

function TOF_BTIMPORTBSV.ConstitueInfoPiece (TOBDatas : TOB) : boolean;
var LL : string;
		LIB : string;
begin
  result := true;
  LL := TOBDatas.GetString('LIBDOC');
  LIB := LL + ' N° '+TOBDatas.getString('NUMFACTUREFOU')+' du '+
  											DateToStr(TOBDatas.GetDateTime ('DATEFAC'));

  if TOBDatas.getString('CHANTIER') <> '' then
  begin
  	LIB := LIB + ' pour le chantier '+TOBDatas.getString('CHANTIER');
	end;
  INFOPIECE.Caption := LIB;

end;

function TOF_BTIMPORTBSV.PieceExists (Naturepiece,Numero : string) : boolean;
var st : string;
begin
  if not IsNumeric(Numero) then BEGIN Result := False; Exit; end;  
	St := 'SELECT GP_NUMERO FROM PIECE WHERE GP_NATUREPIECEG="'+Naturepiece+'" AND GP_NUMERO='+IntToStr(StrToInt(Numero));
  Result := ExisteSQL(St);
end;

function TOF_BTIMPORTBSV.ConstitueDetail (TOBDatas : TOB) : boolean;
var II : Integer;
		LL : string;
    Loc : string;
    LOk : Boolean;
begin
  result := True;
	For II := 0 to TOBDatas.detail.count -1 do
  begin
    LOk := True;
		LL := rechdom('GCNATUREPIECEG',TOBDatas.detail[II].GetString('NATUREPIECE'),false);
    Loc := LL + ' N° '+TOBDatas.detail[II].GetString('NUMEROPIECE');
		if ll = '' then BEGIN result := false; Lok := false; end;
    if Not PieceExists(TOBDatas.detail[II].GetString('NATUREPIECE'),TOBDatas.detail[II].GetString('NUMEROPIECE')) then
    begin
      Result := false;
      LOk := false;
    end;
    if NOt LOk then Loc := Loc + ' <--- Erreur';
    MDETAIL.Lines.Add(Loc);
  end;
end;

procedure TOF_BTIMPORTBSV.AfficheEcran;
var TOBDatas : TOB;
begin
	TOBDatas := fImpFacXml.TOBDatas;
  //
  ConstitueInfoPiece (TOBDatas);

  THPanel(GetControl('PCONSTITUANT')).Font.Color :=  ColorInit;
  NOMFOU.Font.Color := ColorInit;
  THNumEdit(GetControl('MONTANTHT')).Font.Color := ColorInit;
  THNumEdit(GetControl('MONTANTTTC')).Font.Color := ColorInit;
  INFOPiece.Font.Color := ColorInit;
  //
  CODEFOU.Text := TOBDatas.GetString('CODEFOU');
  NOMFOU.Caption := TOBDatas.GetString('LIBFOU');
  MONTANTHT.Value := TOBDatas.GetDouble('MONTANTHT');
  TVA.Value := TOBDatas.GetDouble('TVA');
  MONTANTTTC.Value := TOBDatas.GetDouble('MONTANTTTC');
  PINTERNE.Visible := True;

  if TCsTiersInconnu in fImpFacXml.Statutimp then
  BEGIN
  	NOMFOU.Font.Color := clRed;
  END;
  //
  if TcsNoDetails in fImpFacXml.Statutimp  then
  BEGIN
  	THPanel(GetControl('PCONSTITUANT')).Font.Color := clRed;
  end else
  begin
    ConstitueDetail (TOBDatas);
  end;
  //
  if TcsMtInvalide in fImpFacXml.Statutimp then
  begin
    THNumEdit(GetControl('MONTANTHT')).Font.Color := clRed;
    THNumEdit(GetControl('MONTANTTTC')).Font.Color := clRed;
  end;

  if TCsDateInvalide in fImpFacXml.Statutimp then
  BEGIN
    INFOPiece.Font.Color := clRed;
  end;

  if fImpFacXml.Statutimp = [TCsOk] then
  begin
    THImage(GetControl('IMGSTATEOK')).Visible := true;
    THImage(GetControl('IMGSTATEBAD')).Visible := false;
  end else
  begin
    THImage(GetControl('IMGSTATEOK')).Visible := false;
    THImage(GetControl('IMGSTATEBAD')).Visible := true;
  end;
end;

Initialization
  registerclasses ( [ TOF_BTIMPORTBSV ] ) ;
end.
