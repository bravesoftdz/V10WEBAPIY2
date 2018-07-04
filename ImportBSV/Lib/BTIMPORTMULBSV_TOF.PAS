{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 19/05/2014
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTIMPORTMULBSV ()
Mots clefs ... : TOF;BTIMPORTMULBSV
*****************************************************************}
Unit BTIMPORTMULBSV_TOF ;

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
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UImportFacXml,
     FileCtrl,
     AGLInit,
     UTOF ;

Type
  TOF_BTIMPORTMULBSV = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TOBGenere : TOB;
    TOBRetour : TOB;
    fImpXml : TImportFachXml;
    MCompteRendu : Tmemo;
    procedure TraiteFile (nomFic : string);
    procedure TraiteFichier (nomfic : string);
    procedure AjouteEvent(FF: String);
  end ;

Implementation
uses factgrp;

procedure TOF_BTIMPORTMULBSV.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTMULBSV.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTMULBSV.OnUpdate ;
var nomfic : string;
Rec : TSearchRec;
begin
  Inherited ;
  THEdit(getControl('EMPLACEMENT')).Text := IncludeTrailingBackslash(THEdit(getControl('EMPLACEMENT')).Text);
  if PGIAsk('Vous allez intégrer en lot des factures d''achats.Confirmez-vous?')<>mryes then exit;
  Nomfic:= IncludeTrailingBackslash(THEdit(getControl('EMPLACEMENT')).Text)+'*.xml';
  if FindFirst (Nomfic,faAnyFile,Rec) = 0 then
  begin
    if (rec.name <> '.') and (rec.name <> '..') then TraiteFile (rec.Name);
    while FindNext (Rec) = 0 do
    begin
    	if (rec.name <> '.') and (rec.name <> '..') then TraiteFile (rec.Name);
    end;
  end;
  FindClose (Rec);
end ;

procedure TOF_BTIMPORTMULBSV.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTMULBSV.OnArgument (S : String ) ;
begin
  Inherited ;
  fImpXml := TImportFachXml.create;
  THEdit(getControl('EMPLACEMENT')).Text := fImpXml.RepertDef;
  TMemo(getControl('MCONTERENDU')).Clear;
  MCompteRendu := TMemo(getControl('MCONTERENDU'));
  TOBGenere := TOB.Create('LES PIECES',nil,-1);

end ;

procedure TOF_BTIMPORTMULBSV.OnClose ;
begin
	TOBGenere.free;

  fImpXml.Free;
  Inherited ;
end ;

procedure TOF_BTIMPORTMULBSV.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTMULBSV.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_BTIMPORTMULBSV.AjouteEvent (FF : String);
var QQ: TQuery;
  MotifPiece: TStrings;
  NumEvent: integer;
  II : integer;
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
  MotifPiece.Add('Document : '+FF);
  QQ.FindField('GEV_DATEEVENT').AsDateTime := Date;
  QQ.FindField('GEV_UTILISATEUR').AsString := V_PGI.User;
  if fImpXml.Statutimp = [TCsOk] then
  begin
    QQ.FindField('GEV_ETATEVENT').AsString := 'OK';
    if TOBRetour <> nil then
    begin
      for II := 0 to TOBretour.detail.count -1 do
      begin
        MotifPiece.Add(rechdom('GCNATUREPIECEG',TOBRetour.detail[ii].GetString('NATUREPIECE'),false)+
                       ' N° '+InttOstr(TOBRetour.detail[ii].GetInteger('NUMERO'))+' Générée');
      end;
    end;
  end else
  begin
    QQ.FindField('GEV_ETATEVENT').AsString := 'ERR';
    fImpXml.ContitueMotif (MotifPiece);
  end;
    TMemoField(QQ.FindField('GEV_BLOCNOTE')).Assign(MotifPiece);
  QQ.Post;
  Ferme(QQ);
  MotifPiece.Free;
end;

procedure TOF_BTIMPORTMULBSV.TraiteFile(nomFic: string);
var ff : string;
begin
  ff := IncludeTrailingBackslash(ExtractFilePath(THEdit(getControl('EMPLACEMENT')).Text))+Nomfic;
	//
  TraiteFichier(ff);

end;

procedure TOF_BTIMPORTMULBSV.TraiteFichier(nomfic: string);
var Datefac : TDateTime;
		sMontantHt,sMontantTTC : double;
    Sql : string;
    QQ : TQuery;
    TOBP,TOBDatas : TOB;
    Mess : string;
    i_ret,II : Integer;
    Okok : Boolean;
begin
  okok := True;
  TMemo(getControl('MCONTERENDU')).Lines.Add('----------------------------------------------------------------------------------------------------------------');
  TMemo(getControl('MCONTERENDU')).Lines.Add('Traitement du fichier '+Nomfic);
  TMemo(getControl('MCONTERENDU')).Lines.Add('----------------------------------------------------------------------------------------------------------------');
  fImpXml.TraiteXml(NomFic);
  if not (fImpXml.Statutimp = [TCsOk])  then
  begin
  	fImpXml.MoveToError (NomFic);
    fImpXml.ContitueMotif (MCompteRendu.Lines);
    MCompteRendu.Lines.Add('Fichier rejeté');
  	AjouteEvent (nomfic);
    Exit;
  end;
  TOBDatas := fImpXml.TOBDatas;
  TOBGenere.clearDetail;
  DateFac := TOBDatas.GetDateTime('DATEFAC');
  //
  SMontantHt := 0;
  SMontantTTC := 0;
  for II := 0 to TOBDatas.detail.Count -1 do
  begin
    Sql := 'SELECT * FROM PIECE WHERE GP_NATUREPIECEG="'+TOBDatas.detail[II].GetString('NATUREPIECE')+'" AND '+
    			 'GP_NUMERO='+IntToStr(TOBDatas.detail[II].GetInteger('NUMEROPIECE'));
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
  (*
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
    TOBretour := nil;
    fImpXml.ContitueMotif (MCompteRendu.lines);
    i_ret := RegroupeLesPieces (TOBGenere,TOBDatas.GetString('TYPEDOC'),False,false,True,0,
                        				DateFac,false,True,false,False,false,'',false,'','GENEREFF',TOBDatas.GetDouble('MONTANTHT'),
                        				TOBDatas.GetDouble('MONTANTTTC'),TOBDatas.GetString('NUMFACTUREFOU'));
    if TheTOB <> nil then TOBRetour := TheTOB;
    if i_ret = 0 then
    begin
      fImpXml.MoveToTreated (Nomfic);
    end else
    begin
      fImpXml.MoveToError (Nomfic);
    end;
    AjouteEvent (nomfic);
    if TOBRetour <> nil then
    begin
      for II := 0 to TOBretour.detail.count -1 do
      begin
        MCompteRendu.Lines.Add(rechdom('GCNATUREPIECEG',TOBRetour.detail[ii].GetString('NATUREPIECE'),false)+
                               ' N° '+InttOstr(TOBRetour.detail[ii].GetInteger('NUMERO'))+' Généré');
      end;
    end;
    if TheTOB <> nil then
    begin
    	FreeAndNil(TheTOB);
      TOBRetour := nil;
    end;
  end else
  begin
    MCompteRendu.Lines.Add('En attente de contrôle');
    MCompteRendu.Lines.Add('Document non généré');
  end;
end;

Initialization
  registerclasses ( [ TOF_BTIMPORTMULBSV ] ) ;
end.
