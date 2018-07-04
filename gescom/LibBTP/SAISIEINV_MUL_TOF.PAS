{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 11/06/2014
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : SAISIEINV_MUL ()
Mots clefs ... : TOF;SAISIEINV_MUL
*****************************************************************}
Unit SAISIEINV_MUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,DBGrids,HDB,M3Fp,AglInit,FE_Main,
{$else}
     eMul,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HTB97,
     HMsgBox,
     uTOFComm,
     UTOF ;

Type
  TOF_SAISIEINV_MUL = Class (tTOFComm)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    fListe : THDBGrid;
    procedure ReactiverClick (Sender : TObject);
  	procedure DupliquerClick(Sender : TObject);
    procedure RowEnterGrid (Sender : TObject);
    procedure ListeDblClick (sender : TObject);
    procedure EnregEvtreactiveListe(NomListe: string);
    function ConstitueCodeInv(Racine: string; var CodeInv: string): boolean;
    procedure EnregEvtDupListe(NomListe, NouvelleListe: string; Dateliste: TdateTime);
    function DupliqueListe ( CodeListe,CodeInvent: string; DateInv: TDateTime; Intitule : string) : boolean;
    procedure GetInfosArt(TL, TD, TA: TOB);

  end ;

Implementation
uses SaisieInv,UTofListeInv;

procedure TOF_SAISIEINV_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEINV_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEINV_MUL.OnUpdate ;
begin
  Inherited ;
  if Assigned(fListe) then RowEnterGrid(fliste);
end ;

procedure TOF_SAISIEINV_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEINV_MUL.OnArgument (S : String ) ;
begin
  Inherited ;
  if  (V_PGI.SAV) then TToolbarButton97 (GetControl('BREACTIVER')).Visible := True;
  TToolbarButton97 (GetControl('BREACTIVER')).OnClick := ReactiverClick;
  TToolbarButton97 (GetControl('B_DUPLICATION')).OnClick := DupliquerClick;
  fListe := THDBGrid(GetControl('Fliste'));
  fListe.OnRowEnter := RowEnterGrid;
  fListe.OnDblClick := ListeDblclick;
end ;

procedure TOF_SAISIEINV_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEINV_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEINV_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEINV_MUL.DupliquerClick(Sender: TObject);
var CodeListe : string;
		Intitule,RACINECODE,CodeInvent : string;
    DateInv : TDateTime;
    OneTOB : TOB;
begin
  OneTOB := TOB.Create('UN ECHANGE',nil,-1);
  OneTOB.AddChampSupValeur('INTITULE','');
  OneTOB.AddChampSupValeur('RACINECODE','');
  OneTOB.AddChampSupValeur('DATEINV',Now);
  OneTOB.AddChampSupValeur('OKOK','-');
	CodeListe:=Fliste.datasource.dataset.FindField('GIE_CODELISTE').AsString;
  TheTOB := OneTOB;
  AGLLanceFiche('BTP','BTDUPLICINVENT','','','');
  TheTOB := nil;
  if OneTOB.GetString('OKOK')='X' then
  begin
    DateInv := OneTOB.GetDateTime('DATEINV');
    RACINECODE := OneTOB.GetString('RACINECODE');
    Intitule := OneTOB.GetString('INTITULE');
    if ConstitueCodeInv (RACINECODE,CodeInvent ) then
    begin
      if DupliqueListe ( CodeListe,CodeInvent,DateInv,Intitule) then
      begin
				EnregEvtDupListe (CodeListe,CodeInvent,DateInv);
        refreshDb;
      end else
      begin
        PgiError ('Erreur ecriture de la liste'); 
      end;
    end else
    begin
      PGIError('Impossible : plus de code disponible pour cette racine de liste');
    end;
  end;
  OneTOB.Free;
end;

procedure TOF_SAISIEINV_MUL.ReactiverClick(Sender: TObject);
var CodeListe : string;
		DateInv: TdateTime;
begin
	CodeListe:=Fliste.datasource.dataset.FindField('GIE_CODELISTE').AsString;
  if Fliste.datasource.dataset.FindField('GIE_DATEINVENTAIRE')= nil then
  begin
    PgiInfo ('Vous devez ajouter la date d''inventaire dans la liste de saisie');
  end;
	DateInv:=Fliste.datasource.dataset.FindField('GIE_DATEINVENTAIRE').AsDateTime;

	if PGIAsk('ATTENTION : Vous allez réactiver une liste d''inventaire déjà validée.#13#10Etes-vous sur ?')<>Mryes then exit;
  ExecuteSQL('UPDATE LISTEINVENT SET GIE_VALIDATION="-" WHERE GIE_CODELISTE="'+CodeListe+'"');
  ExecuteSQL('UPDATE LISTEINVLIG SET GIL_DATESAISIE="'+USDATETIME(DateInv)+'" WHERE GIL_CODELISTE="'+CodeListe+'" AND GIL_DATESAISIE IS NULL');
  EnregEvtreactiveListe (CodeListe);
  refreshDb;

end;

procedure TOF_SAISIEINV_MUL.RowEnterGrid(Sender: TObject);
var Valide : string;
begin
  TToolbarButton97 (GetControl('BREACTIVER')).Enabled := false;
	Valide:=Fliste.datasource.dataset.FindField('GIE_VALIDATION').AsString;
  if (Valide = 'X') and (V_PGI.SAV) then TToolbarButton97 (GetControl('BREACTIVER')).Enabled := True;
end;

procedure TOF_SAISIEINV_MUL.ListeDblClick(sender: TObject);
var CodeListe : string;
begin
	CodeListe:=Fliste.datasource.dataset.FindField('GIE_CODELISTE').AsString;
	EntreeSaisieInv(CodeListe);
  RefreshDB;
end;

procedure TOF_SAISIEINV_MUL.EnregEvtDupListe (NomListe,NouvelleListe : string; Dateliste : TdateTime);
var QQ: TQuery;
  MotifPiece: TStrings;
  NumEvent: integer;
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
  QQ.FindField('GEV_TYPEEVENT').AsString := 'INV';
  QQ.FindField('GEV_LIBELLE').AsString := 'Duplication de liste d''inventaire';
  QQ.FindField('GEV_DATEEVENT').AsDateTime := Date;
  QQ.FindField('GEV_UTILISATEUR').AsString := V_PGI.User;
  QQ.FindField('GEV_ETATEVENT').AsString := 'OK';
  MotifPiece.Add('Liste d''inventaire '+NomListe+' --> '+NouvelleListe);
  TMemoField(QQ.FindField('GEV_BLOCNOTE')).Assign(MotifPiece);
  QQ.Post;
  Ferme(QQ);
  MotifPiece.Free;
end;

procedure TOF_SAISIEINV_MUL.EnregEvtreactiveListe(NomListe: string);
var QQ: TQuery;
  MotifPiece: TStrings;
  NumEvent: integer;
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
  QQ.FindField('GEV_TYPEEVENT').AsString := 'INV';
  QQ.FindField('GEV_LIBELLE').AsString := 'Réactivation de liste d''inventaire';
  QQ.FindField('GEV_DATEEVENT').AsDateTime := Date;
  QQ.FindField('GEV_UTILISATEUR').AsString := V_PGI.User;
  QQ.FindField('GEV_ETATEVENT').AsString := 'OK';
  MotifPiece.Add('Liste d''inventaire '+NomListe);
  TMemoField(QQ.FindField('GEV_BLOCNOTE')).Assign(MotifPiece);
  QQ.Post;
  Ferme(QQ);
  MotifPiece.Free;
end;

function TOF_SAISIEINV_MUL.ConstitueCodeInv(Racine: string; var CodeInv : string): boolean;
var FnextCode : Integer;
		FirstTour : Boolean;
begin
  result := True;
  FirstTour := True;
  FnextCode := 1;
	while True do
  begin
		if Not ExisteSQL('SELECT GIE_CODELISTE FROM LISTEINVENT WHERE GIE_CODELISTE="' + Racine + Format('%.6d',[FNextCode]) + '"') then
    	Break;
    if FnextCode  > 999999 then
    begin
      if Not FirstTour then
      BEGIN
        result := false;
      	break;
      end;
      FirstTour := False;
    	FnextCode := 0;
    end;
  	inc(FNextCode);
  end;
  if result then CodeInv := Racine + Format('%.6d',[FNextCode]);
end;

procedure TOF_SAISIEINV_MUL.GetInfosArt (TL,TD,TA : TOB);
var QQ : TQuery;
begin
  TD.InitValeurs;
  TA.InitValeurs;
  QQ := OpenSQL('SELECT * FROM DISPO WHERE GQ_ARTICLE="'+TL.GetString('GIL_ARTICLE')+'" AND '+
  							'GQ_DEPOT="'+TL.GetString('GIL_DEPOT')+'"',true,1,'',true	);
  if not QQ.eof then
  begin
    TD.SelectDB('',QQ);
  end;
  Ferme(QQ);
  QQ := OpenSQL('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+TL.GetString('GIL_ARTICLE')+'"',true,1,'',true	);
  if not QQ.eof then
  begin
    TA.SelectDB('',QQ);
  end;
	ferme (QQ);
end;

function TOF_SAISIEINV_MUL.DupliqueListe(CodeListe, CodeInvent: string; DateInv: TDateTime; Intitule: string) : boolean;
var TT,TL,TD,TA : TOB;
		QQ : TQuery;
    II : Integer;
    FUS,FUV,FPPQ : Double;
begin
  Result := True;
  FPPQ:=1;
  //
	TT := TOB.Create ('LISTEINVENT',nil,-1);
  TD := TOB.Create ('DISPO',nil,-1);
  TA := TOB.Create ('ARTICLE',nil,-1);
  TRY
    QQ := OpenSQL('SELECT * FROM LISTEINVENT WHERE GIE_CODELISTE="'+CodeListe+'"',True,1,'',true);
    if not QQ.eof then
    begin
      TT.SelectDB('',QQ);
      TT.SetString('GIE_CODELISTE',CodeInvent);
      TT.SetString('GIE_LIBELLE',Intitule);
      TT.SetDateTime('GIE_DATEINVENTAIRE',DateInv);
      TT.SetDateTime('GIE_DATECREATION',nowh);
      TT.SetDateTime('GIE_DATEMODIF',nowh);
      TT.SetString ('GIE_CREATEUR',V_PGI.USER);
      TT.SetString('GIE_UTILISATEUR',V_PGI.USER);
      TT.SetString('GIE_VALIDATION','-');
    end else Result := false;
    ferme (QQ);
    if Result then
    begin
      QQ := OpenSQL('SELECT * FROM LISTEINVLIG WHERE GIL_CODELISTE="'+CodeListe+'"',True,-1,'',true);
      if not QQ.Eof then
      begin
        TT.LoadDetailDB('LISTEINVLIG','','',QQ,false);
      end else Result := false;
      Ferme(QQ);
    end;
    //
    if Result then
    begin
      for II := 0 to TT.Detail.Count -1 do
      begin
        TL := TT.detail[II];
        GetInfosArt (TL,TD,TA);
        //
        FUS := RatioMesure('PIE', TA.GetString('GA_QUALIFUNITESTO'));
        FUV := RatioMesure('PIE', TA.GetString('GA_QUALIFUNITEVTE')); if FUV=0 then FUV:=1;
        //
        TL.SetString('GIL_CODELISTE',CodeInvent);
        TL.SetString('GIL_SAISIINV','-');
        TL.SetDouble('GIL_INVENTAIRE',0);
        TL.SetDouble('GIL_QTEPHOTOINV',TD.GetDouble('GQ_PHYSIQUE'));
        //
        TL.SetDouble('GIL_DPA', TD.GetDouble('GQ_DPA'));
        TL.SetDouble('GIL_PMAP', TD.GetDouble('GQ_PMAP'));
        TL.SetDouble('GIL_DPR', TD.GetDouble('GQ_DPR'));
        TL.SetDouble('GIL_PMRP', TD.GetDouble('GQ_PMRP'));
        //
        TL.SetDouble('GIL_DPAART', Ratioize(TA.GetDouble('GA_DPA'), FUS, FUV, FPPQ));
        TL.SetDouble('GIL_PMAPART', Ratioize(TA.GetDouble('GA_PMAP'), FUS, FUV, FPPQ));
        TL.SetDouble('GIL_DPRART', Ratioize(TA.GetDouble('GA_DPR'), FUS, FUV, FPPQ));
        TL.SetDouble('GIL_PMRPART', Ratioize(TA.GetDouble('GA_PMRP'), FUS, FUV, FPPQ));
        //
        TL.SetDouble('GIL_DPASAIS', 0);
        TL.SetDouble('GIL_PMAPSAIS', 0);
        TL.SetDouble('GIL_DPRSAIS', 0);
        TL.SetDouble('GIL_PMRPSAIS', 0);
        //
        TL.SetDateTime('GIL_DATESAISIE', DateInv);
      end;
      if Not TT.InsertDBByNivel(false) then Result :=false;
    end;
  FINALLY
  	TT.Free;
    TD.Free;
    TA.Free;
  end;
end;

Initialization
  registerclasses ( [ TOF_SAISIEINV_MUL ] ) ;
end.
