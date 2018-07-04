{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 04/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : GCLANCEREA_MUL ()
Mots clefs ... : TOF;LANCEREA_MUL;REAPPRO
*****************************************************************}
Unit LanceRea_Mul_Tof ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
      MaineAGL,eMul,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_main,Mul,DBGrids,
{$ENDIF}
{$IFDEF BTP}
      BtpUtil,
      Menus,
{$ENDIF}
			Ent1,
      forms, sysutils, ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,UTob,Htb97,M3FP,EntGc,Grids,
      FactComm,ParamSoc,AssistContreMVtoA ;

Type
  TOF_LANCEREA_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
    procedure ValideSuggestion;
    procedure BRechResponsable(Sender: TObject);
    {$IFDEF BTP}
    private
    	TheTraitement : string;
      procedure FlisteDblClick (Sender : Tobject);
      procedure ModifSuggestion;
    {$ENDIF}
  end ;

Implementation
uses Utilressource;

procedure TOF_LANCEREA_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_LANCEREA_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_LANCEREA_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_LANCEREA_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_LANCEREA_MUL.OnArgument (S : String ) ;
var CC : THValComboBox;
    CCe : Thedit;
begin
  Inherited ;

  // gestion Etablissement (BTP)
  CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ;
  if CC<>Nil then PositionneEtabUser(CC) ;
  //
  //Gestion Restriction Domaine
  CC:=THValComboBox(GetControl('GP_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

{$IFDEF BTP}
{$IFNDEF EAGLCLIENT}
	THGrid(GetControl('Fliste')).OnDblClick := FlisteDblClick;
{$ELSE}
{$ENDIF}
  TheTraitement := S;
{$ENDIF}


	CCE:=ThEdit(GetControl('AFF_RESPONSABLE')) ;
	if CCE<>Nil then
  begin
  	PositionneResponsableUser(CCE) ;
    CCE.OnElipsisClick := BRechResponsable;
	end;

end ;


procedure TOF_LANCEREA_MUL.BRechResponsable(Sender: TObject);
Var QQ  : TQuery;
    SS  : THCritMaskEdit;
begin

  if GetParamSocSecur('SO_AFRECHRESAV', True) then
  begin
    SS := THCritMaskEdit.Create(nil);
    GetRessourceRecherche(SS,'ARS_RESSOURCE=' + ThEdit(Sender).text + ';TYPERESSOURCE=SAL', '', '');
    if (SS.Text <> THEdit(Sender).text) then
    begin
      if SS.text = '' then ss.text := ThEdit(Sender).text;
    end;
    ThEdit(Sender).text  := SS.Text;
    SS.Free;
  end
  else
    GetRessourceRecherche(ThEdit(Sender),'ARS_TYPERESSOURCE="SAL"', '', '');

end;

procedure TOF_LANCEREA_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_LANCEREA_MUL.ValideSuggestion;
var TOBPiece, TOBL, TOBCatalogu,TOBCata, TOBCat : TOB;
    NaturepieceG ,Souche : string;
    Numero, IndiceG, i_ind : Integer ;
    QQ : Tquery ;
    CM_Assist : TFCmdeContreM;
    RefCata,RefFour,WhereCat,QCatCol,RefPiece : String ;
{$IFDEF BTP}
    Annule : boolean;
{$ENDIF}
begin
{$IFDEF BTP}
Annule := false;
{$ENDIF}
NaturepieceG:=TFmul(Ecran).Q.FindField('GP_NATUREPIECEG').asstring ;
Souche:=TFmul(Ecran).Q.FindField('GP_SOUCHE').asstring ;
Numero:=TFmul(Ecran).Q.FindField('GP_NUMERO').asInteger ;
IndiceG:=TFmul(Ecran).Q.FindField('GP_INDICEG').asInteger ;
if (NaturepieceG='') or (Souche='') then exit ;

TOBPiece:=TOB.Create('PIECE',Nil,-1) ;
QQ:=OpenSql('SELECT * From PIECE Where GP_NATUREPIECEG="'+NaturepieceG+'" AND GP_SOUCHE="'+Souche
            +'" AND GP_NUMERO='+IntToStr(Numero)+' AND GP_INDICEG='+IntToStr(IndiceG), True,-1, '', True) ;
if not QQ.Eof then TOBPiece.SelectDB('',QQ) else begin TOBPiece.Free; TOBPiece:=Nil; end;
ferme(QQ);
if TOBPiece=Nil then exit ;
QQ:=OpenSql('SELECT * From LIGNE Where GL_NATUREPIECEG="'+NaturepieceG+'" AND GL_SOUCHE="'+Souche
     +'" AND GL_NUMERO='+IntToStr(Numero)+' AND GL_INDICEG='+IntToStr(IndiceG), True,-1, '', True) ;
if not QQ.Eof then TOBPiece.LoadDetailDB('LIGNE','','',QQ,False,True) ;
ferme(QQ);

TOBCatalogu:=TOB.Create('',Nil,-1) ;
if TOBPiece.Detail.Count > 0 then
   begin
   TOBCata:=TOB.Create('',nil,-1);
   for i_ind:=0 to TOBPiece.Detail.Count - 1 do
      BEGIN
      TOBL := TOBPiece.Detail[i_ind];
      TOBL.AddChampSup('SELECT', False) ;
      TOBL.PutValue('SELECT', 'X') ;
      RefPiece:=EncodeRefPiece(TOBL);
      //TOBL.PutValue('GL_PIECEORIGINE',RefPiece) ;
      if Trim(TOBL.GetValue('GL_DEPOT')) = '' then
          TOBL.PutValue('GL_DEPOT', GetParamSoc('SO_GCDEPOTDEFAUT'));
      TOBL.PutValue('GL_PIECEORIGINE','') ;
      TOBL.PutValue('GL_PIECEPRECEDENTE','') ;
      RefCata:=TOBL.GetValue('GL_REFARTTIERS');
      TOBL.PutValue('GL_REFCATALOGUE',RefCata);
      RefFour:=TOBL.GetValue('GL_FOURNISSEUR');
      if (RefFour='') or (RefCata='') then continue;
      if TOBCata.FindFirst(['REFCATA','REFFOUR'],[RefCata,RefFour],False) = Nil then
         BEGIN
         TOBCat:=TOB.Create('',TOBCata,-1);
         TOBCat.AddChampSup('REFCATA',False);
         TOBCat.PutValue('REFCATA', RefCata);
         TOBCat.AddChampSup('REFFOUR',False);
         TOBCat.PutValue('REFFOUR', RefFour);
         END;
      END;
   WhereCat:='';
   for i_ind:=0 to TOBCata.Detail.Count-1 do
      BEGIN
      RefCata:=TOBCata.Detail[i_ind].GetValue('REFCATA');
      RefFour:=TOBCata.Detail[i_ind].GetValue('REFFOUR');
      if WhereCat <> '' then WhereCat:=WhereCat+' OR ';
      WhereCat:=WhereCat + '(GCA_REFERENCE="'+RefCata+'" AND GCA_TIERS="'+RefFour+'")';
      END;
   TOBCata.Free; //TOBCata:=Nil;
   if WhereCat<>'' then
      BEGIN
      QCatCol:='GCA_REFERENCE,GCA_TIERS,GCA_DELAILIVRAISON';
      QQ:=OpenSQL('SELECT '+QCatCol+' FROM CATALOGU WHERE '+WhereCat,True,-1, '', True) ;
      if not QQ.Eof then TOBCatalogu.LoadDetailDB('CATALOGU','','',QQ,False,True) ;
      Ferme(QQ) ;
      END ;
   CM_Assist := TFCmdeContreM.Create (Application);
   Try
      CM_Assist.TOBLigne:=TOBPiece ;
      CM_Assist.TOBCatalogu:=TOBCatalogu ;
      CM_Assist.QteATraiterName:='GL_QTESTOCK';
      CM_Assist.Mode:='REAP';
      CM_Assist.ShowModal;
      if CM_Assist.ModalResult=mrOk then
         BEGIN
         if GetInfoParpiece(NaturePieceG,'GPP_HISTORIQUE')='X' then
            BEGIN
            TOBPiece.PutValue('GP_VIVANTE', '-');
            for i_ind:=0 to TOBPiece.Detail.Count - 1 do
               BEGIN
               TOBL := TOBPiece.Detail[i_ind];
               TOBL.PutValue('GL_VIVANTE', '-') ;  // DBR : GL_VIVANTE a la place de GP_VIVANTE en meme tps que fiche 10424
               END;
            TOBPiece.UpdateDB(false);
            END else TOBPiece.DeleteDB;
         Self.Update;
         END else
         BEGIN
{$IFDEF BTP}
         Annule := true;
{$ENDIF}
         END;
   Finally
      CM_Assist.free;
   End;
   end;
//TOBPiece.Free; TOBCatalogu.Free;
TFMul(Ecran).BChercheClick (Self); // DBR Fiche eQualite : 10424
{$IFDEF BTP}
//if not annule then if ISGenereLivFromAppro then GenereLivraisonClients(false);
{$ENDIF}
end;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure LanceRea_MulValideSuggestion(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_LANCEREA_MUL) then TOF_LANCEREA_MUL(TOTOF).ValideSuggestion else exit;
end;

{$IFDEF BTP}
procedure TOF_LANCEREA_MUL.FlisteDblClick(Sender: Tobject);
begin
	if TheTraitement = 'MODIFICATION' Then ModifSuggestion
  																	else ValideSuggestion;
end;


procedure TOF_LANCEREA_MUL.ModifSuggestion;
begin
  TMenuItem(GetCOntrol('MnZPiece')).Click;
end;

{$ENDIF}

Initialization
  registerclasses ( [ TOF_LANCEREA_MUL ] ) ;
RegisterAglProc('LanceRea_MulValideSuggestion',TRUE,0,LanceRea_MulValideSuggestion);
end.
