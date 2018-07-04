{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 04/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : PIECEREA_MUL ()
Mots clefs ... : TOF;PIECEREA_MUL
*****************************************************************}
Unit PieceRea_Mul_Tof ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
      MaineAGL,eMul,
{$ELSE}
      db,dbTables,FE_main,Mul,
{$ENDIF}
      forms, sysutils, ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,UTob,Htb97,M3FP,EntGc,
      AssistContreMVtoA ;

Type
  TOF_PIECEREA_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
    procedure ValideSuggestion;
  end ;

Implementation

procedure TOF_PIECEREA_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PIECEREA_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PIECEREA_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PIECEREA_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PIECEREA_MUL.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_PIECEREA_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PIECEREA_MUL.ValideSuggestion;
var TOBPiece, TOBL, TOBCatalogu,TOBCata, TOBCat : TOB;
    NaturepieceG ,Souche : string;
    Numero, IndiceG, i_ind : Integer ;
    QQ : Tquery ;
    CM_Assist : TFCmdeContreM;
    RefCata,RefFour,WhereCat,QCatCol : String ;
begin
NaturepieceG:=TFmul(Ecran).Q.FindField('GP_NATUREPIECEG').asstring ;
Souche:=TFmul(Ecran).Q.FindField('GP_SOUCHE').asstring ;
Numero:=TFmul(Ecran).Q.FindField('GP_NUMERO').asInteger ;
IndiceG:=TFmul(Ecran).Q.FindField('GP_INDICEG').asInteger ;
if (NaturepieceG='') or (Souche='') then exit ;

TOBPiece:=TOB.Create('PIECE',Nil,-1) ;
QQ:=OpenSql('SELECT * From PIECE Where GP_NATUREPIECEG="'+NaturepieceG+'" AND GP_SOUCHE="'+Souche
            +'" AND GP_NUMERO='+IntToStr(Numero)+' AND GP_INDICEG='+IntToStr(IndiceG), True) ;
if not QQ.Eof then TOBPiece.SelectDB('',QQ) else begin TOBPiece.Free; TOBPiece:=Nil; end;
ferme(QQ);
if TOBPiece=Nil then exit ;
QQ:=OpenSql('SELECT * From LIGNE Where GL_NATUREPIECEG="'+NaturepieceG+'" AND GL_SOUCHE="'+Souche
     +'" AND GL_NUMERO='+IntToStr(Numero)+' AND GL_INDICEG='+IntToStr(IndiceG), True) ;
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
      //TOBL.PutValue('GL_QTESTOCK',TOBL.GetValue('GL_QTEFACT'));
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
   TOBCata.Free; TOBCata:=Nil;
   if WhereCat<>'' then
      BEGIN
      QCatCol:='GCA_REFERENCE,GCA_TIERS,GCA_DELAILIVRAISON';
      QQ:=OpenSQL('SELECT '+QCatCol+' FROM CATALOGU WHERE '+WhereCat,True) ;
      if not QQ.Eof then TOBCatalogu.LoadDetailDB('CATALOGU','','',QQ,False,True) ;
      Ferme(QQ) ;
      END ;
   CM_Assist := TFCmdeContreM.Create (Application);
   Try
      CM_Assist.TOBLigne:=TOBPiece ;
      CM_Assist.TOBCatalogu:=TOBCatalogu ;
      CM_Assist.QteATraiterName:='GL_QTESTOCK';
      CM_Assist.Mode:='VTOA';
      CM_Assist.ShowModal;
      if CM_Assist.ModalResult=mrOk then
         BEGIN
         if GetInfoParpiece(NaturePieceG,'GPP_HISTORIQUE')='X' then
            BEGIN
            TOBPiece.PutValue('GP_VIVANTE', '-');
            for i_ind:=0 to TOBPiece.Detail.Count - 1 do
               BEGIN
               TOBL := TOBPiece.Detail[i_ind];
               TOBL.PutValue('GP_VIVANTE', '-') ;
               END;
            TOBPiece.UpdateDBTable ;
            END else TOBPiece.DeleteDB;
         Self.Update;
         END;
   Finally
      CM_Assist.free;
   End;
   end;
TOBPiece.Free; TOBCatalogu.Free
end;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure PieceRea_MulValideSuggestion(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_PIECEREA_MUL) then TOF_PIECEREA_MUL(TOTOF).ValideSuggestion else exit;
end;

Initialization
  registerclasses ( [ TOF_PIECEREA_MUL ] ) ; 
RegisterAglProc('PieceRea_MulValideSuggestion',TRUE,0,PieceRea_MulValideSuggestion);
end.
