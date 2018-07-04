unit UTofGenePiece;

interface

uses  uTofAfBaseCodeAffaire,StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, UTOM,AGLInit,
      Utob,HDB,Messages,HStatus,Ent1,EntGC,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
      Fiche,mul, DBGrids,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ENDIF}
      M3FP,UTofGCDatePiece,FactGrp;

Type

     TOF_GCGenePiece = Class (TOF_AFBASECODEAFFAIRE)
     private
        PieceGenere : string ;
        procedure GenererPiece ;
        procedure ChargeLaPiece ;
        procedure ChargeTobPieces ;
        procedure NomsChampsAffaire( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ; override ;
     public
        procedure OnArgument (Arguments : String ) ; override ;
        procedure OnLoad; override;
     END ;

procedure AGLGenererPiece(parms:array of variant; nb: integer ) ;

implementation

uses
  ParamSoc
  ,wCommuns  
  {$IFDEF GPAOLIGHT}
    ,wOrdreCMP
    ,FactTob
  {$ENDIF GPAOLIGHT}
  ;

Procedure TOF_GCGenePiece.NomsChampsAffaire ( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ;
BEGIN
Aff:=THEdit(GetControl('GP_AFFAIRE'))   ; Aff0:=Nil ;
Aff1:=THEdit(GetControl('GP_AFFAIRE1')) ;
Aff2:=THEdit(GetControl('GP_AFFAIRE2')) ;
Aff3:=THEdit(GetControl('GP_AFFAIRE3')) ;
Aff4:=THEdit(GetControl('GP_AVENANT'))  ;
Tiers:=THEdit(GetControl('GP_TIERS'))   ;
END ;

Procedure TOF_GCGenePiece.OnArgument (Arguments : String ) ;
Var CC : THValComboBox ;
    CB : TCheckBox ;
    EclateDomaine : boolean ;
    St,Arg : string ;
begin
inherited ;
Arg:='';
while Arguments<>'' do
      BEGIN
      St:=Uppercase(ReadTokenSt(Arguments)) ;
      if St='VENTE' then SetControlProperty('GP_TIERS','DataType','GCTIERSCLI')
        else if St='ACHAT' then SetControlProperty('GP_TIERS','DataType','GCTIERSFOURN')
          else Arg:=St;
      END;
//PieceGenere:=Arguments ;
PieceGenere:=Arg ;
EclateDomaine:=(GetInfoParPiece(PieceGenere,'GPP_ECLATEDOMAINE')='X') ;
CB:=TCheckBox(GetControl('ECLATEDOMAINE')) ; if CB<>Nil then CB.Checked:=EclateDomaine ;

if Not(ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) then
  SetControlVisible ('TGP_AFFAIRE1', False);

CC:=THValComboBox(GetControl('GP_DOMAINE')) ; if CC<>Nil then PositionneDomaineUser(CC) ;
CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ; if CC<>Nil then PositionneEtabUser(CC) ;

if Not(ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) then
   begin
   SetControlVisible ('TGP_AFFAIRE1', False);
   SetControlVisible ('GP_AFFAIRE1', False);
   SetControlVisible ('GP_AFFAIRE2', False);
   SetControlVisible ('GP_AFFAIRE3', False);
   SetControlVisible ('GP_AVENANT', False);
   end;
end;

Procedure TOF_GCGenePiece.OnLoad ;
begin
inherited ;
if (ctxAffaire in V_PGI.PGIContexte) And (GetControlText('GP_NATUREPIECEG')='CC') then
    SetControlText('XX_WHERE','GP_AFFAIRE=""');
end;

procedure TOF_GCGenePiece.ChargeLaPiece ;
var TOB_Piece :TOB;
    NaturepieceG ,Souche : string;
    Numero, IndiceG : Integer ;
    QQ : Tquery ;
Begin
NaturepieceG:=TFmul(Ecran).Q.FindField('GP_NATUREPIECEG').asstring ;
Souche:=TFmul(Ecran).Q.FindField('GP_SOUCHE').asstring ;
Numero:=TFmul(Ecran).Q.FindField('GP_NUMERO').asInteger ;
IndiceG:=TFmul(Ecran).Q.FindField('GP_INDICEG').asInteger ;

Tob_Piece:=TOB.Create('PIECE', LaTob, -1) ;
QQ:=OpenSql('SELECT * From PIECE Where GP_NATUREPIECEG="'+NaturepieceG+'" AND GP_SOUCHE="'+Souche
     +'" AND GP_NUMERO='+IntToStr(Numero)+' AND GP_INDICEG='+IntToStr(IndiceG), TRUE) ;
Tob_Piece.SelectDB('',QQ);
ferme(QQ);
end;


//////////////////////// Construire tob detail des pièces ///////////////////////////
procedure TOF_GCGenePiece.ChargeTobPieces ;
var i : integer;
begin
if HShowMessage('0;Confirmation;Confirmez vous la génération des pièces ?;Q;YN;N;N;','','')<>mrYes then exit;
With TFMul(Ecran) do
   BEGIN
   if FListe.AllSelected then
      BEGIN
      InitMove(Q.RecordCount,'');
      Q.First;
      while Not Q.EOF do
        BEGIN
        MoveCur(False);
        ChargeLaPiece;
        Q.NEXT;
        END;
      FListe.AllSelected:=False;
      END ELSE
      BEGIN
      InitMove(FListe.NbSelected,'');
      for i:=0 to FListe.NbSelected-1 do
          BEGIN
          FListe.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
          Q.TQ.Seek(FListe.Row-1) ;
{$ELSE}
{$ENDIF}
          ChargeLaPiece ;
          MoveCur(False);
          END ;
      FListe.ClearSelected;
      END;
   FiniMove;
   END ;
End;

//////////// Modification par lot des articles ///////////////
procedure TOF_GCGenePiece.GenererPiece ;
var CB,CEclate : TCheckBox ;
     ii : integer ;
     {$IFDEF GPAOLIGHT}
       i: Integer;
       TobGP: Tob;
       FoundOrdre: Boolean;
     {$ENDIF GPAOLIGHT}
     DatePiece : TDateTime ;
begin
With TFMul(Ecran) do
   BEGIN
   if (FListe.NbSelected=0) and (not FListe.AllSelected) then
      begin
      MessageAlerte('Aucun élément sélectionné');
      exit;
      end;
   DatePiece:=ChoixDatePiece ;
   LaTOB:=TOB.Create('_GenererPiece',NIL,-1);
   ChargeTobPieces ;
   {$IFDEF GPAOLIGHT}
   { Gestion complète des ordres à la commande ... }
   { dans le cas de livraison de commande de vente }
   FoundOrdre := False;
   if Pos(PieceGenere + ';', GetParamSoc('SO_WMISEENPROD')) <> 0 then
   begin
     i := 0;
     while not FoundOrdre and (i < LaTob.Detail.Count) do
     begin
       TobGP := LaTob.Detail[i];
       if  (TobGP.G('GP_NATUREPIECEG') = GetParamSoc('SO_WORDREVTE'))
       and wLoadTobFromSQL('LIGNE', 'SELECT GL_TYPEREF, GL_ARTICLE, GL_NATUREPIECEG, GL_SOUCHE, GL_NUMERO, GL_INDICEG, GL_NUMORDRE, GL_IDENTIFIANTWOL FROM LIGNE WHERE ' + WherePiece(Tob2CleDoc(TobGP), ttdLigne, False), TobGP)
       then
         FoundOrdre := wExisteOrdreCompletementGere_NonGenere(TobGP);
       TobGP.ClearDetail;
       Inc(i);
     end;
   end;
   {$ENDIF GPAOLIGHT}
   CB:=TCheckBox(FindComponent('DEGROUPEREMISE')) ;
   CEClate:=TCheckBox(FindComponent('ECLATEDOMAINE')) ;
   if CEclate <> Nil then ii:=Ord(CEclate.Checked) ;
   {$IFDEF GPAOLIGHT}
     if not FoundOrdre then
       RegroupeLesPieces(LaTob,PieceGenere,False,CB.Checked,True,ii,DatePiece) ;
   {$ELSE}
     RegroupeLesPieces(LaTob,PieceGenere,False,CB.Checked,True,ii,DatePiece) ;
   {$ENDIF GPAOLIGHT}
   LaTob.Free ; LaTob:=nil ;
   END ;
end;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLGenererPiece(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_GCGenePiece) then TOF_GCGenePiece(TOTOF).GenererPiece else exit;
end;

Initialization
registerclasses([TOF_GCGenePiece]);
RegisterAglProc('GenererPiece',TRUE,0,AGLGenererPiece);
end.
