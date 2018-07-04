{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 14/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RTSELECTRESSOURCES ()
Mots clefs ... : TOF;RTSELECTRESSOURCES
*****************************************************************}
Unit Utof_RTSelectRessources ;

Interface

uses  StdCtrls,Controls,Classes,forms,sysutils,Windows,HTB97,
      HCtrls,HEnt1,HMsgBox,UTOF,
      UTOM,AGLInit,Hstatus,M3FP,UTOB,Hqry,
{$IFNDEF EAGLCLIENT}
     db,HDB,
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     mul,
{$else}
     emul,     
{$ENDIF}
     ComCtrls
     ,TiersUtil
     ;

Type
  TOF_RTSELECTRESSOURCES = Class (TOF)
    G1 : THGRID;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure DeplaceLesRessources;
    procedure RabRessources;
    procedure bSuppClick(Sender: TObject);
    procedure GenereActionsRessources;
    Function PresenceRessource( StRessource : String ) : Boolean ;
  end ;

Implementation

procedure TOF_RTSELECTRESSOURCES.DeplaceLesRessources;
var  F : TFMul ;
     i : integer;
{$IFDEF EAGLCLIENT}
       L : THGrid;
{$ELSE}
       L : THDBGrid;
{$ENDIF}
     Q : THQuery;
begin
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
   begin
   PGIInfo('Aucun élément sélectionné','');
   exit;
   end;

{$IFDEF EAGLCLIENT}
if F.bSelectAll.Down then
   if not F.Fetchlestous then
     begin
     F.bSelectAllClick(Nil);
     F.bSelectAll.Down := False;
     exit;
     end;
{$ENDIF}

L:= F.FListe;
Q:= F.Q;

if L.AllSelected then
   begin
   InitMove(Q.RecordCount,'');
   Q.First;
   while Not Q.EOF do
      begin
      MoveCur(False);
      if ( Q.FindField('ARS_RESSOURCE').asstring <> '') and ( not PresenceRessource(Q.FindField('ARS_RESSOURCE').asstring) ) then
      begin
        G1.cells[1,G1.RowCount-1]:=Q.FindField('ARS_RESSOURCE').asstring;
        G1.cells[2,G1.RowCount-1]:=Q.FindField('ARS_LIBELLE').asstring;
        G1.cells[3,G1.RowCount-1]:=Q.FindField('ARS_LIBELLE2').asstring;
        G1.RowCount:=G1.RowCount+1;
      end;
      Q.Next;
      end;
   L.AllSelected:=False;
   end else
   begin
   InitMove(L.NbSelected,'');
   for i:=0 to L.NbSelected-1 do
    begin
      MoveCur(False);
      L.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
      Q.TQ.Seek(L.Row-1) ;
{$ENDIF}
      if ( Q.FindField('ARS_RESSOURCE').asstring <> '') and ( not PresenceRessource(Q.FindField('ARS_RESSOURCE').asstring) ) then
      begin
        G1.cells[1,G1.RowCount-1]:=Q.FindField('ARS_RESSOURCE').asstring;
        G1.cells[2,G1.RowCount-1]:=Q.FindField('ARS_LIBELLE').asstring;
        G1.cells[3,G1.RowCount-1]:=Q.FindField('ARS_LIBELLE2').asstring;
        G1.RowCount:=G1.RowCount+1;
      end;
    end;
   L.ClearSelected;
   end;

if F.bSelectAll.Down then
    F.bSelectAll.Down := False;

FiniMove;

end;

Function TOF_RTSELECTRESSOURCES.PresenceRessource( StRessource : String ) : Boolean ;
var i : integer;
begin
  result:=False;
  for i:=1 to G1.RowCount-1 Do
  begin
    if G1.cells[1,i] = StRessource then
    begin
      Result:=True;
      break;
    end;
  end;
end ;

procedure TOF_RTSELECTRESSOURCES.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RTSELECTRESSOURCES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_RTSELECTRESSOURCES.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_RTSELECTRESSOURCES.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_RTSELECTRESSOURCES.OnArgument (S : String ) ;
var bSupp : TToolBarButton97;
begin
  Inherited ;
  //G1LesColonnes:='FIXED;C_NOM;C_FONCTION;C_TELEPHONE' ;
  G1:=THGRID(GetControl('LIGRES'));
  G1.RowCount:=2;
  G1.ColCount:=4;
  G1.FixedCols:=1;
  G1.FixedRows:=1;
  G1.ColWidths[0]:=12;
  bSupp := TToolBarButton97(GetControl('BSUPP'));
  bSupp.OnClick := bSuppClick;
end ;

procedure TOF_RTSELECTRESSOURCES.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_RTSELECTRESSOURCES.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_RTSELECTRESSOURCES.OnCancel () ;
begin
  Inherited ;
end ;
procedure TOF_RTSELECTRESSOURCES.bSuppClick(Sender: TObject);
begin
  if (G1.Row <> 0 ) and (G1.cells[1,G1.Row] <> '') then
     G1.DeleteRow (G1.Row);
end;

procedure TOF_RTSELECTRESSOURCES.RabRessources;
begin
G1.videpile (False);
end;

procedure TOF_RTSELECTRESSOURCES.GenereActionsRessources;
var  ListeResp,Select : string;
     i,nbact : integer;
     Q : TQuery;
     TobActGen,TobAction : TOB;
begin
  if(G1.Cells[1,1]='') then
  begin
   PGIInfo('Aucun élément sélectionné','');
   exit;
  end;
  if PGIAsk('Confirmez vous la génération des actions ?','')<>mrYes then exit ;

  TobActGen := LaTob;
  if TobActGen = Nil then
  begin
   PGIInfo('Informations de l'' action non trouvées','');
   exit;
  end;

  TobAction:=TOB.create ('ACTIONS',NIL,-1);
  TobAction.initValeurs (False);
  TobAction.PutValue ('RAC_NUMCHAINAGE',TobActGen.GetValue('RAG_NUMACTGEN'));

  TobAction.PutValue ('RAC_LIBELLE',TobActGen.GetValue('RAG_LIBELLE'));
  TobAction.PutValue ('RAC_TYPEACTION',TobActGen.GetValue('RAG_TYPEACTION'));
  TobAction.PutValue ('RAC_PRODUITPGI',TobActGen.GetValue('RAG_PRODUITPGI'));
  TobAction.PutValue ('RAC_TIERS',TobActGen.GetValue('RAG_INTERVENANT'));
  TobAction.PutValue ('RAC_AUXILIAIRE',TiersAuxiliaire (TobActGen.GetValue('RAG_INTERVENANT'), False));

  TobAction.PutValue ('RAC_DATEACTION',TobActGen.GetValue('RAG_DATEACTION'));
  TobAction.PutValue ('RAC_DATEECHEANCE',TobActGen.GetValue('RAG_DATEECHEANCE'));
  TobAction.PutValue ('RAC_ETATACTION',TobActGen.GetValue('RAG_ETATACTION'));
  TobAction.PutValue ('RAC_TABLELIBRE1',TobActGen.GetValue('RAG_TABLELIBRE1'));
  TobAction.PutValue ('RAC_TABLELIBRE2',TobActGen.GetValue('RAG_TABLELIBRE2'));
  TobAction.PutValue ('RAC_TABLELIBRE3',TobActGen.GetValue('RAG_TABLELIBRE3'));
  TobAction.PutValue ('RAC_TABLELIBREF1',TobActGen.GetValue('RAG_TABLELIBREF1'));
  TobAction.PutValue ('RAC_TABLELIBREF2',TobActGen.GetValue('RAG_TABLELIBREF2'));
  TobAction.PutValue ('RAC_TABLELIBREF3',TobActGen.GetValue('RAG_TABLELIBREF3'));
  TobAction.PutValue ('RAC_COUTACTION',TobActGen.GetValue('RAG_COUTACTION'));
  TobAction.PutValue ('RAC_BLOCNOTE',TobActGen.GetValue('RAG_BLOCNOTE'));
  TobAction.PutValue ('RAC_OPERATION',TobActGen.GetValue('RAG_OPERATION'));
  TobAction.PutValue ('RAC_AFFAIRE',TobActGen.GetValue('RAG_AFFAIRE'));
  TobAction.PutValue ('RAC_AFFAIRE0',TobActGen.GetValue('RAG_AFFAIRE0'));
  TobAction.PutValue ('RAC_AFFAIRE1',TobActGen.GetValue('RAG_AFFAIRE1'));
  TobAction.PutValue ('RAC_AFFAIRE2',TobActGen.GetValue('RAG_AFFAIRE2'));
  TobAction.PutValue ('RAC_AFFAIRE3',TobActGen.GetValue('RAG_AFFAIRE3'));
  TobAction.PutValue ('RAC_AVENANT',TobActGen.GetValue('RAG_AVENANT'));
  ListeResp:='';
  nbact:=0;
  Select := 'SELECT MAX(RAC_NUMACTION) FROM ACTIONS WHERE RAC_AUXILIAIRE = "'+ TobAction.GetValue ('RAC_AUXILIAIRE')+'"';
  Q:=OpenSQL(Select, True);
  if not Q.Eof then
     nbact := Q.Fields[0].AsInteger;
  Ferme(Q) ;

  for i:=1 to G1.RowCount-1 Do
  begin
    if G1.Cells[1,i]<>'' then
    begin
      Inc(nbact);
      TobAction.PutValue ('RAC_NUMACTION',nbact);
      TobAction.PutValue ('RAC_INTERVENANT',G1.Cells[1,i]);
      TobAction.PutValue ('RAC_INTERVINT',G1.Cells[1,i]+';');
      TobAction.putvalue ('RAC_DATECREATION', Date);
      TobAction.InsertDB(Nil);
      ListeResp:=ListeResp+ G1.Cells[1,i]+';';
    end;
  end;
  if (ListeResp <> '') then
    if length(ListeResp) < 86 then
      TFMul(Ecran).Retour:=ListeResp+'|'+TobAction.GetValue('RAC_LIBELLE')+'|'+ TobAction.GetValue('RAC_TYPEACTION')
    else
      TFMul(Ecran).Retour:=Copy(ListeResp,1,85)+'|'+TobAction.GetValue('RAC_LIBELLE')+'|'+ TobAction.GetValue('RAC_TYPEACTION');
  TobAction.free;
end;

procedure AGLDeplaceLesRessources( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul)  then TOTOF:=TFmul(F).LaTOF else exit;
if (TOTOF is TOF_RTSELECTRESSOURCES) then TOF_RTSELECTRESSOURCES(TOTOF).DeplaceLesRessources else exit;
end;

procedure AGLRabRessources( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul)  then TOTOF:=TFmul(F).LaTOF else exit;
if (TOTOF is TOF_RTSELECTRESSOURCES) then TOF_RTSELECTRESSOURCES(TOTOF).RabRessources else exit;
end;

procedure AGLGenereActionsRessources( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul)  then TOTOF:=TFmul(F).LaTOF else exit;
if (TOTOF is TOF_RTSELECTRESSOURCES) then TOF_RTSELECTRESSOURCES(TOTOF).GenereActionsRessources else exit;
end;

Initialization
  registerclasses ( [ TOF_RTSELECTRESSOURCES ] ) ;
  RegisterAglProc('DeplaceLesRessources', TRUE , 0, AGLDeplaceLesRessources);
  RegisterAglProc('RabRessources', TRUE , 0, AGLRabRessources);
  RegisterAglProc('GenereActionsRessources', TRUE , 0, AGLGenereActionsRessources);
end.

