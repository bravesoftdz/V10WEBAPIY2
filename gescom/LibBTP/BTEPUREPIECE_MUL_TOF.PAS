{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 02/10/2009
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTEPUREPIECE_MUL ()
Mots clefs ... : TOF;BTEPUREPIECE_MUL
*****************************************************************}
Unit BTEPUREPIECE_MUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Ent1,
     EntGc,
     UEntCommun,
     AglInit,
     M3FP,
{$IFNDEF EAGLCLIENT}
		 fe_main,
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
		 MainEagl,
     eMul,Hqry,
{$ENDIF}
     uTob,
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox,
     RECALCPIECE_RAP_TOF,
     UTOF ;

Type
  TOF_BTEPUREPIECE_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	TOBPieces : TOB;
    SansControle : Boolean;
    procedure EpureCesPieces;
    procedure AddPieceAtraiter (Cledoc : r_cledoc);
    procedure Controlechamp(Champ, Valeur: string);
  end ;

Implementation

procedure TOF_BTEPUREPIECE_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTEPUREPIECE_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTEPUREPIECE_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTEPUREPIECE_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTEPUREPIECE_MUL.Controlechamp(Champ, Valeur : string);
begin
  if Champ = 'SUPFACTURE' then
  begin
    THEdit(GetCOntrol('XX_WHERE')).Text := ' AND GP_NATUREPIECEG IN ("FBP","FBT","DAC")';
    THvalComboBox(GetControl('GP_NATUREPIECEG')).Plus := ' AND GPP_NATUREPIECEG IN ("FBP","FBT","DAC")';
    ecran.caption := TraduireMemoire('Suppression des factures sans contrôles');
    UpdateCaption(Ecran);
  end;
  if Champ = 'SANSCONTROLE' then
  begin
    SansControle := True;
  end;
end;


procedure TOF_BTEPUREPIECE_MUL.OnArgument (S : String ) ;
var CC : THValComboBox;
    critere,Champ,Valeur : string;
    i :Integer;
begin
  Inherited ;
  TOBPieces := TOB.Create ('LES PIECESS',nil,-1);
  THEdit(GetCOntrol('XX_WHERE')).Text := ' AND GP_AFFAIRE=""';
  CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ;if CC<>Nil then PositionneEtabUser(CC) ;
  Critere := S;

  While (Critere <> '') do
  BEGIN
    i:=pos(':',Critere);
    if i = 0 then i:=pos('=',Critere);
    if i <> 0 then
       begin
       Champ:=copy(Critere,1,i-1);
       Valeur:=Copy (Critere,i+1,length(Critere)-i);
       end
    else
       Champ := Critere;
    Controlechamp(Champ, Valeur);
    Critere:=(Trim(ReadTokenSt(S)));
  END;

end ;

procedure TOF_BTEPUREPIECE_MUL.OnClose ;
begin
  TOBPieces.free;
  Inherited ;
end ;

procedure TOF_BTEPUREPIECE_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTEPUREPIECE_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTEPUREPIECE_MUL.EpureCesPieces;
var
{$IFDEF EAGLCLIENT}
		Q : THquery;
{$ELSE}
		Q : Tquery;
{$ENDIF}
		F : TFmul;
    i : integer;
    Cledoc : r_cledoc;
    SRet : String;
    TOBPass : TOB;
    GetOUt : boolean;
begin
  //
  TOBPieces.clearDetail;
  GetOut := false;
  TOBPass := TOB.create ('UNE TOB',nil,-1);
  TOBPASS.AddChampSupValeur ('RESULTAT','');
  TheTOB := TOBPAss;
  AGLLAnceFiche('BTP','BTCONFIRMEPURPIEC','','','');
  TheTOB := nil;
  if TOBPass.getValue('RESULTAT')<>'OK' then GetOUt := true;
  TOBPass.free;
  //
  if GetOUt then exit;   // pas d'accord
  //
  TOBPass := TOB.create ('UNE TOB',nil,-1);
  TOBPASS.AddChampSupValeur ('DAYPASS','');
  TOBPASS.AddChampSupValeur ('RESULTAT','');
  TheTOB := TOBPAss;
 	Sret := AGLLAnceFiche('BTP','BTCONFIRMPASS','','','');
  TheTOB := nil;
  if TOBPass.GetValue('RESULTAT')<> 'OK' then GetOut := true;
  TOBPass.free;
  //
  if GetOUt then exit; // toujours pas d'accord
  //
  F := TFmul(ecran);
  TRY
    if TFMul(F).Fliste.AllSelected then
    BEGIN
      Q:=TFmul(F).Q;
      Q.First;
      while Not Q.EOF do
      BEGIN
        FillChar(CleDoc, Sizeof(CleDoc), #0);
        cledoc.NaturePiece := Q.FindField('GP_NATUREPIECEG').AsString;
        Cledoc.Souche      := Q.FindField('GP_SOUCHE').AsString;
        cledoc.NumeroPiece := Q.FindField('GP_NUMERO').AsInteger;
        cledoc.DatePiece   := Q.FindField('GP_DATEPIECE').AsDateTime;
        cledoc.Indice      := Q.FindField('GP_INDICEG').AsInteger;
        AddPieceAtraiter (Cledoc);
        Q.NEXT;
      END;
      if TOBPieces.detail.count > 0 then TraiteEpurePieces (TOBPieces);
      TFMul(F).Fliste.AllSelected:=False;
    END ELSE
    BEGIN
      for i:=0 to TFMul(F).Fliste.nbSelected-1 do
      begin
        TFMul(F).Fliste.GotoLeBookmark(i);
        FillChar(CleDoc, Sizeof(CleDoc), #0);
{$IFDEF EAGLCLIENT}
      	TFMul(F).Q.TQ.Seek(TFMul(F).FListe.Row-1) ;

        cledoc.NaturePiece :=TFMul(F).Q.FindField('GP_NATUREPIECEG').AsString;
        Cledoc.Souche      :=TFMul(F).Q.FindField('GP_SOUCHE').AsString;
        cledoc.NumeroPiece :=TFMul(F).Q.FindField('GP_NUMERO').AsInteger;
        cledoc.DatePiece   :=TFMul(F).Q.FindField('GP_DATEPIECE').AsDateTime;
        cledoc.Indice      :=TFMul(F).Q.FindField('GP_INDICEG').AsInteger;
        AddPieceAtraiter (Cledoc);
{$ELSE}
        cledoc.NaturePiece :=TFMul(F).Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
        Cledoc.Souche      :=TFMul(F).Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
        cledoc.NumeroPiece :=TFMul(F).Fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger;
        cledoc.DatePiece   :=TFMul(F).Fliste.datasource.dataset.FindField('GP_DATEPIECE').AsDateTime;
        cledoc.Indice      :=TFMul(F).Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger;
        AddPieceAtraiter (Cledoc);
{$ENDIF}
      end;
    	if TOBPieces.detail.count > 0 then TraiteEpurePieces (TOBPieces);
    END;
  finally
  	TOBPieces.ClearDetail;
  end;
end;

procedure AglBTEpurePieces (parms:array of variant; nb: integer ) ;
var  F : TForm ;
     LaTof : TOF;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFMul) then Latof:=TFMul(F).Latof else exit;
  if (LaTof is TOF_BTEPUREPIECE_MUL) then TOF_BTEPUREPIECE_MUL(LaTof).EpureCesPieces else exit;
end;

procedure TOF_BTEPUREPIECE_MUL.AddPieceAtraiter(Cledoc: r_cledoc);
var TOBPiece : TOB;
begin
	TOBPIece := TOB.Create ('PIECE',TOBPieces,-1);
	TOBPIece.AddChampSupValeur('AFF_GENERAUTO','DIR');
  TOBPiece.putValue('GP_NATUREPIECEG',cledoc.NaturePiece);
  TOBPiece.putValue('GP_SOUCHE',cledoc.Souche);
  TOBPiece.putValue('GP_NUMERO',cledoc.NumeroPiece);
  TOBPiece.putValue('GP_INDICEG',cledoc.Indice);
end;

Initialization
  registerclasses ( [ TOF_BTEPUREPIECE_MUL ] ) ;
	RegisterAglProc('BTEpureCesPieces', True , 0, AglBTEpurePieces);
end.
