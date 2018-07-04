{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 30/09/2013
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSELFACNEG_MUL ()
Mots clefs ... : TOF;BTSELFACNEG_MUL
*****************************************************************}
Unit BTSELFACNEG_MUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,
{$else}
     eMul, 
{$ENDIF}
		 HDB,	
     uTob, 
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     uEntCommun,
     UTofAfBaseCodeAffaire;

Type
  TOF_BTSELFACNEG_MUL = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;
  private
    TobPieces : TOB;
    procedure BLanceClick (Sender : TObject) ;
    procedure AjouteDoc (Cledoc : r_cledoc);
    function DemandeDatesFacturation(var DateFac: TDateTime; ModeGeneration: string): boolean;
  end ;

Implementation
uses HTB97,UtilTOBPiece,FactGrp,Aglinit, uTOFComm;

procedure TOF_BTSELFACNEG_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSELFACNEG_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSELFACNEG_MUL.OnUpdate ;
begin
  Inherited ;

end ;

procedure TOF_BTSELFACNEG_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTSELFACNEG_MUL.OnArgument (S : String ) ;
begin
  fMulDeTraitement := True;
  Inherited ;
  tobpieces := TOB.Create ('LES PIECES',nil,-1);
  TToolbarButton97 (GetControl('BOUVRIR')).onclick := BLanceClick;
end ;

procedure TOF_BTSELFACNEG_MUL.OnClose ;
begin
  tobpieces.Free;
  Inherited ;
end ;

procedure TOF_BTSELFACNEG_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSELFACNEG_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTSELFACNEG_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,
  Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
  inherited;
  Aff0 := THEdit(GetControl('AFFAIRE0'));
  Aff := THEdit(GetControl('GP_AFFAIRE'));
  Aff1 := THEdit(GetControl('GP_AFFAIRE1'));
  Aff2 := THEdit(GetControl('GP_AFFAIRE2'));
  Aff3 := THEdit(GetControl('GP_AFFAIRE3'));
  Aff4 := THEdit(GetControl('GP_AVENANT'));
end;

function TOF_BTSELFACNEG_MUL.DemandeDatesFacturation(var DateFac: TDateTime; ModeGeneration : string ) : boolean;
var TobDates : TOB;
begin
  TOBDates := TOB.Create ('LES DATES', nil,-1);
  TOBDates.AddChampSupValeur('RETOUROK','-');
  TOBDates.AddChampSupValeur('DATFAC',V_PGI.DateEntree);
  TOBDates.AddChampSupValeur('DATESITUATION','-');
  TOBDates.AddChampSupValeur('TYPEDATE','Date de facturation');
  TRY
    TheTOB := TOBDates;
    AGLLanceFiche('BTP','BTDEMANDEDATES','','','');
    TheTOB := nil;
    if TOBDates.getValue('RETOUROK')='X' then
    begin
    	DateFac := TOBDates.GetDateTime('DATFAC');
    end;
  FINALLY
  	result := (TOBDates.getValue('RETOUROK')='X');
  	freeAndNil(TOBDates);
  END;
end;


procedure TOF_BTSELFACNEG_MUL.BLanceClick(Sender: TObject);
var Cledoc : r_cledoc;
		Q : TQuery;
    Fliste : THDbGrid;
    i : Integer;
    DateFac : TDateTime;
begin
  Fliste := THDbGrid(GetCOntrol('Fliste'));
  Q := TFMul(ecran).Q;
  if (PGIAsk ('Confirmez-vous la facturation des éléments sélectionnés', Ecran.Caption)<>mrYes) then exit;
  DateFac := iDate1900;
  if not DemandeDatesFacturation (DateFac,'') then Exit;
  SourisSablier;
  TRY
    TobPieces.ClearDetail;
    if Fliste.AllSelected then
    BEGIN
      Q.First;
      while Not Q.EOF do
      BEGIN
        Cledoc.NaturePiece :=Q.FindField('GL_NATUREPIECEG').AsString;
        Cledoc.Souche :=Q.FindField('GL_SOUCHE').AsString;
        Cledoc.NumeroPiece :=Q.FindField('GL_NUMERO').AsInteger;
        Cledoc.Indice:=Q.FindField('GL_INDICEG').AsInteger;
        AjouteDoc (cledoc);
        Q.NEXT;
      END;
      Fliste.AllSelected:=False;
    END
    else
    begin
      for i:=0 to Fliste.nbSelected-1 do
      begin
        Fliste.GotoLeBookmark(i);
        Cledoc.NaturePiece :=Fliste.datasource.dataset.FindField('GL_NATUREPIECEG').AsString;
        Cledoc.Souche :=Fliste.datasource.dataset.FindField('GL_SOUCHE').AsString;
        Cledoc.NumeroPiece :=Fliste.datasource.dataset.FindField('GL_NUMERO').AsInteger;
        Cledoc.Indice :=Fliste.datasource.dataset.FindField('GL_INDICEG').AsInteger;
        //
        AjouteDoc (cledoc);
      end;
    end;
    if TobPieces.Detail.Count = 0 then Exit;
  	RegroupeLesPieces(TobPieces, 'FBC', false, False, True,0, DateFac,true,false,false,true,false,'',false);
  FINALLY
    SourisNormale ;
    TobPieces.ClearDetail;
    RefreshDB(); 
  END;
end;

procedure TOF_BTSELFACNEG_MUL.AjouteDoc(Cledoc: r_cledoc);
var TOBP : TOB;
		QQ : TQuery;
begin
  QQ := OpenSQL('SELECT * FROM PIECE WHERE '+WherePiece(cledoc,ttdPiece,true),True,1,'',True);
  if not QQ.eof then
  begin
		TOBP := TOB.Create('PIECE',TobPieces,-1);
    TOBP.SelectDB('',QQ);
  end;
  ferme (QQ);
end;

Initialization
  registerclasses ( [ TOF_BTSELFACNEG_MUL ] ) ;
end.
