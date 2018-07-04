{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 31/10/2017
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTGENFACBAST ()
Mots clefs ... : TOF;BTGENFACBAST
*****************************************************************}
Unit BTGENFACBAST_TOF ;

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
     HTB97,
     Urapport,
     galPatience,
     UTOF ; 

Type
  TResultB = class(TObject)
  private
    fResult : Integer;
    fNumBast : String;
    fNaturePiece : string;
    fNumeroPiece : Integer;
    fNumSituation : Integer;
    fMessage : string;
  public
    property result : integer read fResult write fResult;
    property MessageRet : string read fMessage write fMessage;
    property NaturePiece : string read fNaturePiece write fNaturePiece;
    property Numero : integer read fNumeroPiece write fNumeroPiece;
    property Numbast : string read fNumBast write fNumBast;
    property NumSituation : Integer read fNumSituation write fNumSituation;
    procedure Clear;
  end;

  TOF_BTGENFACBAST = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Rapport : TFBTRapport;
    procedure BVAliderClick (Sender : TObject);
    procedure GenereFacture(CodeAffaire,CodeMarche,Fournisseur : string; NumSituation : integer ; PaimentPoc:  Boolean; DateFac : TDateTime; Result : TResultB);
    procedure AddResultToRapport (Result : TResultB);
  end ;

Implementation
uses Paramsoc,CalcOLEGenericBTP,UGenereDocument,FactSituations;

procedure TOF_BTGENFACBAST.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTGENFACBAST.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTGENFACBAST.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTGENFACBAST.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTGENFACBAST.OnArgument (S : String ) ;
begin
  Inherited ;
  TToolbarButton97 (GetControl('BOuvrir')).OnClick := BVAliderClick;
  Rapport := TFBTRapport.create(application.mainform);
  Rapport.MemoRapport.lines.clear;

end ;

procedure TOF_BTGENFACBAST.OnClose ;
begin
  Rapport.free;
  Inherited ;
end ;

procedure TOF_BTGENFACBAST.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTGENFACBAST.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTGENFACBAST.BVAliderClick(Sender: TObject);
var Q : TQuery;
    CodeAffaire,CodeMarche,Fournisseur : string;
    NumSituation : Integer;
    PaiementPoc : Boolean;
    i : Integer;
    Result : TResultB;
    XX : TFPatience;
    DateFac : TDateTime;
begin
  //
  if Not DemandeDatesFacturation(DateFac) then exit;
  //
  Rapport.MemoRapport.lines.Clear;
  XX := FenetrePatience('Gestion des BAST',aoMilieu, False,true);
  XX.lAide.Caption := 'Validation comptable des BAST ...';
  XX.StartK2000 ;
  XX.Refresh;
  //
  Result := TResultB.Create;
  TRY
    Q:=TFMul(Ecran).Q;
    if TFMul(Ecran).Fliste.AllSelected then
    BEGIN
      Q.First;
      while Not Q.EOF do
      BEGIN
        CodeAffaire:=Q.FindField('BM4_AFFAIRE').AsString;
        CodeMarche:=Q.FindField('BM4_CODEMARCHE').AsString;
        Fournisseur:=Q.FindField('BM4_FOURNISSEUR').AsString;
        NumSituation:=Q.FindField('BM4_NUMSITUATION').AsInteger;
        PaiementPoc := (Q.FindField('BM4_PAIEMENTPOC').AsString ='X');
        //
        Result.Clear;
        GenereFacture(CodeAffaire,CodeMarche,Fournisseur,NumSituation,PaiementPoc,DateFac,Result);
        AddResultToRapport (Result);
        Q.NEXT;
      END;
      TFMul(Ecran).Fliste.AllSelected:=False;
    END else
    begin
      for i:=0 to TFMul(Ecran).Fliste.nbSelected-1 do
      begin
        TFMul(Ecran).Fliste.GotoLeBookmark(i);
        CodeAffaire:=TFMul(Ecran).Fliste.datasource.dataset.FindField('BM4_AFFAIRE').AsString;
        CodeMarche:=TFMul(Ecran).Fliste.datasource.dataset.FindField('BM4_CODEMARCHE').AsString;
        Fournisseur:=TFMul(Ecran).Fliste.datasource.dataset.FindField('BM4_FOURNISSEUR').AsString;
        NumSituation:=TFMul(Ecran).Fliste.datasource.dataset.FindField('BM4_NUMSITUATION').AsInteger;
        PaiementPoc:=(TFMul(Ecran).Fliste.datasource.dataset.FindField('BM4_PAIEMENTPOC').AsString='X');
        //
        Result.Clear;
        GenereFacture(CodeAffaire,CodeMarche,Fournisseur,NumSituation,PaiementPoc,DateFac,Result);
        AddResultToRapport (Result);
      end;
    end;
  FINALLY
    XX.StopK2000 ;
    XX.Free;
    Rapport.ShowModal;
    Result.free;
    TToolbarButton97 (GetControl('Bcherche')).Click;  // refresh
  END;
end;

{ TResultB }

procedure TResultB.Clear;
begin
  fResult := 0;
  fMessage := '';
end;

procedure TOF_BTGENFACBAST.GenereFacture(CodeAffaire,CodeMarche,Fournisseur: string; NumSituation: integer; PaimentPoc: Boolean; DateFac : TDateTime; Result: TResultB);
var TT : TOB;
    QQ : TQuery;
    SQL : String;
    XX : TGenerePiece;
    TheResult : Tresult;
begin
  Result.result := -1;
  Result.MessageRet := 'Abandonné par l''utilisateur';
  Result.fNumBast := '';
  Result.fNumSituation := 0;
  //

  TT := TOB.Create('BASTENT',nil,-1);
  TRY
    SQL :='SELECT * FROM BASTENT '+
                   'LEFT JOIN BTFACTST ON BM4_AFFAIRE=BM3_AFFAIRE AND BM4_FOURNISSEUR=BM3_FOURNISSEUR AND BM4_CODEMARCHE=BM3_CODEMARCHE AND BM4_NUMSITUATION=BM3_NUMSITUATION '+
//                   'LEFT JOIN BTMARCHEST ON BM4_AFFAIRE=BM0_AFFAIRE AND BM4_FOURNISSEUR=BM0_FOURNISSEUR AND SUBSTRING(BM4_CODEMARCHE,1,8)=BM0_MARCHE '+
                   'WHERE '+
                   'BM4_AFFAIRE="'+CodeAffaire+'" AND '+
                   'BM4_FOURNISSEUR="'+FOurnisseur+'" AND '+
                   'BM4_CODEMARCHE="'+CodeMarche+'" AND '+
                   'BM4_PAIEMENTPOC="'+BoolToStr(PaimentPoc)+'" AND '+
                   'BM4_NUMSITUATION='+InttoStr(NumSituation);
    QQ := OpenSQL (SQL,True,1,'',true);
    if not QQ.Eof then TT.SelectDB('',QQ);
    ferme (QQ);
    //
    Result.fNumBast := TT.GetString('BM3_NUMERODOC');
    Result.fNumSituation := TT.GetInteger('BM3_NUMSITUATION');
    //
    Result.result := -1;
    Result.MessageRet := 'Document innexistant';
    if TT.GetString('BM4_AFFAIRE') = '' then Exit;
    SQL := 'SELECT * FROM BASTLIG '+
           'WHERE '+
           'BM5_AFFAIRE="'+CodeAffaire+'" AND '+
           'BM5_FOURNISSEUR="'+FOurnisseur+'" AND '+
           'BM5_CODEMARCHE="'+CodeMarche+'" AND '+
           'BM5_PAIEMENTPOC="'+BoolToStr(PaimentPoc)+'" AND '+
           'BM5_NUMSITUATION='+InttoStr(NumSituation)+' AND '+
           'BM5_TYPELIG="LIG" ORDER BY BM5_TYPELBAST,BM5_NUMLIGNE';
    Result.MessageRet := 'Détail du BAST innexistant';
    TT.LoadDetailDBFromSQL('BASTLIG',SQL,false);
    if TT.detail.count = 0 then exit;
    XX := TGenerePiece.create;
    Result.MessageRet := 'Erreur en génération de document';
    TRY
      TheResult := XX.GenereDocFromBAST(TT,DateFac);
      //
      if TheResult.ErrorResult=OeOk then
      begin
        Result.result := 0;
        Result.NaturePiece := TheResult.NatureDoc;
        Result.Numero  := TheResult.NumeroDoc;
        TT.SetString('BM4_REFGESCOM',TheResult.RefPiece);
        TT.UpdateDB;
      end else
      begin
        Result.MessageRet := TheResult.LibError;
      end;
    FINALLY
      XX.free;
    END;
  FINALLY
    TT.Free;
  END;
end;

procedure TOF_BTGENFACBAST.AddResultToRapport(Result: TResultB);
var St : string;
begin
  if Result.result = 0 then
  begin
    St := Format('BAST N° %s situation N° %d --> %s N° %d',[Result.Numbast,Result.NumSituation,RechDom('GCNATUREPIECEG','FF',false),Result.fNumeroPiece]);
  end else
  begin
    St := Format('ERREUR en génération du BAST N° %s situation N° %d --> %s',[Result.Numbast,Result.NumSituation,Result.MessageRet]);
  end;
  Rapport.MemoRapport.lines.Add(St);
end;


Initialization
  registerclasses ( [ TOF_BTGENFACBAST ] ) ;
end.
