{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 22/11/2017
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSELMOISAVANC ()
Mots clefs ... : TOF;BTSELMOISAVANC
*****************************************************************}
Unit BTSELMOISAVANC_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     uDbxDataSet,
     mul, fe_Main,M3Fp,
{$else}
     eMul,
{$ENDIF}
     AglInit,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     uEntCommun,
     UTOF ;

Type
  TOF_BTSELMOISAVANC = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    cledoc : r_cledoc;
    Affaire : string;
    procedure BValideClick (Sender : TObject);
    procedure BInsertClick (Sender : Tobject);
  end ;

procedure SelectionMoisAvancPoc (Parms : array of variant ; nb : integer);

Implementation

procedure NewMois (cledoc : r_cledoc; Affaire : string);
var TOBP : TOB;
    Params : string;
begin
  TOBP := TOB.Create('LES PARAMS',nil,-1);
  TRY
    TOBP.AddChampSupValeur('NATUREPIECEG',cledoc.NaturePiece);
    TOBP.AddChampSupValeur('SOUCHE',cledoc.Souche);
    TOBP.AddChampSupValeur('NUMERO',cledoc.NumeroPiece);
    TOBP.AddChampSupValeur('INDICEG',cledoc.Indice);
    TOBP.AddChampSupValeur('MOIS','');
    TheTOB := TOBP;
    AGLLanceFiche('BTP','BTSAISMOIAVANC','','','ACTION=MODIFICATION');
    TheTOB := nil;
    if TOBP.GetString('MOIS')<>'' then
    begin
      Params := 'BNATURE='+cledoc.NaturePiece+';BSOUCHE='+cledoc.Souche+';BNUMERO='+InttoStr(cledoc.NumeroPiece)+
                ';BINDICEG='+IntToStr(cledoc.Indice)+';BMOIS='+TOBP.GetString('MOIS')+';BAFFAIRE='+Affaire;
      AGLLanceFiche('BTP','BSAISIEAVANCPOC','','','ACTION=MODIFICATION;'+Params);
    end;
  FINALLY
    TOBP.Free;
  END;
end;

procedure SelectionMoisAvancPoc (Parms : array of variant ; nb : integer);
var SQL,Params,Affaire : string;
    cledoc : R_CLEDOC;
begin
  Affaire := Parms[5];
  cledoc.NaturePiece := Parms[1];
  cledoc.Souche := Parms[2];
  cledoc.NumeroPiece := Parms[3];
  cledoc.Indice := Parms[4];
  Sql := 'SELECT 1 FROM BAVANCTECH WHERE '+
         'BM9_NATUREPIECEG="'+cledoc.NaturePiece+'" AND '+
         'BM9_SOUCHE="'+cledoc.Souche+'" AND '+
         'BM9_NUMERO='+IntTostr(cledoc.NumeroPiece)+' AND '+
         'BM9_INDICEG='+IntToStr(cledoc.Indice);
  if ExisteSQL(SQL) then
  begin
    Params :='BM9_NATUREPIECEG='+cledoc.NaturePiece+';BM9_SOUCHE='+cledoc.Souche+';BM9_NUMERO='+InttoStr(cledoc.NumeroPiece)+';BM9_INDICEG='+InttoStr(cledoc.Indice)+';BM9_AFFAIRE='+Affaire;
    AGLLanceFiche('BTP','BTSELMOISAVANC',Params,'','ACTION=MODIFICATION');
  end else
  begin
    NewMois (cledoc,Affaire);
  end;
end;

procedure TOF_BTSELMOISAVANC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSELMOISAVANC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSELMOISAVANC.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTSELMOISAVANC.OnLoad ;
begin
  Inherited ;
  cledoc.NaturePiece := GetControlText('BM9_NATUREPIECEG');
  cledoc.Souche := GetControlText('BM9_SOUCHE');
  cledoc.NumeroPiece := StrToInt(GetControlText('BM9_NUMERO'));
  cledoc.Indice := StrToInt(GetControlText('BM9_INDICEG'));
  Affaire := GetControlText('BM9_AFFAIRE');
end ;

procedure TOF_BTSELMOISAVANC.OnArgument (S : String ) ;
begin
  Inherited ;
  //
  TToolbarButton97 (GetControl('BVALIDE')).OnClick := BValideClick;
  TToolbarButton97 (GetControl('BINSERT')).OnClick := BInsertClick;
end ;

procedure TOF_BTSELMOISAVANC.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTSELMOISAVANC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSELMOISAVANC.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTSELMOISAVANC.BValideClick(Sender: TObject);
begin

end;

procedure TOF_BTSELMOISAVANC.BInsertClick(Sender: Tobject);
begin
  NewMois(cledoc,Affaire);
  TFMul(ecran).ChercheClick;
end;

Initialization
  RegisterAglProc( 'SelectionMoisAvancPoc',True,5,SelectionMoisAvancPoc);
  registerclasses ( [ TOF_BTSELMOISAVANC ] ) ;
end.
