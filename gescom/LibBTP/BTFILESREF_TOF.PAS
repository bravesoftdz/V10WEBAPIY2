{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 19/04/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTFILESREF ()
Mots clefs ... : TOF;BTFILESREF
*****************************************************************}
Unit BTFILESREF_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     uDbxDataSet,
     mul,
{$else}
     eMul,
{$ENDIF}
     fe_main,
     AGLInit,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     ParamSoc, 
     UTOF ; 

Type
  TOF_BTFILESREF = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    EmplacementStockage : string;
    procedure BInsertClick (Sender : TObject);
    procedure GSDblClick (Sender : TObject);
    procedure BVoirDocClick (Sender : TObject);
  end ;

Implementation
uses UtilFichiers;

procedure TOF_BTFILESREF.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTFILESREF.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTFILESREF.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTFILESREF.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTFILESREF.OnArgument (S : String ) ;
begin
  Inherited ;
  EmplacementStockage := GetParamSocSecur ('SO_BTEMPLFILEREF','');
  TToolbarButton97 (GetControl('BInsert')).onclick := BInsertClick;
  TToolbarButton97 (GetControl('BVOIR')).onclick := BVoirDocClick;
  THGrid(GetControl('fliste')).OnDblClick := GSDblClick;
  EmplacementStockage := GetParamSocSecur ('SO_BTEMPLFILEREF','');
end ;

procedure TOF_BTFILESREF.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTFILESREF.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTFILESREF.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTFILESREF.BInsertClick(Sender: TObject);
var TT : TOB;
    ONETT : TOB;
begin
  TT := TOB.Create ('UNE TOB',nil,-1);
  ONETT := TOB.Create('BFILES',nil,-1);
  TRY
    TT.AddChampSupValeur('CODE','REFERENCE');
    TT.AddChampSupValeur('FILENAME','');
    TT.AddChampSupValeur('FILESTOCK','');
    TT.AddChampSupValeur('FILEDESC','');
    TT.AddChampSupValeur('OKOK','-');
    TheTOB := TT;
    AglLanceFiche('BTP','BTNEWFILESTD','','','ACTION=MODIFICATION');
    TheTOB := nil;
    if TT.GetString('OKOK')='X' then
    begin
      ONETT.SetString('BF0_CODE','REFERENCE');
      ONETT.SetString('BF0_FILENAME',TT.GetString('FILENAME'));
      ONETT.SetString('BF0_LIBELLE',TT.GetString('FILEDESC'));
      ONETT.SetString('BF0_CRITSTOCKFILE',TT.GetString('FILESTOCK'));
      ONETT.SetAllModifie(true);
      ONETT.InsertDB(nil);
      TToolbarButton97(GetControl('Bcherche')).Click;
    end;
  FINALLY
    TT.free;
    ONETT.free;
  END;
end;

procedure TOF_BTFILESREF.GSDblClick(Sender: TObject);
var FileUid,Desc,Crit : string;
    TT : TOB;
begin
  FileUid := TFMul(Ecran).Fliste.datasource.dataset.FindField('BF0_FILENAME').AsString;
  Desc := TFMul(Ecran).Fliste.datasource.dataset.FindField('BF0_LIBELLE').AsString;
  Crit := TFMul(Ecran).Fliste.datasource.dataset.FindField('BF0_CRITSTOCKFILE').AsString;
  TT := TOB.Create('BFILES',LaTOB,-1);
  TT.SetString('BF0_CODE',LaTOB.GetString('CODE'));
  TT.SetString('BF0_FILENAME',FileUid);
  TT.SetString('BF0_CRITSTOCKFILE',Crit);
  TT.SetString('BF0_LIBELLE',Desc);
  LaTOB.SetString('MODIF','X');
  Ecran.close;
end;

procedure TOF_BTFILESREF.BVoirDocClick(Sender: TObject);
var FileName : string;
begin
  FileName := IncludeTrailingBackslash (EmplacementStockage)+TFMul(Ecran).Fliste.datasource.dataset.FindField('BF0_FILENAME').AsString;
  OuvreDocument (FileName);
end;

Initialization
  registerclasses ( [ TOF_BTFILESREF ] ) ; 
end.
