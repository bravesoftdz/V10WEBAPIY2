{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 07/07/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PARFOU_DUPLIC ()
Mots clefs ... : TOF;PARFOU_DUPLIC
*****************************************************************}
Unit TofParfouDuplic ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
      MaineAgl,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOB,
     AglInit,
     UTOF ;

Type
  TOF_PARFOU_DUPLIC = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    function ControleSaisieOk: boolean;
  end ;

procedure DuplicRecupTarifFour (CodeImport : string);

Implementation

procedure MiseAjour (TobParFou: TOB);
var Indice : integer;
begin
  V_PGI.IoError:=oeOk;
  TOBParFou.SetAllModifie (true);
  for indice := 0 to TOBParFou.detail.count -1 do TOBParFou.detail[indice].SetAllModifie (true);
  BEGINTRANS;
  TRY
    TOBPArFOu.InsertDB (nil,true);
  EXCEPT
    V_PGI.IoError:=oeUnknown;
  END;
  if V_PGI.IoError=oeOK then CommitTrans else Rollback;
end;

procedure DupliquePARFOU (CodeImport : string ; LaTob : TOB);
var TOBPARFOU,TOBPARFOULIG : TOB;
    QQ : TQuery;
    Indice : integer;
begin
  TOBPARFOU := TOB.create ('PARFOU',nil,-1);
  TRY
    // chargement de PARFOU
    QQ := OpenSql ('SELECT * FROM PARFOU WHERE GRF_PARFOU="'+CodeImport+'"',true,-1,'',true);
    if not QQ.eof then
    begin
      TOBPARFOU.SelectDB ('',QQ);
    end;
    ferme (QQ);
    // chargement de PARFOULIG
    QQ := OpenSql ('SELECT * FROM PARFOULIG WHERE GFL_PARFOU="'+CodeImport+'"',true,-1,'',true);
    if not QQ.eof then
    begin
      TOBPARFOU.LoadDetailDB ('PARFOULIG','','',QQ,false,true);
    end;
    ferme (QQ);
    // Mise en place des nouvelles valeurs
    TOBPARFOU.PutValue('GRF_PARFOU',laTOB.GetValue('GRF_PARFOU'));
    TOBPARFOU.PutValue('GRF_TIERS',laTOB.GetValue('GRF_TIERS'));
    //
    for Indice := 0 to TOBPARFOU.detail.count -1 do
    begin
      TOBPARFOU.detail[indice].PutValue('GFL_PARFOU',laTOB.GetValue('GRF_PARFOU'));
    end;
    MiseAjour (TobParFou);
  FINALLY
    TOBPARFOU.free;
  END;
end;

procedure DuplicRecupTarifFour (CodeImport : string);
var MyTOB : TOB;
    result : string;
begin
  MyTOb := TOB.Create ('The_PARAMETRES',nil,-1);
  TRY
    MyTOB.addChampSupValeur ('GRF_PARFOU','',false);
    MyTOB.addChampSupValeur ('GRF_TIERS','',false);
    MyTOB.addChampSupValeur ('RETOUR','0',false);
    TheTOb := MyTob;
    AglLanceFiche('GC', 'GCPARFOU_DUPLIC', '', '', 'ACTION=MODIFICATION');
    if TheTOb <> nil then
    begin
      if TheTOB.GetValue ('RETOUR') = 1 then
      DupliquePARFOU (CodeImport,TheTob);
    end;
  FINALLY
    MyTob.free;
    TheTob := nil;
  END;
end;

procedure TOF_PARFOU_DUPLIC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PARFOU_DUPLIC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PARFOU_DUPLIC.OnUpdate ;
begin
  Inherited ;
  NextPrevControl (TForm(Ecran),true);
  if ControleSaisieOk then
  begin
    LaTOB.putValue('RETOUR',1);
    TheTOB := LaTob;
  end else
  begin
    TForm(Ecran).ModalResult:=0;
  end;
end ;

procedure TOF_PARFOU_DUPLIC.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PARFOU_DUPLIC.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_PARFOU_DUPLIC.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PARFOU_DUPLIC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PARFOU_DUPLIC.OnCancel () ;
begin
  Inherited ;
end ;

function TOF_PARFOU_DUPLIC.ControleSaisieOk : boolean;
begin
  result := true;
  if THedit(GetControl('GRF_PARFOU')).Text = '' then
  begin
    PGIBox (TraduireMemoire('Vous devez renseigner un code de paramétrage'));
    THedit(GetControl('GRF_PARFOU')).SetFocus;
    result := false;
    exit;
  end;
  if THedit(GetControl('GRF_PARFOU')).Text <> '' then
  begin
    if ExisteSQL ('SELECT GRF_PARFOU FROM PARFOU WHERE GRF_PARFOU="'+THedit(GetControl('GRF_PARFOU')).Text+'"') then
    begin
    PGIBox (TraduireMemoire('Ce code existe déjà. Veuillez en saisir en autre'));
    THedit(GetControl('GRF_PARFOU')).SetFocus;
    result := false;
    exit;
    end;
  end;
  if THedit(GetControl('GRF_TIERS')).Text = '' then
  begin
    PGIBox (TraduireMemoire('Le code du fournisseur est obligatoire. Veuillez le saisir'));
    THedit(GetControl('GRF_PARFOU')).SetFocus;
    result := false;
    exit;
  end;
  if THedit(GetControl('GRF_TIERS')).Text <> '' then
  begin
    if not ExisteSQL ('SELECT T_TIERS FROM TIERS WHERE T_NATUREAUXI="FOU" AND T_TIERS="'+THedit(GetControl('GRF_TIERS')).Text+'"') then
    begin
    PGIBox (TraduireMemoire('Ce fournisseur n''existe pas. Veuillez en saisir en autre'));
    THedit(GetControl('GRF_TIERS')).SetFocus;
    result := false;
    exit;
    end;
  end;
end;

Initialization
  registerclasses ( [ TOF_PARFOU_DUPLIC ] ) ;
end.

