{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 22/11/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RECUPSUSPECT ()
Mots clefs ... : TOF;RECUPSUSPECT
*****************************************************************}
Unit UtofRECUPSUSPECT ;

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
     MainEAGL,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     UtofRTDuplImport,
     RecupSuspect,
     Paramsoc,
     UTOF ;

Type
  TOF_RECUPSUSPECT = Class (TOF)
    public

    BDuplique           : TToolbarButton97;
    BImportParTiers     : TToolbarButton97;           //import par code tiers     //TJA 02/07/2007
    Fliste              : THGrid;
    BInsert             : TToolbarButton97;
    BDelete             : TToolbarButton97;
    BSauve              : TToolbarButton97;   // FQ 10848
    BRestau             : TToolbarButton97;   // FQ 10848

    procedure OnArgument (S : String ) ; override ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    procedure BDupliqueOnCLick (Sender : Tobject);
    procedure FlisteOnDblClick (Sender : Tobject);
    procedure BInsertOnClick (Sender : Tobject);
    procedure BDeleteOnClick (Sender : Tobject);
    procedure BSauveOnCLick (Sender : Tobject);   // FQ 10848
    procedure BRestauOnCLick (Sender : Tobject);  // FQ 10848

  end ;

Function RTLanceFiche_RECUPSUSPECT(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

{$IFDEF VER150}
  uses
    Variants
    ;
{$ENDIF VER150}

Function RTLanceFiche_RECUPSUSPECT(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_RECUPSUSPECT.OnArgument (S : String ) ;
begin
  Inherited ;

  Fliste                := THGrid(GetControl('FLISTE'));
  Fliste.OnDblClick     := FlisteOnDblClick;

  BInsert               := TToolbarButton97(GetControl('BINSERT'));
  BInsert.OnClick       := BInsertOnClick;

  BDelete               := TToolbarButton97(GetControl('BDELETE'));
  BDelete.OnClick       := BDeleteOnClick;

  BDuplique             := TToolbarButton97(GetControl('BDUPLIQUE'));
  BDuplique.OnClick     := BDupliqueOnCLick;

  BImportParTiers       := TToolbarButton97(GetControl('BIMPORTPARTIERS'));
  if Assigned(BImportParTiers) then
  begin
    if GetParamSocSecur('SO_AFCLIENT', 0, True) <> 8 then         //Si Kpmg
      BImportParTiers.Visible := False
    else
//      BImportParTiers.Visible := True;
      BImportParTiers.Visible := False;

  end;

  // FQ 10848
  BSauve                := TToolbarButton97(GetControl('BSAUVE'));
  if Assigned(BSauve) then
    BSauve.OnClick        := BSauveOnCLick;
  BRestau               := TToolbarButton97(GetControl('BRESTAU'));
  if Assigned(BRestau) then
    BRestau.OnClick       := BRestauOnCLick;

end ;

procedure TOF_RECUPSUSPECT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RECUPSUSPECT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_RECUPSUSPECT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_RECUPSUSPECT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_RECUPSUSPECT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_RECUPSUSPECT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_RECUPSUSPECT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_RECUPSUSPECT.BDupliqueOnCLick(Sender: Tobject);
var
  CodePar : String;

begin
  if VarIsNull(GetField('RSS_PARSUSPECT')) then
    exit;

  CodePar               := GetField('RSS_PARSUSPECT');

  LanceDuplImport ('CODE='+CodePar);
  TFMul(Ecran).BChercheClick(Ecran);

end;


procedure TOF_RECUPSUSPECT.FlisteOnDblClick(Sender: Tobject);
begin
  if VarIsNull(GetField('RSS_PARSUSPECT')) then
    exit;

  EntreeRecupSuspect(GetField('RSS_PARSUSPECT'));
  TFMul(Ecran).BChercheClick(Ecran);

end;

procedure TOF_RECUPSUSPECT.BDeleteOnClick(Sender: Tobject);
begin
  If VarIsNull(GetField('RSS_PARSUSPECT')) then
    exit;

  SupRecupSuspect(GetField('RSS_PARSUSPECT'));
  TFMul(Ecran).BChercheClick(Ecran);

end;

procedure TOF_RECUPSUSPECT.BInsertOnClick(Sender: Tobject);
begin
  EntreeRecupSuspect('');
  TFMul(Ecran).BChercheClick(Ecran);

end;

// FQ 10848 : sauvegarde du paramétrage tablette RTCORRESPIMPORT
procedure TOF_RECUPSUSPECT.BSauveOnCLick(Sender: Tobject);
var Nomfic : string;
  TOBC : TOB;
  Q : TQuery;
begin
  Nomfic := AGLLanceFiche('RT','RTNOMFICSAV','','','');
  if Nomfic<>'' then
  begin
    TOBC := TOB.Create ('Tablette RTCORRESPIMPORT',nil,-1);
    Q :=OpenSQL('Select * from choixcod where cc_type="CIM"',True);
    if not Q.Eof then TOBC.LoadDetailDB('CHOIXCOD','','',Q,False,True);
    Ferme(Q);
    TOBC.SaveToFile(Nomfic,False,true,true,'');
    TOBC.Free;
    PGIInfo('Sauvegarde terminée.',Nomfic);
  end;
end;

// FQ 10848 : restauration du paramétrage tablette RTCORRESPIMPORT
procedure TOF_RECUPSUSPECT.BRestauOnCLick(Sender: Tobject);
var Nomfic : string;
  TOBC : TOB;
begin
  TOBC := Nil;
  Nomfic := AGLLanceFiche('RT','RTNOMFICRES','','','');
  if Nomfic<>'' then
  begin
  if FileExists (NomFic) then
    try
    TOBC := TOB.Create ('CHOIXCOD',nil,-1);
    TobLoadFromFile(NomFic, Nil, TOBC ) ;
    finally
        if (TOBC.Detail.count > 0) then
        begin
          if (TOBC.Detail[0].NomTable<>'CHOIXCOD')or (TOBC.Detail[0].GetString('CC_TYPE')<>'CIM') then
            PGIInfo('Les lignes importées ne sont pas des lignes de la tablette RTCORRESPIMPORT ', 'Restauration RTCORRESPIMPORT')
          else
          begin
            try
              BeginTrans;
              ExecuteSql('Delete from choixcod where cc_type="CIM"');
              TOBC.InsertOrUpdateDB(false);
              CommitTrans;
              PGIInfo('Restauration terminée.',Nomfic);
            Except
              Rollback;
              PGIInfo('Erreur en écriture : restauration non effectuée.',Nomfic);
            end;
          end;
        end;
        TOBC.Free;
    end;

  end;
end;


Initialization
  registerclasses ( [ TOF_RECUPSUSPECT ] ) ;
end.
