{***********UNITE*************************************************
Auteur  ...... : Lionel SANTUCCI
Créé le ...... : 12/02/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSELECTPHASE ()
Mots clefs ... : TOF;BTSELECTPHASE
*****************************************************************}
Unit SelectPhase ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     AglInit,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     FE_Main,
{$else}
     eMul,
     uTob,
     Maineagl,
{$ENDIF}
     forms,
     UtilPhases,
     HTB97,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOB,
     UTOF ;

Type
  TOF_BTSELECTPHASE = Class (TOF)
  private
    TOBParam : TOB;
    TV_PHASE : TTreeView;
    TOBphases,lesPhases : TOB;
    TheNodes : TTreeNode;
    procedure ChargeTV;
    procedure constitueTV (TV : TTreeView;TheNode : TTreeNode;thePhases : TOB; Niveau : integer);
    procedure TV_PHASESelect(Sender: TOBject);
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

function SelectionPhase (Chantier  : string; var Code : string) : boolean;

Implementation

function SelectionPhase (Chantier  : string; var Code : string) : boolean;
var QQ : TQuery;
    UneTOB : TOB;
begin
  result := false;
  QQ := OpenSql ('SELECT COUNT(*) FROM PHASESCHANTIER WHERE BPC_AFFAIRE="'+Chantier+'"',true,-1,'',true);
  if not QQ.eof Then
  begin
    UneTOB := TOB.Create ('LE RETOUR',nil,-1);
    TRY
      UneTOB.AddChampSupValeur ('CHANTIER',Chantier);
      UneTOB.AddChampSupValeur ('PHASE','');
      TheTOB := UneTOB;
      AGLLanceFiche('BTP','BTSELECTPHASE','','','MODIFICATION') ;
      if TheTOB <> nil then
      begin
        Code := TheTOB.GetValue('PHASE');
        result := true;
      end;
    FINALLY
      UneTOB.free;
      TheTOB := nil;
    END;
  end else PGIBox (TraduireMemoire('Aucune phase n''est définie'),TraduireMemoire('Sélection de phase'));
  ferme (QQ);
end;

procedure TOF_BTSELECTPHASE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSELECTPHASE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSELECTPHASE.OnUpdate ;
var TheNode : TTreeNode;
    UneTOB : TOB;
begin
  Inherited ;
  TheNode := TV_PHASE.Selected;
  if TheNode.Data <> nil then
  begin
    UneTOB := TOB(TheNode.data);
    TOBparam.PutValue ('PHASE', UneTOB.GetValue('BPC_PHASETRA'));
    TheTOB := TOBParam;
  end else
  begin
    // sélection sur le chantier donc phase --> non définie
    TOBparam.PutValue ('PHASE', '');
    TheTOB := TOBParam;
  end;
end ;

procedure TOF_BTSELECTPHASE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTSELECTPHASE.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBPhases := TOB.create ('LES PHASES',nil,-1);
  lesPhases := TOB.create ('LES PHASES TRIES',nil,-1);
  TOBParam := LaTOB;
  TV_PHASE := TTreeView (GetControl('TV_PHASE'));
  TV_Phase.OnDblClick := TV_PHASESelect;
  ChargeTV;
end ;

procedure TOF_BTSELECTPHASE.OnClose ;
begin
  Inherited ;
  TOBPhases.free;
  lesPhases.free;
end ;

procedure TOF_BTSELECTPHASE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSELECTPHASE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTSELECTPHASE.constitueTV (TV : TTreeView;TheNode : TTreeNode;thePhases : TOB; Niveau : integer);
var Indice : integer;
    NodeSuiv : TTreeNode;
    niveausuiv : integer;
    TOBL : TOB;
begin
  for Indice := 0 to ThePhases.detail.count -1 do
  begin
    TOBL := ThePhases.detail[Indice];
    NodeSuiv := TV.Items.AddChild (TheNode,TOBL.GetValue('BPC_PHASETRA')+' '+TOBL.GetValue('BPC_LIBELLE'));
    NodeSuiv.Data := TOBL;
    if TOBL.detail.count > 0 then
    begin
      NiveauSuiv := Niveau + 1;
      ConstitueTV (TV,NodeSUiv,TOBL, NiveauSuiv);
    end;
  end;
end;

procedure TOF_BTSELECTPHASE.ChargeTV;
var QQ : TQuery;
    req : String;
    TheNode : TTreeNode;
begin
(*
  Req := 'SELECT * FROM PHASESCHANTIER WHERE BPC_AFFAIRE="'+
          TOBParam.getValue('CHANTIER')+'" ORDER BY BPC_PN1,BPC_PN2,BPC_PN3,BPC_PN4,BPC_PN5,BPC_PN6,'+
          'BPC_PN7,BPC_PN8,BPC_PN9';
*)
  Req := 'SELECT * FROM PHASESCHANTIER WHERE BPC_AFFAIRE="'+
          TOBParam.getValue('CHANTIER')+'" ORDER BY BPC_PHASETRA';
  QQ := OpenSql (Req,true,-1,'',true);
  TOBPhases.LoadDetailDB ('PHASESCHANTIER','','',QQ,false,true);
  ferme (QQ);

  ConstitueLaTOBTrieSurphase (TOBPhases,lesPhases);

  TheNode := TV_PHASE.Items.Add( nil,'Chantier '+TOBparam.GetValue('CHANTIER'));
  constitueTV (TV_PHASE,TheNode,LesPhases,1);
  TheNode.Expanded := true;
end;

procedure TOF_BTSELECTPHASE.TV_PHASESelect (Sender : TOBject);
begin
end;

Initialization
  registerclasses ( [ TOF_BTSELECTPHASE ] ) ;
end.

