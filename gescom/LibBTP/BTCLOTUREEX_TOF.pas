{***********UNITE*************************************************
Auteur  ...... : LS
Créé le ...... : 23/04/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCLOTUREEX ()
Mots clefs ... : TOF;BTCLOTUREEX
*****************************************************************}
Unit BTCLOTUREEX_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     eAgl,
     uTob,
{$ENDIF}
     Ent1,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     vierge,
     hpanel,
     windows,
     messages,
     UTOF ;

Type
  TOF_BTCLOTUREEX = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    // Exo à cloturer et Exo suivant
    Exo1    : tExoDate;
    DateDebutEx,DateFinEx : THEdit;
    procedure LanceTraitement (Sender : Tobject);
    procedure ClotureExBtp;

  end ;

procedure BTClotureEx;

Implementation

function ExistePieceNonPasseCtpa : boolean;
var QQ : TQuery;
begin
  QQ := OpenSql('SELECT E_NUMEROPIECE FROM ECRITURE WHERE E_DATECOMPTABLE >="'+USDATETIME(VH^.EnCours.Deb)+'" '+
        'AND E_DATECOMPTABLE <="'+USDATETIME(VH^.EnCours.Fin)+'" AND E_EXPORTE<>"X"',True);
  result := not QQ.eof;
  ferme (QQ);
end;

function ControleAvantClotureOk : boolean;
begin
  result := true;
  if (VH^.Suivant.Code = '') then
  begin
    PGIError ('La clôture ne peut être effectuée. L''exercice suivant n''est pas ouvert') ;
    result := false;
    exit;
  end;
  if ExistePieceNonPasseCtpa then
  begin
    PGIError ('La clôture ne peut être effectuée. Il existe des pièces sur l''exercice précédent non passées en comptabilité') ;
    result := false;
    exit;
  end;
end;

procedure BTClotureEx;
begin
  if not ControleAvantClotureOk then exit;
  AGLLanceFiche ('BTP','BTCLOTUREEX','','','ACTION=MODIFICATION');
end;

procedure TOF_BTCLOTUREEX.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTCLOTUREEX.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTCLOTUREEX.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTCLOTUREEX.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTCLOTUREEX.OnArgument (S : String ) ;
begin
  Inherited ;
  //
  DATEDEBUTEX := THedit (GetControl('DATEDEBUTEX'));
  DATEFINEX := THedit (GetControl('DATEFINEX'));
  //
  TToolbarButton97 (GetControl('BCLOTURE')).OnClick := LanceTraitement;
  Exo1     := VH^.EnCours;
  ExotoDates(Exo1.Code,DateDebutEx,DateFinEx);
end ;

procedure TOF_BTCLOTUREEX.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTCLOTUREEX.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCLOTUREEX.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTCLOTUREEX.LanceTraitement(Sender: Tobject);
begin
  if (PgiAsk ('Etes-vous sur de vouloir clôturer l''exercice désigné ?')=Mryes) then
  begin
    if (PgiAsk ('Les pièces de cet exercice ne seront plus modifiable. Continuer ?')=Mryes) then
    begin
      V_PGI.Ioerror := OeOk;
      if TRANSACTIONS (ClotureExBtp,0) <> OeOk then
      begin
        PgiError ('Une erreur à eu lieu durant la clôture.'); 
      end else
      begin
        RechargeParamSoc;
        PgiInfo ('Clôture effectuée.');
        PostMessage (Ecran.Handle, WM_CLOSE, 0, 0);
      end;
    end;
  end;
end;

procedure TOF_BTCLOTUREEX.ClotureExBtp;
var Sql : String;
begin
  Sql := 'UPDATE PIECE SET GP_VIVANTE="-" WHERE GP_NATUREPIECEG IN ("FBT","ABT") AND GP_DATEPIECE <= "'+USDATETIME(Exo1.Fin)+'"';
  ExecuteSql (Sql);
  Sql := 'UPDATE EXERCICE SET EX_ETATCPTA="CDE" WHERE EX_EXERCICE="'+Exo1.Code +'"';
  if ExecuteSql (Sql) < 0 then V_PGI.IOError := OeUnKnown;
end;

Initialization
  registerclasses ( [ TOF_BTCLOTUREEX ] ) ;
end.

