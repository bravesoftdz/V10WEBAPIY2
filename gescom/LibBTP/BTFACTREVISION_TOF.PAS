{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 02/02/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTFACTREVISION ()
Mots clefs ... : TOF;BTFACTREVISION
*****************************************************************}
Unit BTFACTREVISION_TOF ;

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
     uEntCommun,
     UTOF ;

Type
  TOF_BTFACTREVISION = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;


procedure RevisionFromDevis (Tobpiece,TOBRevisions : TOb);
procedure RevisionFromFacture (Tobpiece,TOBRevisions : TOb);
procedure LoadLesRevisions (TOBPiece,TOBrevisions : TOB);
procedure ValideLesRevisions (TOBpiece,TOBrevisions : TOB);

Implementation
uses UtilTOBPiece,FactTOB;

procedure ValideLesRevisions (TOBpiece,TOBrevisions : TOB);
begin

end;

procedure LoadLesRevisions (TOBPiece,TOBrevisions : TOB);
var Sql : string;
    cledoc : r_cledoc;
    QQ : TQuery;
begin
  (*
  CleDoc := TOB2CleDoc(TOBPiece);
  Sql := 'SELECT * FROM BREVISIONS WHERE '+WherePiece(cledoc,ttdRevision,false);
  QQ := OpenSQL(Sql,True,-1,'',true);
  if not QQ.eof then
  begin
    TOBrevisions.LoadDetailDB('BREVISIONS','','',QQ,false);
  end;
  ferme (QQ);
  *)
end;

procedure RevisionFromDevis (Tobpiece,TOBRevisions : TOb);
begin

end;

procedure RevisionFromFacture (Tobpiece,TOBRevisions : TOb);
begin

end;

procedure TOF_BTFACTREVISION.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTFACTREVISION.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTFACTREVISION.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTFACTREVISION.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTFACTREVISION.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_BTFACTREVISION.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTFACTREVISION.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTFACTREVISION.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTFACTREVISION ] ) ; 
end.
