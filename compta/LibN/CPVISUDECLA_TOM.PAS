{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 24/01/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : ECRITURE (ECRITURE)
Mots clefs ... : TOM;ECRITURE
*****************************************************************}
Unit CPVISUDECLA_TOM ;       

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     FichGrid,
     HDB, 
     Fe_Main,  //AGLLanceFiche
{$else}
     MaineAGL, //AGLLanceFiche
     eFichGrid,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM, 
     UTob,
     uLibExportTVA ;

procedure LanceVisuDeclaTVA(Argument : string);

Type
  TOM_ECRITURE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    public
    private
      {$IFDEF EAGLCLIENT}
      FListe : THGrid;
      {$ELSE}
      FListe : THDBGrid;
      {$ENDIF EAGLCLIENT}
      maDeclaration      : TDeclaration;
      NumDeclaration     : Integer
    end ;

Implementation
procedure LanceVisuDeclaTVA(Argument : string);
begin
  AGLLanceFiche('CP','CPVISUDECLATVA','','',Argument);
end;

procedure TOM_ECRITURE.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_ECRITURE.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_ECRITURE.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_ECRITURE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_ECRITURE.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_ECRITURE.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_ECRITURE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_ECRITURE.OnArgument ( S: String ) ;
var
  x: integer;
  critere: string;
  Arg, Val: string;
begin
  Inherited ;

  // Gestion des arguments
  repeat
    Critere := uppercase(Trim(ReadTokenSt(S)));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        Arg := copy(Critere, 1, x - 1);
        Val := copy(Critere, x + 1, length(Critere));
        if Arg = 'CODECPT' then NumDeclaration := StrToInt(Val);
      end;
    end;
  until Critere = '';

  // Création des objets
  maDeclaration := TDeclaration.Create(NumDeclaration);
  FListe := TFFicheGrid(ecran).FListe;

  // Chargement de la grille
  maDeclaration.Enregistrements.PutGridDetail(FListe,true,true,'',true);
end ;

procedure TOM_ECRITURE.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_ECRITURE.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_ECRITURE ] ) ; 
end.
