{***********UNITE*************************************************
Auteur  ...... : BTY
Créé le ...... : 19/07/2005
Modifié le ... :   /  /
Description .. : Source TOT de la TABLETTE : TIMOTIFCESSION (TIMOTIFCESSION)
Mots clefs ... : TOT;TIMOTIFCESSION
Suite ........ : BTY 12/05 FQ 17122 Interdire la saisie d'un code débutant par des blancs
Suite ........ : BTY 09/07 Interdire la suppression du code 999 Rempl. composant 2e catégorie
*****************************************************************}
Unit TIMOTIFCESSION_TOT ;

Interface

Uses StdCtrls, 
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     uTob, // TQuery
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOT ;

Type
  TOT_TIMOTIFCESSION = Class ( TOT )
    procedure OnNewRecord              ; override ;
    procedure OnDeleteRecord           ; override ;
    procedure OnUpdateRecord           ; override ;
    procedure OnAfterUpdateRecord      ; override ;
    procedure OnClose                  ; override ;
    procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

procedure TOT_TIMOTIFCESSION.OnNewRecord () ;
begin
  Inherited ;
end ;

procedure TOT_TIMOTIFCESSION.OnDeleteRecord () ;

var      CodeSortie : string;
         Existe : integer;
         Q : TQuery;
begin
  Inherited ;
  // BTY 19/07/05
  // Si le code existe pour une cession déjà pratiquée,
  // bloquer la suppression
   CodeSortie := '';
   CodeSortie := GetField ('CC_CODE');
   //YCP 21/10/05 Existe := 0;
   // Lire immolog avec le code opération = CES, sélecter champ il_motifces

   Q:= OpenSQL ('SELECT Count(IL_IMMO) FROM IMMOLOG WHERE IL_TYPEOP="CES" AND IL_MOTIFCES="' + CodeSortie + '"', True);
   Existe := Q.Fields[0].AsInteger;
   Ferme(Q);
//   if (Existe > 0) then
   if (Existe > 0) or (CodeSortie='999') then
     begin
     HShowMessage('8;Motifs de sortie;Suppression non autorisée, motif utilisé.;W;O;O;O;', '', '');
     LastError := 1;
     end;

end ;

procedure TOT_TIMOTIFCESSION.OnUpdateRecord () ;
var CodeSortie : string;
begin
  Inherited ;
  // BTY 12/05 Interdire la saisie d'un code commençant par un espace
  CodeSortie := GetField ('CC_CODE');
  if (DS.State=dsInsert) then
    begin
    if (POS(' ', CodeSortie) = 1) then
      begin
      HShowMessage('8;Motifs de sortie;Le caractère espace n''est pas autorisé dans le code.;W;O;O;O;', '', '');
      LastError := 1;
      end;
    end;
end ;

procedure TOT_TIMOTIFCESSION.OnAfterUpdateRecord () ;
//var CodeSortie : string; BTY 01/06
begin
  Inherited ;
end ;

procedure TOT_TIMOTIFCESSION.OnClose () ;
begin
  Inherited ;
end ;

procedure TOT_TIMOTIFCESSION.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOT_TIMOTIFCESSION ] ) ; 
end.
