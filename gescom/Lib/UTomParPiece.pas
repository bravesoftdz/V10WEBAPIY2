unit UTomParPiece;

interface

uses  UTOM,
      HCtrls,
      Classes,
      Controls,
      ComCtrls,
      Sysutils,
      HDB,
      Forms,
      Hent1,
      LicUtil,
      Graphics,
      StdCtrls,
{$IFNDEF EAGLCLIENT}
      db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$IFDEF V530} EdtEtat, {$ELSE} EdtREtat, {$ENDIF}
      mul,
{$else}
      eMul,
      uTob,
      UtileAGL,
{$ENDIF}
{$IFNDEF SANSCOMPTA}	Souche_TOM, {$ENDIF}
      HTB97,
      HMsgBox,
      M3FP;

Type
     TOM_ParPieceCompl = Class (TOM)
       procedure OnNewRecord  ; override ;
       procedure OnUpdateRecord  ; override ;
       procedure OnArgument (stArgument : String ) ; override ;
       procedure OnLoadRecord ; override ;
     Private
       Nature : string;
     END ;


const TexteMessage: array[1..3] of string 	= (
          {1}        'Vous devez renseigner un établissement.',
          {2}        'Vous devez renseigner un compteur.',
          {3}        'Vous devez renseigner un libellé.'
          );

implementation

// Récupération de la nature de pièce en argument
procedure TOM_ParPieceCompl.OnArgument(stArgument : String );
var x : integer;
    critere: string;
    Arg,Val : string;
BEGIN
Inherited;

  Critere := (Trim(ReadTokenSt(stArgument)));

  if Critere<>'' then
  BEGIN
    x:= pos('=', Critere);
    if x<>0 then
    BEGIN
      Arg:= copy (Critere, 1, x-1);
      Val:= copy (Critere, x+1, length(Critere));
      if Arg='NATURE' then Nature := Val;
    END;
  END;

  //mcd 05/04/02 pour être en phase avec appel fihce parpiece ou parpieceusr en fct mot passe du jour
  // pas fait pour mode qui n'accede qu'a parpiece
  if not(ctxmode in V_PGI.PGICOntexte) and (V_PGI.PassWord <> CryptageSt(DayPass(V_PGI.DateEntree))) then
  begin
    SetControlEnabled('GPC_TYPEECRCPTA',False);
    SetControlProperty('GPC_TYPEECRCPTA','Color',ClbTnface);
    SetControlEnabled('GPC_TYPEPASSACC',False);
    SetControlProperty('GPC_TYPEPASSACC','Color',Clbtnface);
  end;

END;

procedure TOM_ParPieceCompl.OnUpdateRecord;
BEGIN
Inherited;
  if GetControlText('GPC_ETABLISSEMENT')='' then
  Begin
    SetFocusControl('GPC_ETABLISSEMENT'); LastError:=1; LastErrorMsg:=TexteMessage[LastError]; exit;
  End;

  if GetControlText('GPC_SOUCHE')='' then
  Begin
    SetFocusControl('GPC_SOUCHE'); LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit;
  End;

  if GetControlText('GPC_LIBELLE')='' then
  Begin
    SetFocusControl('GPC_LIBELLE'); LastError:=3; LastErrorMsg:=TexteMessage[LastError]; exit;
  End;

END;

procedure TOM_ParPieceCompl.OnNewRecord;
BEGIN
Inherited;
if (Nature <> '') then SetField ('GPC_NATUREPIECEG',Nature);
END;


procedure TOM_ParPieceCompl.OnLoadRecord;
begin
Inherited;
//THEdit(Ecran.findComponent('NATURE')).text := GetField('GPC_NATUREPIECEG') ;
end;

{$IFNDEF SANSCOMPTA}
Procedure TOMParPiece_Souche(Parms : Array of variant; nb: integer) ;
BEGIN
FicheSouche('GES');
END;

procedure InitTOMParPiece();
begin
RegisterAglProc( 'ParPiece_OuvreSouche', True , 0, TOMParPiece_Souche);
end;
{$ENDIF}

// MAJ du 01/08/2001
// Procédure d'édition de la fiche nature par pièce
procedure AGLLanceEtatNature ( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     NaturePiece, stWhere,Modele : string;
begin
     F:=TForm(Longint(Parms[0])) ;
     if F.Name<>'GCPARPIECE'then exit;
     NaturePiece:=string(Parms[1]);
     stWhere:=' GPP_NATUREPIECEG="'+NaturePiece+'"';
     Modele:='NAT';   //GetParamSoc('SO_GCETATARTICLE');
     if Modele<>'' then LanceEtat('E','NAT',Modele,True,False,False,Nil,stWhere,'',False)
end;

{ TOM_ParPiece }

Initialization
registerclasses([TOM_PARPIECECOMPL]) ;
RegisterAglProc( 'LanceEtatNature', TRUE , 1, AGLLanceEtatNature);
{$IFNDEF SANSCOMPTA}
InitTOMParPiece
{$ENDIF}
end.
