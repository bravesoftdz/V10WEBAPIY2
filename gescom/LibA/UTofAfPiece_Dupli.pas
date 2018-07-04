unit UTofAfPiece_Dupli;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
{$IFDEF EAGLCLIENT}
   eMul,
{$ELSE}
   Mul,dbTables, db,
{$ENDIF}
   HCtrls,HEnt1,HMsgBox,UTOF, AffaireUtil,DicoAF,SaisUtil,UTofAfBasePiece_Mul,
   M3FP;

Type
     TOF_AFPIECE_DUPLI = Class (TOF_AFBASEPIECE_MUL)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure AlimPieceVivante;
     END ;


implementation
//****************  TOF_AFPIECE_DUPLI ******************************
procedure TOF_AFPIECE_DUPLI.OnArgument(stArgument : String );
var Critere, Champ, valeur, LibMul,znewnat  : String;
    x : integer;
Begin
Inherited;
LibMul :='';
// Recup des critères
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN
        X:=pos(':',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;
        if Champ = 'DUPPLIC' then Znewnat := Valeur;
        END;
    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;
    SetControltext('NEWNATURE',znewnat);

    Ecran.caption := 'Création de factures définitives par duplication';
    if (Znewnat = 'FPR') then
       Ecran.caption := 'Création de factures provisoires par duplication';
    if (Znewnat = 'AVC') then
       Ecran.caption := 'Création d''avoirs par duplication';
    if (Znewnat = 'APR') then
       Ecran.caption := 'Création d''avoirs provisoires par duplication';

    UpdateCaption(Ecran);


End;

procedure TOF_AFPIECE_DUPLI.AlimPieceVivante;
BEGIN
inherited;
if Natpiece.value = '' then   SetControlProperty('GP_VIVANTE','State',cbgrayed); //mcd 07/10/03 10433
END;

procedure AGLAlimPieceVivante2( parms: array of variant; nb: integer );
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFMul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_AFPIECE_DUPLI) then TOF_AFPIECE_DUPLI(MaTOF).AlimPieceVivante else exit;
end;

Initialization
registerclasses([TOF_AFPIECE_DUPLI]);
RegisterAglProc( 'AlimPieceVivante2',True,0,AGLAlimPieceVivante2);

end.
