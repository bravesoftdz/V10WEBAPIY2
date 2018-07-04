unit UTofAfPiece_Mul;

interface
uses
{$IFDEF EAGLCLIENT}
   eMul,
{$ELSE}
   Mul , {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db,
{$ENDIF}
   StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
   HCtrls,HEnt1,HMsgBox,UTOF, AffaireUtil,Dicobtp,SaisUtil,UTofAfBasePiece_Mul,
   M3FP,aglinit, paramsoc;

Type                      
     TOF_AFPIECE_MUL = Class (TOF_AFBASEPIECE_MUL)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure AlimPieceVivante;
     END ;

Type
     TOF_AFPIECEF_MUL = Class (TOF)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnUpdate ; override ;
     END ;

implementation
//****************  TOF_AFPIECE_MUL ******************************
procedure TOF_AFPIECE_MUL.OnArgument(stArgument : String );
var Critere, Champ, valeur,Nat  : String;
    x : integer;
Begin
Inherited;
// Recup des critères
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN
        X:=pos(':',Critere);
        if x=0 then  X:=pos('=',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end
           else Champ := Critere;
        if Champ = 'ACTION' then SetControlText('XXAction',Critere);
        END;
    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;
{$IFDEF BTP}
SetControlEnabled('GP_VIVANTE',False);
SetControlVisible('GP_VIVANTE',false);
SetControlVisible('AFFAIRE',false);
{$else}
if ((NatPiece<>nil) and  (natpiece.value = '')) then SetControlProperty('GP_VIVANTE','State',cbgrayed);  //mcd 17/04/02 pour avoir cette zone en grisée et passer outre ce qui est fait dans la tof ancêtre si tous en nature (cas appel depuis mission)
if ((NatPieceMul<>nil) and  (natpieceMul.text = '')) then SetControlProperty('GP_VIVANTE','State',cbgrayed);  //mcd 09/10/02
{$ENDIF}

if (NatPiece<>nil) then nat:=   natpiece.value;//mcd 09/10/02
if (NatPieceMul<>Nil) then nat:=   natpieceMul.text;//mcd 09/10/02
if (nat= 'FPR') then Ecran.Caption := 'Modification des factures provisoires'
else if (nat = 'APR') then Ecran.Caption := 'Modification des avoirs provisoires'
else if (nat = 'FRE') then Ecran.Caption := 'Modification des factures externes';

// CB 14/10/2003 factures de regul   
SetControlVisible('FACTREGUL', GetParamSoc('SO_AFREVISIONPRIX'));

UpdateCaption(Ecran);
End;

procedure TOF_AFPIECE_MUL.AlimPieceVivante;
BEGIN
inherited;

END;

//****************  TOF_AFPIECEF_MUL ******************************
//*****************************************************************
procedure TOF_AFPIECEF_MUL.OnArgument(stArgument : String );
var zaff : string;
    Critere, Champ, valeur, LibMul  : String;
    x : integer;
    zdateD,ZdateF,zwhere : String;
Begin
Inherited;
          // mcd 31/05/02 tout revu car la fiche n'est plus faite sur PIECE mais sur LIGNE en regroupé
// Recup des critères
zaff :=' '; LibMul :='';
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
        if Champ = 'GP_AFFAIRE' then Zaff := Valeur;
        if Champ = 'GP_DATEPIECE'then ZDateD := Valeur;
        if Champ = 'GP_DATEPIECE_'then ZDateF := Valeur;
        if Champ = 'XXWHERE'then Zwhere := Valeur;
        END;
    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;


(*if (zaff = ' ') then     // mcd 31/05/02 ??? je ne sais pas qd c'est utiliser, ces champs AFF_AFF* n'existe pas sur la fiche AfPieceF_Mul
   // mcd 26/06/01 ajout avenant
   ChargeCleAffaire(Nil,THEDIT(GetControl('AFF_AFFAIRE1')), THEDIT(GetControl('AFF_AFFAIRE2')),
        THEDIT(GetControl('AFF_AFFAIRE3')),THEDIT(GetControl('AFF_AVENANT')),Nil, taCreat,'',False)
else  *)
ChargeCleAffaire(Nil,THEDIT(GetControl('ZZAFFAIRE1')), THEDIT(GetControl('ZZAFFAIRE2')),
         THEDIT(GetControl('ZZAFFAIRE3')),THEDIT(GetControl('ZZAVENANT')),Nil, taConsult,zaff,False);

UpdateCaption(Ecran);
SetControlText('GL_AFFAIRE',zaff);
SetControlText('GL_NATUREPIECEG','FAC;FRE;AVC');
SetControlText('GL_TYPELIGNE','ART');

if (ZDateD = '') then  ZDateD := DateToStr(idate1900);
if (ZDateF = '') then  ZDateF := DateToStr(idate2099);
SetControlText('GL_DATEPIECE',zdateD);
SetControlText('GL_DATEPIECE_',zdateF);
SetControlText('XX_WHERE',zwhere);

End;

procedure TOF_AFPIECEF_MUL.OnUpdate;
Begin
{$IFDEF AFFAIRE}
 {$IFDEF EAGLCLIENT}
  ModifColAffaireGrid ( TFMul(Ecran).FListe,TFMul(Ecran).Q.tq);  //mcd 27/12/02 ??, manque
  {$ELSE}
//  ModifColAffaireGrid ( TFMul(Ecran).FListe);
  ModifColAffaireGrid ( TFMul(Ecran));
  {$ENDIF}
{$ENDIF}
TFMul(Ecran).FListe.Repaint;
End;


procedure AGLAlimPieceVivante1( parms: array of variant; nb: integer );
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFMul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_AFPIECE_MUL) then TOF_AFPIECE_MUL(MaTOF).AlimPieceVivante else exit;
end;


Initialization
registerclasses([TOF_AFPIECE_MUL,TOF_AFPIECEF_MUL]);
RegisterAglProc( 'AlimPieceVivante1',True,0,AGLAlimPieceVivante1);


end.
