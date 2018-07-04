{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 29/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : APPRECONANUMUL ()
Mots clefs ... : TOF;APPRECONMUL
*****************************************************************}
Unit UTOFAPPRECONANU_MUL ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
   emul,MaineAGL,
{$ELSE}
   dbTables, db, mul, FE_Main ,
{$ENDIF}
	 forms, sysutils,  ComCtrls,
   HCtrls, HEnt1, HMsgBox, UTOF,
   utofAfBaseCodeAffaire,M3FP,utob,utilmultrt, UtofAppreciation;

Type
  TOF_APPRECONANU_MUL = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnArgument (Stargument : String ) ; override ;
    procedure OnLoad ; override ;

    private
    procedure Annuleapprec(stpar : string);
    public
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
    
  end ;

Procedure AFLanceFiche_Mul_Annu_Apprec;
Implementation



procedure TOF_APPRECONANU_MUL.OnArgument (StArgument : String ) ;
Var Critere, Champ, valeur  : String;
    x : integer; //,znumapp
    zaff : string;
begin

  zaff :='';
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
          if Champ = 'ZAFF' then Zaff := Valeur;
          END;
      Critere:=(Trim(ReadTokenSt(stArgument)));
      END;
  SetControlText('AFF_AFFAIRE',zaff);
// Init du code affaire dans la tof ancêtre
  Inherited ;

  SetControlText('AFA_TYPECHE','APP');
end;

procedure TOF_APPRECONANU_MUL.OnLoad ;
begin
  Inherited ;
  SetActiveTabSheet('P_PARAMETRE');
end ;


procedure TOF_APPRECONANU_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
  Aff:=THEdit(GetControl('AFF_AFFAIRE'));
  Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
  Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
  Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
  Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
  Aff4:=THEdit(GetControl('AFF_AVENANT'));
end;
    

procedure TOF_APPRECONANU_MUL.AnnuleApprec(stpar : string);
var   //Critere, Champ, valeur ,zaff, zdatapp,zori  : String;
			wi : integer;
      TobMAff,TobDet : TOB;
begin
  if PGIAsk('Confirmez l''annulation des appréciations sélectionnées?',Ecran.Caption)=mrno
    then  Exit;
  TobMAff := Tob.Create('les affaires',NIL,-1);
  TraiteEnregMulListe (TFMul(Ecran), 'AFF_AFFAIRE','AFACTTIERSAFFAIRE', TobMAff, True);
  for wi :=0 to TobMAff.Detail.count-1  do
    begin
    TobDet := TobMAff.Detail[wi];
    AnnulationApprec (Tobdet.GetValue('AFA_AFFAIRE'),Tobdet.GetValue('AFA_NUMECHE'),False);
    end;
  TobMaff.free;
end;

procedure AGLAnnuleApprec( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
begin
  F:=TForm(Longint(Parms[0]));
  if (F is TFMul) then Latof:=TFMul(F).Latof else laTof := nil;
  if (Latof is TOF_APPRECONANU_MUL) then TOF_APPRECONANU_MUL(LaTof).AnnuleApprec(Parms[1]);
end;


Procedure AFLanceFiche_Mul_Annu_Apprec;
begin
AGLLanceFiche('AFF','AFAPPRECONANU_MUL','','','');
end;

Initialization
  registerclasses ( [ TOF_APPRECONANU_MUL ] ) ;
  RegisterAglProc('AnnuleApprec',TRUE,1,AGLAnnuleApprec);
end.
