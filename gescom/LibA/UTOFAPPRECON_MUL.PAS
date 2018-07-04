{***********UNITE*************************************************
Auteur  ...... :                                      
Créé le ...... : 29/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : APPRECONMUL ()
Mots clefs ... : TOF;APPRECONMUL
*****************************************************************}
Unit UTOFAPPRECON_MUL ;

Interface

Uses StdCtrls, Controls, Classes,paramsoc,
{$IFDEF EAGLCLIENT}
   emul,MaineAGL,
{$ELSE}
   dbTables, db, mul, FE_Main ,
{$ENDIF}
	 forms, sysutils,  ComCtrls,
   HCtrls, HEnt1, HMsgBox, UTOF,
   utofAfBaseCodeAffaire,M3FP,UtofAppreciation, UtofAppreConAnu_mul;

Type
  TOF_APPRECON_MUL = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (Stargument : String ) ; override ;
    procedure OnClose                  ; override ;

    private
        GcbAffReadOnly:boolean;
    //procedure TypeAppelChargeCle(var AvActionFiche:TActionFiche);
    procedure Appelapprec(stpar : string);
    public
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
  end ;
Procedure AFLanceFiche_Mul_Consult_Apprec(Argument : string);

Implementation

procedure TOF_APPRECON_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_APPRECON_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_APPRECON_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_APPRECON_MUL.OnLoad ;
begin
  Inherited ;
  SetActiveTabSheet('P_PARAMETRE');
end ;

procedure TOF_APPRECON_MUL.OnArgument (StArgument : String ) ;
Var Critere, Champ, valeur  : String;
    x : integer;
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
          if Champ = 'ZAFF' then Zaff := Valeur
          else if champ='TIERS' then
            begin
            SetcontrolText ('AFA_TIERS',Valeur);
            SetcontrolText ('AFA_TIERS_',Valeur);
            SetControlEnabled ('AFA_Tiers',false);
            SetControlEnabled ('AFA_Tiers_',false);
            end
          else if Champ ='DATEDEB' then SetControlText ('AFA_DATEECHE',Valeur)
          else if Champ ='DATEFIN' then SetControlText ('AFA_DATEECHE_',Valeur);
          END;
      Critere:=(Trim(ReadTokenSt(stArgument)));
      END;

  if zAff='' then  GcbAffReadOnly:=false
  else begin
   GcbAffReadOnly:=true;
      //mcd 10/06/02
   SetControlEnabled ('AFf_AFFAIRE1',False);
   SetControlEnabled ('AFf_AFFAIRE2',False);
   SetControlEnabled ('AFf_AFFAIRE3',False);
   SetControlEnabled ('AFf_AVENANT',False);
   SetControlEnabled ('BEFFACEAFF1',False);
   SetControlEnabled ('BSELECTAFF1',False);
   BchangeTiers :=False;
   SetControlText('AFF_AFFAIRE',zaff);
   end;

// Init du code affaire dans la tof ancêtre
  Inherited ;

  SetControlText('AFA_TYPECHE','APP');
end ;

procedure TOF_APPRECON_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_APPRECON_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('AFF_AFFAIRE'));
Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
Tiers:=THEdit(GetControl('AFA_TIERS'));
Tiers_:=THEdit(GetControl('AFA_TIERS_'));
end;

{procedure TOF_APPRECON_MUL.TypeAppelChargeCle(var AvActionFiche:TActionFiche);
begin
if GcbAffReadOnly then AvActionFiche := taConsult else  AvActionFiche := taCreat;
end;
}

procedure TOF_APPRECON_MUL.Appelapprec(stpar : string);
var  Critere, Champ, valeur,zaff,zdatapp,zori  : String;
			x : integer;
begin

 // Recup des critères
Critere:=(Trim(ReadTokenSt(stpar)));
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
        if Champ = 'AFF' then Zaff := Valeur;
        if Champ = 'DATAPP' then Zdatapp := Valeur;

        END;
    Critere:=(Trim(ReadTokenSt(stpar)));
    END;

		if (GetParamSoc('SO_AFAPPPOINT')=true)  then zori:='COM' else zori :='SIM';

    if (zori = 'SIM') then
     // mcd 12/06/02 AglLanceFiche('AFF','AFAPPRECNIV0','','','ZORI:'+zori+';ZAFF:'+zaff+';ZDATAPP:'+zdatapp+';CONSULT:;')
     AFLanceFiche_AppreciationNiv0('ZORI:'+zori+';ZAFF:'+zaff+';ZDATAPP:'+zdatapp+';CONSULT:;')
    else
      // mcd 12/06/02 AglLanceFiche('AFF','AFAPRECOMPLETE','','','ZORI:'+zori+';ZAFF:'+zaff+';ZDATAPP:'+zdatapp+';CONSULT:;');
      AFLanceFiche_AppreciationCompl ('ZORI:'+zori+';ZAFF:'+zaff+';ZDATAPP:'+zdatapp+';CONSULT:;');
end ;



procedure AGLAPPRECONMUL_AppelApprec( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
begin
F:=TForm(Longint(Parms[0]));
if (F is TFMul) then Latof:=TFMul(F).Latof else laTof := nil;
if (LaTof is TOF_APPRECON_MUL) or (Latof is TOF_APPRECONANU_MUL) then
   TOF_APPRECON_MUL(LaTof).AppelApprec(Parms[1]);
end;

Procedure AFLanceFiche_Mul_Consult_Apprec(Argument : string);
begin
AGLLanceFiche('AFF','AFAPPRECON_MUL','','',Argument);
end;

Initialization
  registerclasses ( [ TOF_APPRECON_MUL ] ) ;
  RegisterAglProc( 'AppelApprec',True,1,AGLAPPRECONMUL_AppelApprec);
end.
