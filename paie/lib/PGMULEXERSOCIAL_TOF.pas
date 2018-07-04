{***********UNITE*************************************************
Auteur  ...... : EPI
Créé le ...... : 14/02/2006
Modifié le ... : 14/02/2006
Description .. : Source TOF de la FICHE : PGMULEXERSOCIAL ()
Mots clefs ... : TOF;PGMULEXERSOCIAL
La TOF n'est pas utilisée, l'appel de la fiche se fait dans le multicritères
*****************************************************************}
Unit PGMULEXERSOCIAL_TOF ;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,Mul,FE_Main,
{$else}
     eMul, uTob, MaineAgl,
{$ENDIF}
     HCtrls,
     UTOF ,
     HTB97;

Type
  TOF_PGMULEXERSOCIAL = Class (TOF)
  public
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (StArgument : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnDisplay                ; override ;
    procedure OnCancel                 ; override ;

  private
    Appel: string;  // selon appel du menu
    procedure BinsertClick(Sender: TObject);
    procedure GrilleDblClick(Sender: TObject);
  end ;

Implementation

procedure TOF_PGMULEXERSOCIAL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGMULEXERSOCIAL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGMULEXERSOCIAL.OnArgument (StArgument : String ) ;
 var
 	Arg : string;
	Binsert : ttoolbarbutton97;
	Creat: Boolean;

begin
  Inherited ;
   Arg := StArgument;
   Appel := readtokenst(Arg);

   TFMul(Ecran).FListe.OnDblClick := GrilleDblClick;

	 Creat := True;
 	 // Appel sur le menu Paie Outils
   If Appel = 'O'
   then Creat := False;
   SetControlEnabled('BINSERT',Creat);

   Binsert := ttoolbarbutton97(getcontrol('BINSERT'));
   if Binsert <> nil then Binsert.OnClick := BinsertClick;
end ;

procedure TOF_PGMULEXERSOCIAL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGMULEXERSOCIAL.BinsertClick(Sender: TObject);
begin
  AglLanceFiche ('PAY','EXERSOCIAL','' , '', 'ACTION=CREATION');
end;

procedure TOF_PGMULEXERSOCIAL.GrilleDblClick(Sender: TObject);
var
   Annee : string;
begin
   Annee :=TFMUL(Ecran).Q.findfield('PEX_EXERCICE').Asstring;
   // AglLanceFiche('PAY','EXERSOCIAL','PEX_EXERCICE='+Annee,Annee,appel);
   // AglLanceFiche('PAY','EXERSOCIAL','',Annee,'ACTION=MODIFICATION');
   AglLanceFiche('PAY','EXERSOCIAL','',Annee,Appel);
end;

procedure TOF_PGMULEXERSOCIAL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGMULEXERSOCIAL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGMULEXERSOCIAL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGMULEXERSOCIAL.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [TOF_PGMULEXERSOCIAL] ) ;
end.
