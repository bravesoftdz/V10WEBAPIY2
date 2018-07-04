{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 12/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPARAMGENS1 ()
Mots clefs ... : TOF;BTPARAMGENS1
*****************************************************************}
Unit BTPARAMGENS1_TOF ;

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
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     BTPARAMSOCGEN_TOF,
     HTB97,
     FileCtrl,
     UTob,
     UtilLine,
     UTOF ;

Type
  TOF_BTPARAMGENS1 = Class (TOF_BTPARAMSOCGEN)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private
    Metre : THCheckBox;
    BMail	: TToolbarButton97;

  	function VerifInfos : boolean;
    procedure MetreClick (sender : Tobject);
    procedure OutlookClick (sender : Tobject);

  end ;

Implementation

procedure TOF_BTPARAMGENS1.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMGENS1.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMGENS1.OnUpdate ;
begin
  Inherited ;
  if not verifInfos then
  begin
  	TForm(Ecran).ModalResult:=0;
    exit;
  end;
	StockeInfos(Ecran, LaTob);
end ;

procedure TOF_BTPARAMGENS1.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMGENS1.OnArgument (S : String ) ;
Var UneTob : TOB;
    Grp		 : TGroupBox;
begin
  Inherited ;

  chargeEcran(ecran, LaTob);

  //Affichage du groupe paramétre Echange comptable
  UneTOB := LaTOB.findFirst(['SOC_NOM'],['SO_BTLIENCPTAS1'],true);
  if UneTOB <> Nil then
     if UneTOB.GetValue('SOC_DATA')='AUC' then
        begin
        Grp := TGroupBox(Ecran.FindComponent('GP_GENERAL'));
        ecran.Height := Grp.Height;
        Grp := TGroupBox(Ecran.FindComponent('GP_METRE'));
        ecran.Height := Ecran.height + (Grp.Height * 2 );
				Grp := TGroupBox(Ecran.FindComponent('GP_SAUVEGARDE'));
        ecran.Height := Ecran.Height + Grp.Height;
        SetControlProperty('GP_CHEMACC', 'Visible', False);
        end;

  Metre := THCheckBox(ecran.FindComponent('SO_METRESEXCEL'));
  Metre.onClick := MetreClick;

  BMail := TToolbarButton97(ecran.FindComponent('BMAIL'));
  BMail.onclick := OutlookClick;

  MetreClick(self);

end ;

procedure TOF_BTPARAMGENS1.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMGENS1.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMGENS1.OnCancel () ;
begin
  Inherited ;
end ;

function TOF_BTPARAMGENS1.VerifInfos: boolean;
var CH 				: THedit;
    Cx 				: THCheckbox;
begin

	result := false;

  CH := THEdit(GetCOntrol('SO_BTREPMETR')) ;
  Cx := THCheckBox(GetControl('SO_BTMETREDOC'));
  if Cx.checked then
     Begin
     if (CH.text = '' ) or (not DirectoryExists(CH.text)) then
    	  Begin
   	    HShowMessage('3;Société;Vous devez renseigner un répertoire valide pour stocker vos Fichiers métrés;W;O;O;O;','','') ;
        CH.SetFocus ;
        Exit ;
        end;
     end;

  CH := THEdit(GetControl('SO_CPRDREPERTOIRE'));
  if (CH.text = '' ) or (not DirectoryExists(CH.text)) then
 	   Begin
   	 HShowMessage('3;Société;Vous devez renseigner un répertoire valide pour stocker vos Fichiers d''export;W;O;O;O;','','') ;
     CH.SetFocus ;
     Exit ;
     end;

  CH := THEdit(GetControl('SO_FICHIERSAUVE'));
  if (CH.text = '' ) or (not DirectoryExists(CH.text)) then
 	   Begin
   	 HShowMessage('3;Société;Vous devez renseigner un répertoire valide pour stocker vos sauvegardes;W;O;O;O;','','') ;
     CH.SetFocus ;
     Exit ;
     end;

  result := true;

end;

procedure TOF_BTPARAMGENS1.MetreClick(sender: Tobject);
begin
	THCritMaskEdit(GetCOntrol('SO_BTREPMETR')).enabled := (not Metre.Checked);
end;

procedure TOF_BTPARAMGENS1.OutlookClick(sender: Tobject);
Var ParametreMail : ParamMail;
    AdrMail				: ThEdit;
Begin

  AdrMail := THEdit(GetControl('SO_CPRDEMAILCLIENT'));

  ParametreMail.Email := AdrMail.text;
  ParametreMail.MessMail := 'Test d''envoi de Mail à l''adresse suivante : ' + ParametreMail.Email;
  ParametreMail.Sujet := 'Envoi Mail Cabinet comptable';
  ParametreMail.Fichier := '';
  ParametreMail.CopieA := '';

  EnvoieMail(ParametreMail);

end;

Initialization
  registerclasses ( [ TOF_BTPARAMGENS1 ] ) ; 
end.

