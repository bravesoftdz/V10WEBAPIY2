{***********UNITE*************************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 15/11/2001
Modifié le ... : 10/04/2002
Description .. : Source TOF de la TABLE : PARAMLET ()
Suite ........ : Ecran : CPARAMLET
Mots clefs ... : TOF;PARAMLET
*****************************************************************}
Unit UTOFPARAMLET ;

Interface

Uses StdCtrls, Controls, Classes,  forms, sysutils, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOF,
     Vierge,
     windows,
     HTB97,
     messages
      ;

Type
  TOF_PARAMLET = Class (TOF)
   private
    FStRetour : string;  // chaine contenant l'option choisit par l'utilisateur
    FRdgChoix : THRadioGroup;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BValiderClick(Sender: TObject); // Bouton Valider
    procedure BFermeClick(Sender: TObject); // bouton ferme
    procedure RadioButtonClick(Sender: TObject);
   protected
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay ; override ;    
  end ;

Implementation


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... : 30/09/2003
Description .. : Initialisation des contrôle de la fiche
Suite ........ : si on passe une chaine diff de '1' en parametre, on n'a pas
Suite ........ : acces a l'option 'ecriture de regul' de lettrage'
Suite ........ : - LG - 30/09/2003 - la variable FStRetour pouvait ne pas 
Suite ........ : avoir de valeur par defaut
Mots clefs ... : 
*****************************************************************}
procedure TOF_PARAMLET.OnArgument (S : String ) ;
var
 lStArg  : string;
begin
 Inherited ;
 V_PGI.ZoomOLE:=false;
 lStArg := ReadTokenST(S);
 FStRetour := 'AL1'; // ecriture de regul par defaut
 if lStArg = '1' then
  begin
   TRadioButton(GetControl('GSOLDE')).Checked         := true;
   TRadioButton(GetControl('GPARTIEL')).Checked       := false;
  end
   else
    if lStArg = '2' then
     begin
      TRadioButton(GetControl('GPARTIEL')).Checked     := true ;
      TRadioButton(GetControl('GSOLDE')).Checked       := false ;
      TRadioButton(GetControl('GSOLDE')).Enabled       := false ;
      TRadioButton(GetControl('GECR')).Enabled         := false ;
      FStRetour                                        := 'AL2'; // lettrage partiel par defaut
     end
      else
       begin
        FStRetour := 'AL2'; // lettrage partiel par defaut
        TRadioButton(GetControl('GSOLDE')).Enabled       := false;
        TRadioButton(GetControl('GSOLDE')).Checked       := false;
        TRadioButton(GetControl('GPARTIEL')).Checked     := true;
        Ecran.ActiveControl                              := TRadioButton(GetControl('GPARTIEL')); // donne le focus a la case cocher 'lettrage partiel'
       end; // if
 FRdgChoix                                            := THRadioGroup(GetControl('FE__HRADIOGROUP_CHOIX'));

 Ecran.OnKeyDown                                      := FormKeyDown;
 TToolbarButton97(GetControl('BValider')).Onclick     := BValiderClick;
 TToolbarButton97(GetControl('BFerme')).Onclick       := BFermeClick;
 TRadioButton(GetControl('GSOLDE')).Onclick           := RadioButtonClick;
 TRadioButton(GetControl('GPARTIEL')).Onclick         := RadioButtonClick;
 TRadioButton(GetControl('GECR')).Onclick             := RadioButtonClick;
 TRadioButton(GetControl('GANNUL')).Onclick           := RadioButtonClick;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... :   /  /    
Description .. : Le F10 ferme la fiche
Mots clefs ... :
*****************************************************************}
procedure TOF_PARAMLET.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 if ( csDestroying in Ecran.ComponentState ) then Exit ;
 case Key of
  VK_F10 : begin
            Key := 0 ;
            TFVierge(Ecran).retour := FStRetour;
            Ecran.Close;
           end;
 end; // case
end;

procedure TOF_PARAMLET.BValiderClick(Sender: TObject);
begin
 TFVierge(Ecran).retour := FStRetour;
end;

procedure TOF_PARAMLET.BFermeClick(Sender: TObject);
begin
 TFVierge(Ecran).retour := 'AL4';
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... :   /  /    
Description .. : decode le tag des bouttons pour retourner l'options choisit. 
Suite ........ : les valeurs de retours correspondent à la tablette 
Suite ........ : TTLETMODE.
Suite ........ : AL1 : ecriture de regul
Suite ........ : AL2 : lettrage partiel
Suite ........ : AL3 : ecriture simplifie
Suite ........ : AL4 : annulation
Mots clefs ... : 
*****************************************************************}
procedure TOF_PARAMLET.RadioButtonClick(Sender: TObject);
begin

 case TRadioButton(Sender).Tag of 
  0 : FStRetour := 'AL1';
  1 : FStRetour := 'AL2';
  2 : FStRetour := 'AL3';
  3 : FStRetour := 'AL4';
 end; // case

end;

procedure TOF_PARAMLET.OnDisplay;
begin
  inherited;
end;

Initialization
  registerclasses ( [ TOF_PARAMLET ] ) ;
end.

