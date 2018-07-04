{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 19/02/2003
Modifié le ... : 04/03/2003
Description .. : Source TOM de la TABLE : PLANREF (PLANREF)
Suite ........ : Passage en eAGL
Mots clefs ... : TOM;PLANREF
*****************************************************************}
Unit CPPLANREF_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     forms,
     sysutils,
     ComCtrls,
{$IFDEF EAGLCLIENT}
     MaineAGL,  // AGLLanceFiche
     eFichList, // TFFicheListe
{$ELSE}
     db,
     dbtables,
     FE_Main,   // AGLLanceFiche
     FichList,  // TFFicheListe
{$ENDIF}
     UTILPGI,   // _Blocage,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     UTob ;

function  FicheImportPlanRef(NumPlan : integer ; Compte : String) : String;
procedure FichePlanRef(Compte : String);

Type
  TOM_PLANREF = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BHelpClick(Sender: TObject);
//  private
end ;

Implementation

function FicheImportPlanRef(NumPlan : integer ; Compte : String) : String;
begin
  Result := AGLLanceFiche('CP','CPPLANREF',Compte,Compte,'ACTION=MODIFICATION;0;'+IntToStr(NumPlan)+';'+Compte);
end;

procedure FichePlanRef(Compte : String);
begin
  if _Blocage(['nrCloture'],True,'nrAucun') then Exit ;
  AGLLanceFiche('CP','CPPLANREF','',Compte,'ACTION=MODIFICATION;1');
end;

procedure TOM_PLANREF.OnArgument ( S: String ) ;
var
  szAction : String;
  szNum : String;
  szCompte : String;
begin
  Inherited ;

  // Récupère les arguments
  szAction := ReadTokenSt(S);  // Action
  szNum := ReadTokenSt(S);    // FicheImportPlanRef ou FichePlanRef

  // FicheImportPlanRef
  if (szNum = '0') then begin
    szNum := ReadTokenSt(S);    // NumPlan
    szCompte := ReadTokenSt(S); // Compte
    SetControlEnabled('PAPPLI',False);
    Ecran.Caption := 'Compte : '+szCompte+' du plan de référence n°'+szNum;
    Ecran.HelpContext := 7109080;
    {$IFDEF CCS3}
      SetControlVisible('PR_VENTILABLE2',False);
      SetControlVisible('PR_VENTILABLE3',False);
      SetControlVisible('PR_VENTILABLE4',False);
      SetControlVisible('PR_VENTILABLE5',False);
    {$ENDIF}
    SetControlVisible('HPANEL1',True);
    SetControlVisible('HPANEL2',False);
    SetControlVisible('HPANEL3',True);
    SetControlVisible('FLISTE',False);
    TFFicheListe(Ecran).Width := 331;

    // Evénements des contrôles
    TButton(GetControl('BVALIDER1',True)).OnClick := BValiderClick;
    TButton(GetControl('BFERME1',True)).OnClick := BFermeClick;
    TButton(GetControl('HELPBTN1',True)).OnClick := BHelpClick;
    end
  // FichePlanRef
  else begin
//    NumPlan:=GetParamSocSecur('SO_NUMPLANREF',0) ;
    Ecran.HelpContext := 1315000;

    SetControlVisible('HPANEL1',True);
    SetControlVisible('HPANEL2',False);
    SetControlVisible('HPANEL3',False);
  end;
end ;

procedure TOM_PLANREF.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_PLANREF.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_PLANREF.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_PLANREF.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_PLANREF.OnLoadRecord ;
begin
  Inherited ;
end;

procedure TOM_PLANREF.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_PLANREF.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_PLANREF.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_PLANREF.BValiderClick(Sender: TObject);
begin
  TFFicheListe(Ecran).Retour := '1';
  TFFicheListe(Ecran).BFermeClick(Sender);
end;

procedure TOM_PLANREF.BFermeClick(Sender: TObject);
begin
  TFFicheListe(Ecran).BFermeClick(Sender);
end;

procedure TOM_PLANREF.BHelpClick(Sender: TObject);
begin
  TFFicheListe(Ecran).HelpBtnClick(Sender);
end;

Initialization
  registerclasses ( [ TOM_PLANREF ] ) ;
end.

