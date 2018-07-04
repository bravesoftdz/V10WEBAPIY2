{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 19/10/2001
Modifié le ... : 18/06/2007
Description .. : Source TOF de la TABLE : BALSITCREATION ()
Suite ........ : 
Suite ........ : FQ 17107 : SBO 03/01/2006 : Ajout des écritures de
Suite ........ : type révision...A cette occasion, transformation des cases à
Suite ........ : cocher en THMultiValComboBox
Suite ........ : 
Suite ........ : FQ 14314 : CA  18/06/2007 : Message explicite si balance 
Suite ........ : non créée
Mots clefs ... : TOF;BALSITCREATION
*****************************************************************}
Unit uTofBalSitCreation;

Interface

Uses Windows, StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
     HCtrls, HEnt1, HMsgBox, HTB97, UTOF,
     Ent1, CritEdt,
     BalSit
{$IFDEF EAGLCLIENT}
      , MaineAGL
      , uTOB
{$ELSE}
     {$IFNDEF DBXPRESS},dbtables {$ELSE} ,uDbxDataSet {$ENDIF}
     , db
     ,fe_main
{$ENDIF}
     ;

Type
  TOF_BALSITCREATION = Class (TOF)

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
      FAuto   : boolean;
      FDebEx  : TDateTime;
      FFinEx  : TDateTime;
      procedure OnEcranKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
      procedure OnChangeEtablissement(Sender: TObject) ;
      procedure OnChangeDevise(Sender: TObject) ;
      procedure OnChangeExercice(Sender: TObject) ;
      procedure OnDate1Exit(Sender: TObject) ;
      procedure OnChangeCompte1(Sender: TObject) ;
      procedure OnChangeCompte3(Sender: TObject) ;
      procedure OnChangeAxe(Sender: TObject) ;
      procedure OnChangeTypeCumul(Sender: TObject) ;
      procedure OnChangeTypeBalance(Sender : TObject) ;
      procedure OnClickCumulMois(Sender: TObject) ;
      procedure OnClickCumulPeriode(Sender: TObject) ;
      procedure OnInfoBalSit (Sender : TObject; Msg : string );
      procedure OnAutomatiqueClick(Sender: TObject) ;
      procedure InitZonesSaisie ( bAuto : boolean );
      function  DecodeQualifPiece : String ;
  end ;

procedure CPLanceFiche_BALSIT_CREATION ;

Implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  CPTypeCons,
  ULibExercice,
  {$ENDIF MODENT1}
  Paramsoc; // GetParamSocSecur



////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/07/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_BALSIT_CREATION ;
begin
  AGLLanceFiche('CP', 'BALSIT_CREATION', '', '', '');
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_BALSITCREATION.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BALSITCREATION.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BALSITCREATION.OnUpdate ;
var BS             : TBalSit;
    InfoBS         : RBalSitInfo;
    lStQualP       : String ;
begin
  Inherited ;
  BS := TBalSit.Create(GetControlText('CODE'));

  // GCO - 23/06/2006 - FQ 18467
  (*
  if not VH^.TenueEuro then
  begin
    MessageAlerte('Le dossier doit être en Euro pour créer une balance de situation.');
    BS.Free;
    SetFocusControl('CODE') ;
    exit;
  end else *)
  // FIN GCO

  if BS.CodeBal = '' then
  begin
    PGIInfo('Veuillez renseigner le code correctement.',ECRAN.Caption);
    BS.Free;
    SetFocusControl('CODE') ;
    exit;
  end
  else
  if Copy(BS.CodeBal, 1, 4) = 'BDS@' then
  begin
    SetFocusControl('CODE') ;
    PgiInfo('Vous ne pouvez pas créer de balance de situation commençant par BDS@.', Ecran.Caption);
    Exit;
  end
  else
  if BS.Existe then
  begin
    PGIInfo ('Attention, ce code existe déjà !'#10#13'Veuillez en choisir un autre.',ECRAN.Caption);
    BS.Free;
    SetFocusControl('CODE') ;
    exit;
  end else
  if  THEdit(GetControl('LIBELLE')).Text='' then
  begin
    PGIInfo('Veuillez renseigner le libellé correctement.',ECRAN.Caption);
    BS.Free;
    SetFocusControl('LIBELLE') ;
    exit;
  end else
  { FQ 16188 - Contrôle de la cohérence des dates }
  if (StrToDate(THEdit(GetControl('DATE1')).Text) > StrToDate(THEdit(GetControl('DATE2')).Text)) then
  begin
    PGIInfo('Dates incohérentes.#10#13Veuillez saisir une date de début antérieure à la date de fin.',ECRAN.Caption);
    BS.Free;
    SetFocusControl('DATE1') ;
    exit;
  end;
  BS.OnInformation := OnInfoBalSit;
  SetControlVisible('AFFTRAITEMENT',True) ;
  SetControlCaption('AFFTRAITEMENT',TraduireMemoire('Traitement en cours, veuillez patienter ...')) ;
  InfoBS.Libelle := THEdit(GetControl('LIBELLE')).Text ;
  InfoBS.Abrege := THEdit(GetControl('ABREGE')).Text ;
  InfoBS.Plan := THValComboBox(GetControl('TYPECUM')).Value;
  InfoBS.Compte1Inf := THEdit(GetControl('COMPTE1')).Text ;
  InfoBS.Compte1Sup := THEdit(GetControl('COMPTE2')).Text;
  InfoBS.Compte2Inf := THEdit(GetControl('COMPTE3')).Text;
  InfoBS.Compte2Sup := THEdit(GetControl('COMPTE4')).Text;
  InfoBS.DateInf := StrToDate(THEdit(GetControl('DATE1')).Text);
  InfoBS.DateSup := StrToDate(THEdit(GetControl('DATE2')).Text);
  InfoBS.CumExo := GetCheckBoxState('CUMEXO')=cbChecked;
  InfoBS.Exo := THValComboBox(GetControl('EXERCICE')).Value;
  InfoBS.Ano := GetCheckBoxState('CUMANO')=cbChecked;
  InfoBS.CumMois := GetCheckBoxState('CUMMOIS')=cbChecked;
  InfoBS.CumPer := GetCheckBoxState('CUMPER')=cbChecked;
  InfoBS.Etab := THValComboBox(GetControl('ETABLISSEMENT')).Value;
  InfoBS.CumEtab := GetCheckBoxState('CUMETAB')=cbChecked;
  InfoBS.Devise := THValComboBox(GetControl('DEVISE')).Value;
  InfoBS.CumDev := GetCheckBoxState('CUMDEV')=cbChecked;
  InfoBS.TypeBalance := THValComboBox(GetControl('TYPEBAL')).Value;

  // FQ 17107 : SBO 03/01/2006 : Ajout des écritures de révision...
  lStQualP := DecodeQualifPiece ;
  InfoBS.TypeEcr[Reel]  := pos( 'N', lStQualP ) > 0 ;
  InfoBS.TypeEcr[Simu]  := pos( 'S', lStQualP ) > 0 ;
  InfoBS.TypeEcr[Situ]  := pos( 'U', lStQualP ) > 0 ;
  InfoBS.TypeEcr[Previ] := pos( 'P', lStQualP ) > 0 ;
  InfoBS.TypeEcr[Revi]  := pos( 'R', lStQualP ) > 0 ;
  InfoBS.TypeEcr[Ifrs]  := pos( 'I', lStQualP ) > 0 ;

  InfoBS.Axe := THValComboBox(GetControl('AXE')).Value;
  InfoBS.DebCreColl := GetCheckBoxState('FSOLDEDEBCRECOLL')=cbChecked;
  if (TToolBarButton97(GetControl('FAUTOMATIQUE')).Tag = 0) then
  begin
    try
      BS.GenereAuto ( InfoBS );
      // CA - 18/06/2007 - FQ 14314 : Message explicite si balance non créée
      if ExisteSQL ('SELECT BSI_CODEBAL FROM CBALSIT WHERE BSI_CODEBAL="'+BS.CodeBal+'"') then
        PGIInfo('Traitement terminé.',ECRAN.Caption)
      else PGIInfo ('Aucun enregistrement. La balance n''a pas été générée.');
    finally
      BS.Free;
    end;
  end
  else
  begin
    BS.GenereManu (InfoBS );
    BS.Free;
    if PGIAsk ('Préparation de la balance terminée.#10#13Voulez-vous passer en saisie ?',ECRAN.Caption)=mrYes then
      AGLLanceFiche ('CP','BALSIT_MODIF','','',GetControlText('CODE'));
  end;
  Ecran.Close;
end ;

procedure TOF_BALSITCREATION.OnLoad ;
begin
  Inherited ;

  // GCO - 23/06/2006 - FQ 18467
  (*
  if not VH^.TenueEuro then
  begin
    MessageAlerte('Le dossier doit être en Euro pour créer une balance de situation.');
  end;*)

  // GCO - 05/01/2005 - FQ 15188
  PositionneEtabUser(THValComboBox(GetControl('ETABLISSEMENT')));
end ;

procedure TOF_BALSITCREATION.OnArgument (S : String ) ;
begin
  Inherited ;
  FAuto := True;
  ECRAN.OnKeyDown := OnEcranKeyDown;
  TToolBarButton97(GetControl('FAUTOMATIQUE')).OnClick := OnAutomatiqueClick ;
  THValComboBox(GetControl('ETABLISSEMENT')).OnChange := OnChangeEtablissement;
  THValComboBox(GetControl('DEVISE')).OnChange := OnChangeDevise;
  THValComboBox(GetControl('EXERCICE')).OnChange := OnChangeExercice;
  THEdit(GetControl('DATE1')).OnExit := OnDate1Exit;
  THCritMaskEdit(GetControl('COMPTE1')).OnChange := OnChangeCompte1;
  THCritMaskEdit(GetControl('COMPTE3')).OnChange := OnChangeCompte3;
  THValComboBox(GetControl('AXE')).OnChange := OnChangeAxe;
  THValComboBox(GetControl('TYPECUM')).OnChange := OnChangeTypeCumul;
  TCheckBox(GetControl('CUMMOIS')).OnClick := OnClickCumulMois;
  TCheckBox(GetControl('CUMPERIODE')).OnClick := OnClickCumulPeriode;
  InitZonesSaisie ( True ); // On arrive sur la fenêtre en mode auto. sytématiquement
  SetControlChecked('CUMANO',True);
  SetControlVisible('AFFTRAITEMENT',FALSE) ;
  OnChangeCompte1(nil);
  OnChangeCompte3(nil);
  THValComboBox(GetControl('EXERCICE')).Value := VH^.Entree.Code;
  OnChangeExercice (nil);
  OnDate1Exit(nil);
  THValComboBox(GetControl('DEVISE')).Value := V_PGI.DevisePivot;
  OnChangeDevise(nil);
  THValComboBox(GetControl('ETABLISSEMENT')).Value := VH^.EtablisDefaut;
  OnChangeEtablissement (nil);
  THValComboBox(GetControl('AXE')).ItemIndex := 0;
  OnChangeAxe(nil);
  THValComboBox(GetControl('TYPECUM')).Value := 'GEN';
  OnChangeTypeCumul (nil);

  // GCO - 10/01/2006
  THValComboBox(GetControl('TYPEBAL')).Value := 'BDS'; // Init par défaut = Statique
  THValComboBox(GetControl('TYPEBAL')).OnChange := OnChangeTypeBalance;
  if not GetParamSocSecur('SO_CPBDSDYNA', True) then
  begin
    SetControlVisible('TYPEBAL', False);
    SetControlVisible('TTYPEBAL', False);
  end;
  SetControlChecked('FSOLDEDEBCRECOLL',True);

  SetFocusControl('CODE');
end ;

procedure TOF_BALSITCREATION.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BALSITCREATION.OnChangeEtablissement(Sender: TObject);
begin
  SetControlEnabled('TCUMETAB',THValComboBox(GetControl('ETABLISSEMENT')).Value='');
  SetControlEnabled('CUMETAB',THValComboBox(GetControl('ETABLISSEMENT')).Value='') ;
  if THValComboBox(GetControl('ETABLISSEMENT')).Value<>'' then
    SetControlChecked('CUMETAB',FALSE)
end;

procedure TOF_BALSITCREATION.OnChangeDevise(Sender: TObject);
begin
  SetControlEnabled('TCUMDEV',THValComboBox(GetControl('DEVISE')).Value='');
  SetControlEnabled('CUMDEV',THValComboBox(GetControl('DEVISE')).Value='') ;
  if THValComboBox(GetControl('DEVISE')).Value<>'' then
    SetControlChecked('CUMDEV',FALSE)
end;

procedure TOF_BALSITCREATION.OnChangeExercice(Sender: TObject);
var   Q : TQuery ;
      DebExo, FinExo : String ;
      Exercice : THValComboBox;
begin
  Exercice := THValComboBox(GetControl('EXERCICE'));
  if Exercice.Value='' then
  begin
    Q:=OpenSQL('SELECT MIN(EX_DATEDEBUT) DEBEXO, MAX(EX_DATEFIN) FINEXO FROM EXERCICE',True) ;
    FDebEx := Q.FindField('DEBEXO').AsDateTime;
    FFinEx := Q.FindField('FINEXO').AsDateTime;
    DebExo:=DateToStr(FDebEx) ;
    FinExo:=DateToStr(FFinEx) ;
    Ferme(Q) ;
    SetControlText('DATE1',DebExo);
    SetControlText('DATE2',FinExo);
  end else
  begin
    ExoToDates(Exercice.Value, THEdit(GetControl('DATE1')), THEdit(GetControl('DATE2')));
    FDebEx := StrToDate(THEdit(GetControl('DATE1')).Text);
    FFinEx := StrToDate(THEdit(GetControl('DATE2')).Text);
  end;
  if FAuto then
  begin
    SetControlEnabled('TCUMEXO',Exercice.Value='') ;
    SetControlEnabled('CUMEXO',Exercice.Value='') ;
    SetControlChecked('CUMEXO',FALSE) ;
  end;
  OnDate1Exit(nil);
end;

procedure TOF_BALSITCREATION.OnChangeCompte1(Sender: TObject);
begin
  SetControlVisible('TCOMPTE2',Pos('*',THEdit(GetControl('COMPTE1')).text)<=0) ;
  SetControlVisible('COMPTE2',Pos('*',THEdit(GetControl('COMPTE1')).text)<=0) ;
end;

procedure TOF_BALSITCREATION.OnChangeCompte3(Sender: TObject);
begin
  SetControlVisible('TCOMPTE4',Pos('*',THEdit(GetControl('COMPTE3')).text)<=0) ;
  SetControlVisible('COMPTE4',Pos('*',THEdit(GetControl('COMPTE3')).text)<=0) ;
end;

procedure TOF_BALSITCREATION.OnChangeAxe(Sender: TObject);
begin
  OnChangeTypeCumul (nil);
end;

procedure TOF_BALSITCREATION.OnChangeTypeCumul(Sender: TObject);
var AxeG, AxeS : string;
  TypeCum : string;
begin
  SetControlText('COMPTE1','') ; SetControlText('COMPTE2','') ;
  SetControlText('COMPTE3','') ; SetControlText('COMPTE4','') ;

  case THValComboBox(GetControl('AXE')).ItemIndex of
    0 :  begin AxeG:='1' ; AxeS:=''  ; end ;
    1 :  begin AxeG:='2' ; AxeS:='2' ; end ;
    2 :  begin AxeG:='3' ; AxeS:='3' ; end ;
    3 :  begin AxeG:='4' ; AxeS:='4' ; end ;
    4 :  begin AxeG:='4' ; AxeS:='5' ; end ;
  end;
  SetControlEnabled('FSOLDEDEBCRECOLL',False);
  SetControlChecked('FSOLDEDEBCRECOLL',False);
  TypeCum := THValComboBox(GetControl('TYPECUM')).Value;
  if TypeCum='GEN' then
  begin
    SetControlCaption('TCOMPTE1','Généraux de') ;
    SetControlVisible('TCOMPTE3',FALSE) ; SetControlVisible('TCOMPTE4',FALSE) ;
    SetControlVisible('COMPTE3',FALSE)  ; SetControlVisible('COMPTE4',FALSE)  ;
    SetControlVisible('AXE',FALSE)      ; SetControlVisible('TAXE',FALSE) ;
    SetControlProperty('COMPTE1','DATATYPE','TZGENERAL') ;
    SetControlProperty('COMPTE2','DATATYPE','TZGENERAL') ;
    SetControlEnabled('FSOLDEDEBCRECOLL',FAuto=True);
    THEdit(GetControl('COMPTE1')).MaxLength := VH^.Cpta[fbGene].Lg;
    THEdit(GetControl('COMPTE2')).MaxLength := VH^.Cpta[fbGene].Lg;
  end else if TypeCum='ANA' then
  begin
    SetControlCaption('TCOMPTE1','Sections de') ;
    SetControlVisible('TCOMPTE3',FALSE) ; SetControlVisible('TCOMPTE4',FALSE) ;
    SetControlVisible('COMPTE3',FALSE)  ; SetControlVisible('COMPTE4',FALSE)  ;
    SetControlVisible('AXE',TRUE )      ; SetControlVisible('TAXE',TRUE) ;
    SetControlProperty('COMPTE1','DATATYPE','TZSECTION'+AxeS) ;
    SetControlProperty('COMPTE2','DATATYPE','TZSECTION'+AxeS) ;
  end  else if TypeCum='TIE' then
  begin
    SetControlCaption('TCOMPTE1','Tiers de') ;
    SetControlVisible('TCOMPTE3',FALSE) ; SetControlVisible('TCOMPTE4',FALSE) ;
    SetControlVisible('COMPTE3',FALSE)  ; SetControlVisible('COMPTE4',FALSE)  ;
    SetControlVisible('AXE',FALSE)      ; SetControlVisible('TAXE',FALSE) ;
    SetControlProperty('COMPTE1','DATATYPE','TZTTOUS') ;
    SetControlProperty('COMPTE2','DATATYPE','TZTTOUS');
    THEdit(GetControl('COMPTE1')).MaxLength := VH^.Cpta[fbAux].Lg;
    THEdit(GetControl('COMPTE2')).MaxLength := VH^.Cpta[fbAux].Lg;
  end else if TypeCum='G/A' then
  begin
    SetControlCaption('TCOMPTE1','Généraux de') ;
    SetControlCaption('TCOMPTE3','Sections de') ;
    SetControlVisible('TCOMPTE3',TRUE) ; SetControlVisible('TCOMPTE4',TRUE) ;
    SetControlVisible('COMPTE3',TRUE)  ; SetControlVisible('COMPTE4',TRUE)  ;
    SetControlVisible('AXE',TRUE )      ; SetControlVisible('TAXE',TRUE) ;
    SetControlProperty('COMPTE1','DATATYPE','TZGVENTIL'+AxeG) ;
    SetControlProperty('COMPTE2','DATATYPE','TZGVENTIL'+AxeG) ;
    SetControlProperty('COMPTE3','DATATYPE','TZSECTION'+AxeS) ;
    SetControlProperty('COMPTE4','DATATYPE','TZSECTION'+AxeS) ;
    THEdit(GetControl('COMPTE1')).MaxLength := VH^.Cpta[fbGene].Lg;
    THEdit(GetControl('COMPTE2')).MaxLength := VH^.Cpta[fbGene].Lg;
  end else if TypeCum='G/T' then
  begin
    SetControlCaption('TCOMPTE1','Collectifs de') ; SetControlCaption('TCOMPTE3','Tiers de') ;
    SetControlVisible('TCOMPTE3',TRUE) ; SetControlVisible('TCOMPTE4',TRUE) ;
    SetControlVisible('COMPTE3',TRUE)  ; SetControlVisible('COMPTE4',TRUE)  ;
    SetControlVisible('AXE',FALSE )    ; SetControlVisible('TAXE',FALSE) ;
    SetControlProperty('COMPTE1','DATATYPE','TZGCOLLECTIF');
    SetControlProperty('COMPTE2','DATATYPE','TZGCOLLECTIF');
    SetControlProperty('COMPTE3','DATATYPE','TZTTOUS') ;
    SetControlProperty('COMPTE4','DATATYPE','TZTTOUS');
    THEdit(GetControl('COMPTE1')).MaxLength := VH^.Cpta[fbGene].Lg;
    THEdit(GetControl('COMPTE2')).MaxLength := VH^.Cpta[fbGene].Lg;
    THEdit(GetControl('COMPTE3')).MaxLength := VH^.Cpta[fbAux].Lg;
    THEdit(GetControl('COMPTE4')).MaxLength := VH^.Cpta[fbAux].Lg;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/06/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_BALSITCREATION.OnChangeTypeBalance(Sender: TObject);
begin
  if THValComBoBox(Sender).Value = 'DYN' then
  begin
    SetControlEnabled('FSOLDEDEBCRECOLL', False);
    SetControlChecked('FSOLDEDEBCRECOLL', False);
  end
  else
  begin
    SetControlEnabled('FSOLDEDEBCRECOLL', True);
    SetControlChecked('FSOLDEDEBCRECOLL', True);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_BALSITCREATION.OnInfoBalSit(Sender: TObject; Msg: string);
begin
  SetControlCaption('AFFTRAITEMENT',Msg);
  Application.ProcessMessages ;
end;

procedure TOF_BALSITCREATION.OnAutomatiqueClick(Sender: TObject);
var TBB : TToolBarButton97;
begin
  TBB := TToolBarButton97(GetControl('FAUTOMATIQUE'));
  if TBB.Tag = 0 then
  begin
    FAuto := False;
    TBB.Tag := 1;
    TBB.Caption := TraduireMemoire ('Manuelle');
    InitZonesSaisie ( False );
  end else
  begin
    FAuto := True;
    TBB.Tag := 0;
    TBB.Caption := TraduireMemoire ('Automatique');
    InitZonesSaisie ( True );
  end;
  SetControlEnabled('FSOLDEDEBCRECOLL',FAuto);
  SetControlVisible('FSOLDEDEBCRECOLL',FAuto);
  if not FAuto then SetControlChecked('FSOLDEDEBCRECOLL',False);
end;

procedure TOF_BALSITCREATION.OnClickCumulMois(Sender: TObject);
var PeriodeExiste : boolean;
begin
  PeriodeExiste := ExisteSQL('SELECT * FROM CPPERIODE WHERE CPE_NATURE="CPT"') and (not (GetCheckBoxState('CUMMOIS')=cbChecked));
  if PeriodeExiste then
  begin
    SetControlEnabled('CUMPERIODE',not (GetCheckBoxState('CUMMOIS')=cbChecked));
    SetControlEnabled('TCUMPER',not (GetCheckBoxState('CUMMOIS')=cbChecked));
    SetControlEnabled('CUMMOIS',not (GetCheckBoxState('CUMPERIODE')=cbChecked));
    SetControlEnabled('TCUMMOIS',not (GetCheckBoxState('CUMPERIODE')=cbChecked));
  end;
end;

procedure TOF_BALSITCREATION.OnClickCumulPeriode(Sender: TObject);
var PeriodeExiste : boolean;
begin
  PeriodeExiste := ExisteSQL('SELECT * FROM CPPERIODE WHERE CPE_NATURE="CPT"') and (not (GetCheckBoxState('CUMMOIS')=cbChecked));
  if PeriodeExiste then
  begin
    SetControlEnabled('CUMMOIS',not (GetCheckBoxState('CUMPERIODE')=cbChecked));
    SetControlEnabled('TCUMMOIS',not (GetCheckBoxState('CUMPERIODE')=cbChecked));
    SetControlEnabled('CUMPERIODE',not (GetCheckBoxState('CUMMOIS')=cbChecked));
    SetControlEnabled('TCUMPER',not (GetCheckBoxState('CUMMOIS')=cbChecked));
  end;
end;

procedure TOF_BALSITCREATION.OnEcranKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if ( csDestroying in Ecran.ComponentState ) then Exit ;

  case Key of
{ALT+M}
    77 : if Shift=[ssAlt]  then OnAutomatiqueClick(nil);
    VK_F10 :
      begin
        TToolBarButton97(GetControl('BVALIDER')).Click;
        Key := 0;
      end;
    VK_ESCAPE :
      begin
        TToolBarButton97(GetControl('BFERME')).Click;
        Key := 0;
      end;
  end;
end;

procedure TOF_BALSITCREATION.InitZonesSaisie(bAuto: boolean);
var PeriodeExiste : boolean;
begin

  SetControlEnabled('QUALIFPIECE',bAuto);
{  SetControlEnabled('TYPECRN',bAuto);
  SetControlEnabled('TYPECRS',bAuto);
  SetControlEnabled('TYPECRU',bAuto);
  SetControlEnabled('TYPECRP',bAuto);
  SetControlEnabled('TYPECRI',bAuto);
}
  SetControlEnabled('CUMANO',bAuto);
  SetControlEnabled('TYPECR',bAuto);
  SetControlEnabled ('CUMEXO',bAuto);
  SetControlEnabled ('TCUMEXO',bAuto);
  SetControlEnabled ('CUMMOIS',bAuto);
  SetControlEnabled ('TCUMMOIS',bAuto);
  SetControlEnabled ('CUMPERIODE',bAuto);
  SetControlEnabled ('TCUMPERIODE',bAuto);
  SetControlEnabled ('CUMETAB',bAuto);
  SetControlEnabled ('TCUMETAB',bAuto);
  SetControlEnabled ('CUMDEV',bAuto);
  SetControlEnabled ('TCUMDEV',bAuto);
  SetControlEnabled('CUMPERIODE',False);
  SetControlEnabled('TCUMPERIODE',False);
  if bAuto then
  begin
    SetControlProperty('EXERCICE','VideString',TraduireMemoire('<<Tous>>'));
    THValComboBox(GetControl('EXERCICE')).ReLoad;
    PeriodeExiste := ExisteSQL('SELECT * FROM CPPERIODE WHERE CPE_NATURE="CPT"') and (not (GetCheckBoxState('CUMMOIS')=cbChecked));
    SetControlEnabled('CUMPERIODE',PeriodeExiste);
    SetControlEnabled('TCUMPERIODE',PeriodeExiste);
    THValComboBox(GetControl('EXERCICE')).Value := VH^.Entree.Code;
  end else
  begin
    SetControlProperty('EXERCICE','VideString',TraduireMemoire('<<Autre>>'));
    THValComboBox(GetControl('EXERCICE')).ReLoad;
    THValComboBox(GetControl('EXERCICE')).Value := VH^.Entree.Code;
  end;
end;

procedure TOF_BALSITCREATION.OnDate1Exit(Sender: TObject);
var Date1 : TDateTime;
begin
  if IsValidDate(THEdit(GetControl('DATE1')).Text) then
  begin
    Date1 := StrToDate(THEdit(GetControl('DATE1')).Text);
    if THValComboBox(GetControl('EXERCICE')).ItemIndex<>0  then// Tous non sélectionné
    begin
      if FDebEx < Date1 then  // La période prise en compte ne comprend pas la date de début d'exercice
      begin
        SetControlChecked('CUMANO',False);
        SetControlEnabled('CUMANO',False);
      end else SetControlEnabled('CUMANO',FAuto);
    end else SetControlEnabled('CUMANO',FAuto);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 03/01/2006
Modifié le ... :   /  /
Description .. : FQ 17107 : SBO 03/01/2006 : Ajout des écritures de
Suite ........ : révision...A cette occasion, transformation des cases à
Suite ........ : cocher en THMultiValComboBox
Suite ........ :
Suite ........ : Retroune la liste des Qualifpiece sélectionnées séparées
Suite ........ : par des ";"
Mots clefs ... : 
*****************************************************************}
function TOF_BALSITCREATION.DecodeQualifPiece: String;
var lMultiVal : THMultiValComboBox ;
begin

  lMultiVal := THMultiValComboBox( GetControl('QUALIFPIECE', True) ) ;

  if lMultiVal.Tous or ( lMultiVal.Text = '' )
                    or ( lMultiVal.Text = TraduireMemoire('<<Tous>>') ) then
    begin

    result := 'N;S;R;' ;

    if not (ctxPCL in V_PGI.PGIContexte) then
      begin
      // Coche pour écriture Prévision visible uniquement sur S7 en PGE
      if EstSerie(S7) then
        result := result + 'P;' ;
      // Coche pour écriture situation non visible en S3 PGE
      if not EstSerie(S3) then
        result := result + 'U;' ;
      end
    // PCL on prend tout
    else result := result + 'P;U;' ;

    // IAS / IFRS
    if EstComptaIFRS then
      result := result + 'I;' ;

    end
  else
    result := lMultiVal.Text ;

end;

Initialization
  registerclasses ( [ TOF_BALSITCREATION ] ) ;
end.
