{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 05/03/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : EXTOURNE_PARAM ()
Mots clefs ... : TOF;EXTOURNE_PARAM
*****************************************************************}
Unit uTofExtourneParam ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     uTOB,
     HTB97,
     AGLInit,
     Ent1,
     SaisUtil,
     ParamSoc;

Type
  TOF_EXTOURNE_PARAM = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
      procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
      function ControleOK : boolean;
      function DateOk     : boolean;
      procedure RefOrigChange(Sender: TObject);
  end ;

Implementation

procedure TOF_EXTOURNE_PARAM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_EXTOURNE_PARAM.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_EXTOURNE_PARAM.OnUpdate ;
var T : TOB;
begin
  LastError := 0;
  if not ControleOK then
  begin
    T := nil;
    LastError := -1;
  end else
  begin
    T := TOB.Create ('', nil, -1);
    T.AddChampSupValeur('CONTRE_DATE',StrToDate(GetControlText('CONTRE_DATE')));
    T.AddChampSupValeur('CONTRE_TYPE',GetControlText('CONTRE_TYPE'));
    T.AddChampSupValeur('CONTRE_NEGATIF',GetControlText('CONTRE_NEGATIF'));
    T.AddChampSupValeur('CONTRE_LIBELLE',GetControlText('CONTRE_LIBELLE'));
    T.AddChampSupValeur('CONTRE_REFORIGINE',GetControlText('CONTRE_REFORIGINE'));
    T.AddChampSupValeur('CONTRE_REFINTERNE',GetControlText('CONTRE_REFINTERNE'));
    T.AddChampSupValeur('CONTRE_REFEXTERNE',GetControlText('CONTRE_REFEXTERNE'));
    T.AddChampSupValeur('CONTRE_JOURNAL',GetControlText('CONTRE_JOURNAL')); {YMO 09/03/07 FQ13822 Choix du Jal de géné}
  end;
  TheTOB := T;
  if LastError=0 then Ecran.Close;
end ;

procedure TOF_EXTOURNE_PARAM.OnLoad ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 09/03/2007
Modifié le ... :   /  /    
Description .. : FQ12700 Choix de la référence interne
Suite ........ :
Suite ........ : 1 Par défaut, la case est cochée : on reprend la ref d'origine
Suite ........ : 2 On décoche : on reprend la date:'Extourne JJ/MM/AAAA'
Suite ........ : 3 On décoche et on saisit une référence au choix
Suite ........ :
Mots clefs ... : 12700 REFINTERNE INTERNE ORIGINE
*****************************************************************}
procedure TOF_EXTOURNE_PARAM.RefOrigChange(Sender: TObject);
var SaisieLibre:Boolean;
begin
  SaisieLibre:=Not TCheckBox(GetControl('CONTRE_REFORIGINE')).Checked;
  SetControlEnabled('CONTRE_REFINTERNE', SaisieLibre);

  If SaisieLibre then ThEdit(GetControl('CONTRE_REFINTERNE')).SetFocus;
end;

procedure TOF_EXTOURNE_PARAM.OnArgument (S : String ) ;
begin
  Inherited ;
  Ecran.OnKeyDown := FormKeyDown;

  SetControlText('CONTRE_DATE', DateToStr(V_PGI.DateEntree));

  {YMO 09/03/07 FQ13822 Choix du Jal de génération}
  THValComboBox(GetControl('CONTRE_JOURNAL')).Plus:='J_NATUREJAL="'+LaTOB.GetString('NATUREJAL')
  +'" AND J_MODESAISIE="'+LaTOB.GetString('MODESAISIE')+'"'; {FQ13822 YMO 20.06.07}
  SetControlText('CONTRE_JOURNAL', LaTOB.GetValue('JOURNAL'));
  {YMO 09/03/07 FQ12700 Choix de la référence interne}
  TCheckBox(GetControl('CONTRE_REFORIGINE')).OnClick:=RefOrigChange;

  SetControlText('INFO_JOURNAL',LaTOB.GetValue('JOURNAL')+' - '+GetColonneSQL('JOURNAL','J_LIBELLE','J_JOURNAL="'+LaTOB.GetValue('JOURNAL')+'"'));
  SetControlText('INFO_LIGNES',IntToStr(LaTOB.GetValue('LIGNES'))+' '+
    TraduireMemoire('lignes d''écritures sélectionnées pour un montant de ')+
    StrFMontant(LaTOB.GetValue('MONTANT'),12,V_PGI.OkDecV,GetParamSocSecur('SO_DEVISEPRINC',''),True));
  SetControlEnabled('CONTRE_TYPE',LaTOB.GetValue('MODESAISIE')='-');
  if LaTOB.GetValue('MODESAISIE')<>'-' then SetControlText('CONTRE_TYPE','N')
  else SetControlText('CONTRE_TYPE',LaTOB.GetValue('QUALIFPIECE'));
  if not VH^.MontantNegatif then
  begin
    SetControlChecked('CONTRE_NEGATIF', False);
    SetControlEnabled('CONTRE_NEGATIF', False);
  end;
  TheTOB := nil;
end ;

procedure TOF_EXTOURNE_PARAM.OnClose ;
begin
  Inherited ;
end ;

function TOF_EXTOURNE_PARAM.ControleOK: boolean;
begin
  Result := False;
  if not DateOK then exit;
  if ((LaTOB.GetValue('QUALIFPIECE') <> 'N') and (GetControlText('CONTRE_TYPE')= 'N')) then
  begin
    MessageAlerte('Vous ne pouvez pas extourner en écriture courante des écritures de type '
      +
      RechDom('TTQUALIFPIECE',LaTOB.GetValue('QUALIFPIECE'),False));
    Exit;
  end;
  Result := True;
end;

function TOF_EXTOURNE_PARAM.DateOk: boolean;
var
  Err: integer;
begin
  Result := False;
  Err := ControleDate(THCritMaskEdit(GetControl('CONTRE_DATE')).Text);
  if ((Err > 0) and (Err < 5)) then
  begin
    case Err of
      1: MessageAlerte('La date que vous avez renseignée n''est pas valide');
      2:
        MessageAlerte('La date que vous avez renseignée n''est pas dans un exercice ouvert');
      3:
        MessageAlerte('La date que vous avez renseignée est antérieure à une clôture');
      4:
        MessageAlerte('La date que vous avez renseignée est antérieure à une clôture');
// Non pris en compte pour PCL - voir pour PGE
//      5:
//        MessageAlerte('La date que vous avez renseignée est en dehors des limites autorisées');
    end;
    exit;
  end;
  Result := True;
end;

procedure TOF_EXTOURNE_PARAM.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (csDestroying in Ecran.ComponentState) then
    Exit;
  if (key =  vk_valide) then
  begin
    Key:=0 ;
    TToolBarButton97(GetControl('BVALIDER')).Click;
  end;
end;

Initialization
  registerclasses ( [ TOF_EXTOURNE_PARAM ] ) ;
end.
