{-------------------------------------------------------------------------------------
    Version   |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 8.00.001.001  09/01/07  JP   Création de l'unité : Mul des mouvements bancaires
 8.00.001.012  20/04/07  JP   FQ TRESO 10437 : branchement des lookups sur les références de pointage
 8.00.001.025  19/07/07  JP   FQ 21113 : Ajout d'un argument dans l'appel de la TOM pour connaître
                              la fiche d'origine : FO_MULEEXBQLIG
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
 8.10.001.010  19/09/07  JP   FQ 21317 : Gestion des mouvements saisis manuellement lors de la suppression
--------------------------------------------------------------------------------------}
unit CPMULEEXBQLIG_TOF;

interface

uses
  Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  mul, FE_Main,
  {$ELSE}
  eMul, MaineAGL,
  {$ENDIF}
  {$IFDEF TRCONF}
  uLibConfidentialite,
  {$ELSE}
  UTOF,
  {$ENDIF TRCONF}
  Variants, Forms, SysUtils, HCtrls, HEnt1, HMsgBox, Menus;

type
  {$IFDEF TRCONF}
  TOF_CPMULEEXBQLIG = class (TOFCONF)
  {$ELSE}
  TOF_CPMULEEXBQLIG = class (TOF)
  {$ENDIF TRCONF}
    procedure OnArgument(S : string); override;
  private
    PopupMenu : TPopUpMenu; {04/08/04}

    procedure BInsertClick (Sender : TObject);
    procedure BDeleteClick (Sender : TObject);
    procedure ListDblClick (Sender : TObject);
    procedure SlctAllClick (Sender : TObject);
    procedure RefPoiOnClick(Sender : TObject); {FQ 10437}
  end;

procedure CPLanceFiche_MulEexBqLig(Range, Lequel, Arguments : string);


implementation

uses
  EEXBQLIG_TOM, AglInit, HTB97, Ent1, ULibPointage, LookUp, Commun, Constantes;

{---------------------------------------------------------------------------------------}
procedure CPLanceFiche_MulEexBqLig(Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('CP', 'CPMULEEXBQLIG', Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEEXBQLIG.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  cPlus : string;
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  inherited;
  Ecran.HelpContext := 50000128;

  {Boutons d'insertion / suppression}
  SetControlEnabled('BINSERT', ExJaiLeDroitConcept(TConcept(ccSaisMvtBqe), False));
  SetControlEnabled('BDELETE', ExJaiLeDroitConcept(TConcept(ccSaisMvtBqe), False));
  TToolbarButton97(GetControl('BINSERT')).OnClick := BInsertClick;
  TToolbarButton97(GetControl('BDELETE')).OnClick := BDeleteClick;

  {PopupMenus d'insertion / suppression}
  SetControlEnabled('DEL', ExJaiLeDroitConcept(TConcept(ccSaisMvtBqe), False));
  SetControlEnabled('INS', ExJaiLeDroitConcept(TConcept(ccSaisMvtBqe), False));
  PopupMenu := TPopUpMenu(GetControl('POPUPMENU'));
  PopupMenu.Items[0].OnClick := BInsertClick;
  PopupMenu.Items[1].OnClick := BDeleteClick;
  AddMenuPop(PopupMenu, '', '');

  TToolbarButton97(GetControl('BOUVRIR')).OnClick := ListDblClick;
  TFMul(Ecran).FListe.OnDblClick := ListDblClick;

  {CEL_GENERAL contient des valeurs différentes selon le type de pointage}
  if EstPointageSurTreso then begin
    SetControlProperty('CEL_GENERAL' , 'DATATYPE', 'TRBANQUECP');
    SetControlProperty('CEL_GENERAL_', 'DATATYPE', 'TRBANQUECP');
    TFMul(Ecran).SetDBListe('TRMULEEXBQLIG'); {08/08/07 : Confidentiel}
  end
  else if VH^.PointageJal then begin
    SetControlProperty('CEL_GENERAL' , 'DATATYPE', 'TZJBANQUE');
    SetControlProperty('CEL_GENERAL_', 'DATATYPE', 'TZJBANQUE');
    SetControlCaption('TCEL_GENERAL', TraduireMemoire('Journal              de'));
  end
  else begin
    SetControlProperty('CEL_GENERAL' , 'DATATYPE', 'TTBANQUECP');
    SetControlProperty('CEL_GENERAL_', 'DATATYPE', 'TTBANQUECP');
  end;

  (GetControl('CEL_REFPOINTAGE' ) as THEdit).OnElipsisClick := RefPoiOnClick; {FQ TRESO 10437}
  (GetControl('CEL_REFPOINTAGE_') as THEdit).OnElipsisClick := RefPoiOnClick; {FQ TRESO 10437}

  TFMul(Ecran).bSelectAll.Visible := True;
  TFMul(Ecran).bSelectAll.OnClick := SlctAllClick;

  {08/08/07 : Pas valable en Tréso multi dossiers : utiliser FiltreBanqueCp
  SetPlusBanqueCp(GetControl('CEL_GENERAL'));
  SetPlusBanqueCp(GetControl('CEL_GENERAL_'));}
  cPlus := FiltreBanqueCp(THEdit(GetControl('CEL_GENERAL')).DataType, tcb_Bancaire, '');
  THEdit(GetControl('CEL_GENERAL')).Plus := cPlus;
  THEdit(GetControl('CEL_GENERAL_')).Plus := cPlus;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEEXBQLIG.BInsertClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  CPLanceFiche_EEXBQLIG('', '', ActionToString(taCreatEnSerie) + ';' + FO_MULEEXBQLIG);
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEEXBQLIG.BDeleteClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n  : Integer;
  BO : Boolean;
  MO : Boolean;
begin
  if (TFMul(Ecran).FListe.nbSelected = 0) and not (TFMul(Ecran).FListe.AllSelected) then begin
    HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner une ligne.;W;O;O;O;', '', '');
    Exit;
  end;

  BO := False; {19/09/07 : FQ 21317}
  MO := False; {19/09/07 : FQ 21317}

  if PgiAsk(TraduireMemoire('Seuls les mouvements bancaires non pointés peuvent être supprimés') + #13 +
            TraduireMemoire('Souhaitez-vous poursuivre ?'), Ecran.Caption) = mrYes then begin

    {$IFNDEF EAGLCLIENT}
    TFMul(Ecran).Q.First;
    if TFMul(Ecran).FListe.AllSelected then
      while not TFMul(Ecran).Q.EOF do begin
        if (VarToDateTime(TFMul(Ecran).Q.FindField('CEL_DATEPOINTAGE').AsString) <= iDate1900) then begin
          {19/09/07 : FQ 21317 : Gestion des mouvements saisis manuellement s'ils appartiennent à une session équilibrée}
          if (VarToStr(TFMul(Ecran).Q.FindField('CEL_VALIDE').AsString) = 'X') and not BO and
             (VarToStr(TFMul(Ecran).Q.FindField('EE_AVANCEMENT').AsString) = 'X') then begin
            BO := True;
            MO := HShowMessage('0;' + Ecran.Caption + ';' +
                  TraduireMemoire('ATTENTION ! Des mouvements saisis manuellement appartiennent à une amplitude de pointage terminée.') + #13#13 +
                  TraduireMemoire('Êtes-vous certain de vouloir les supprimer ?') + ';W;YN;N;N;', '', '') = mrNo;
          end;
          if (VarToStr(TFMul(Ecran).Q.FindField('CEL_VALIDE').AsString) = 'X') and MO and
             (VarToStr(TFMul(Ecran).Q.FindField('EE_AVANCEMENT').AsString) = 'X') then Continue;

          {Suppression du contrat courant ...}
          ExecuteSQL('DELETE FROM EEXBQLIG WHERE CEL_GENERAL = "'  + TFMul(Ecran).Q.FindField('CEL_GENERAL'  ).AsString + '" AND ' +
                                                'CEL_NUMRELEVE = ' + TFMul(Ecran).Q.FindField('CEL_NUMRELEVE').AsString + ' AND ' +
                                                'CEL_NUMLIGNE  = ' + TFMul(Ecran).Q.FindField('CEL_NUMLIGNE' ).AsString);
        end;
        TFMul(Ecran).Q.Next;
      end
    else
    {$ENDIF}

    {On boucle sur la sélection}
    for n := 0 to TFMul(Ecran).FListe.nbSelected - 1 do begin
      TFMul(Ecran).FListe.GotoLeBookmark(n);
      if VarToDateTime(GetField('CEL_DATEPOINTAGE')) <= iDate1900 then begin
        {19/09/07 : FQ 21317 : Gestion des mouvements saisis manuellement s'ils appartiennent à une session équilibrée}
        if (VarToStr(TFMul(Ecran).Q.FindField('CEL_VALIDE').AsString) = 'X') and not BO and
           (VarToStr(TFMul(Ecran).Q.FindField('EE_AVANCEMENT').AsString) = 'X') then begin
          BO := True;
          MO := HShowMessage('0;' + Ecran.Caption + ';' +
                TraduireMemoire('ATTENTION ! Des mouvements saisis manuellement appartiennent à une amplitude de pointage terminée.') + #13#13 +
                TraduireMemoire('Êtes-vous certain de vouloir les supprimer ?') + ';W;YN;N;N;', '', '') = mrNo;
        end;

        if (VarToStr(TFMul(Ecran).Q.FindField('CEL_VALIDE').AsString) = 'X') and MO and
           (VarToStr(TFMul(Ecran).Q.FindField('EE_AVANCEMENT').AsString) = 'X') then Continue;

        {Suppression du contrat courant ...}
        ExecuteSQL('DELETE FROM EEXBQLIG WHERE CEL_GENERAL = "' + VarToStr(GetField('CEL_GENERAL')) + '" AND ' +
                                              'CEL_NUMRELEVE = ' + VarToStr(GetField('CEL_NUMRELEVE')) + ' AND ' +
                                              'CEL_NUMLIGNE  = ' + VarToStr(GetField('CEL_NUMLIGNE')));
      end;
    end;

    TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEEXBQLIG.ListDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  s : string;
  B : Boolean;
begin
  B := (VarToStr(GetField('CEL_VALIDE')) = 'X') and ExJaiLeDroitConcept(TConcept(ccSaisMvtBqe), False);
  s := VarToStr(GetField('CEL_GENERAL')) + ';' + VarToStr(GetField('CEL_NUMRELEVE')) + ';' + VarToStr(GetField('CEL_NUMLIGNE')) + ';';
  if B then CPLanceFiche_EEXBQLIG('', s, ActionToString(taModif) + ';' + FO_MULEEXBQLIG)
       else CPLanceFiche_EEXBQLIG('', s, ActionToString(taConsult) + ';' + FO_MULEEXBQLIG);
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;


{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEEXBQLIG.SlctAllClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Fiche : TFMul;
begin
  Fiche := TFMul(Ecran);
  {$IFDEF EAGLCLIENT}
  if not Fiche.FListe.AllSelected then begin
    if not Fiche.FetchLesTous then Exit;
  end;
  {$ENDIF}
  Fiche.bSelectAllClick(nil);
end;

{FQ TRESO 10437 : branchement des lookups
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEEXBQLIG.RefPoiOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  wh : string;
begin
  wh := '';
  
  if GetControlText('CEL_GENERAL') <> '' then
    wh := 'EE_GENERAL >= "' + GetControlText('CEL_GENERAL') + '"';

  if GetControlText('CEL_GENERAL_') <> '' then begin
    if wh = '' then
      wh := 'EE_GENERAL <= "' + GetControlText('CEL_GENERAL_') + '"'
    else
      wh := wh + ' AND EE_GENERAL <= "' + GetControlText('CEL_GENERAL_') +  '"';
  end;

  LookUpList(THEdit(Sender), TraduireMemoire('Date de pointage'), 'EEXBQ', 'EE_REFPOINTAGE',
      'EE_DATEPOINTAGE', wh, 'EE_DATEPOINTAGE DESC', True, 0);

end;

initialization
  RegisterClasses([TOF_CPMULEEXBQLIG]);

end.

