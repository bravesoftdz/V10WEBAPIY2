{ Unité : Source TOF de la FICHE : TRMODIFIERUBRIQUE
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 6.00.018.001    11/10/04   JP     Création de l'unité FQ 10175
 7.05.001.001    04/10/06   JP     FQ 10364 : Gestion du BSelectAll
--------------------------------------------------------------------------------------}
unit TRMODIFIERUBRIQUE_TOF;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, Mul, FE_Main, HDB,
  {$ELSE}
  eMul, MaineAGL,
  {$ENDIF}
  Forms, sysutils, HCtrls, HEnt1, HMsgBox, UTOF, HTB97;


type
  TOF_TRMODIFIERUBRIQUE = class (TOF)
    procedure OnArgument(S : string); override;
  private
    procedure LancerRubrique(Sender : TObject);
    procedure ListeDbClick  (Sender : TObject);
    procedure SlctAllClick  (Sender : TObject); {FQ 10364}
  end ;

procedure TRLanceFiche_ModifieRubrique(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  Constantes, UObjGen;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_ModifieRubrique(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMODIFIERUBRIQUE.OnArgument(S : string );
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 500003;
  {On ne traite que les écritures venant de la comptabilité ou les écritures de simulations}
  SetControlText('XX_WHERE', '(TE_QUALIFORIGINE = "' + QUALIFCOMPTA + '" ) OR (TE_QUALIFORIGINE = "' + QUALIFTRESO + '" AND TE_NATURE = "' + na_Simulation + '")');

  {FQ 10364 : 04/10/06 : Branchement du AllSelect en eAgl}
  SetControlVisible('BSELECTALL', True);
  TToolbarButton97(GetControl('BSELECTALL')).OnClick := SlctAllClick;

  TFMul(Ecran).FListe.OnDblClick := ListeDbClick;
  TToolbarButton97(GetControl('BOUVRIR' )).OnClick := ListeDbClick;
  TToolbarButton97(GetControl('BVALIDER')).OnClick := LancerRubrique;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMODIFIERUBRIQUE.ListeDbClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  {$IFDEF EAGLCLIENT}
  if TFMul(Ecran).FListe.RowCount = 0 then Exit;
  {$ELSE}
  if TFMul(Ecran).Q.Eof and TFMul(Ecran).Q.Bof then Exit;
  {$ENDIF}
  s := GetField('TE_NODOSSIER') + ';' + GetField('TE_NUMTRANSAC') + ';' +
       VarToStr(GetField('TE_NUMEROPIECE')) + ';' + VarToStr(GetField('TE_NUMLIGNE'));
  AGLLanceFiche('TR', 'TRFICECRITURE', '', s, GetField('TE_NATURE') + ';');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMODIFIERUBRIQUE.LancerRubrique(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  {$IFDEF EAGLCLIENT}
  F   : THGrid;
  {$ELSE}
  F   : THDBGrid;
  {$ENDIF}
  Nb  : Integer;
  n   : Integer;
  s   : string;
  Obj : TObjRubrique;
  CIB : TObjCIBModePaie;

    {----------------------------------------------------------------------}
    procedure ModifierRubrique;
    {----------------------------------------------------------------------}
    var
      Gen,
      Rub,
      Lib,
      Pai : string;
    begin
      try
        Gen := GetField('TE_CONTREPARTIETR');
        if Valeur(VarToStr(GetField('TE_MONTANT'))) >= 0 then begin
          Pai := CIB.GetModePaie(GetField('TE_GENERAL'), GetField('TE_CODECIB'), 'ENC');
          Obj.GetCorrespondance(Gen, Rub, Lib, 'C', Pai);
        end
        else begin
          Pai := CIB.GetModePaie(GetField('TE_GENERAL'), GetField('TE_CODECIB'), 'DEC');
          Obj.GetCorrespondance(Gen, Rub, Lib, 'D', Pai);
        end;

        {On ne met à jour le code rubrique que si on en a trouvé un autre}
        if Rub = '' then Exit;

        Lib := 'TE_NUMTRANSAC = "' + GetField('TE_NUMTRANSAC') + '" AND ' +
               'TE_NUMEROPIECE = ' + VarToStr(GetField('TE_NUMEROPIECE'))  + ' AND ' +
               'TE_NUMLIGNE = ' + VarToStr(GetField('TE_NUMLIGNE')) + ' AND ' +
               'TE_EXERCICE = "' + GetField('TE_EXERCICE') + '"';
        ExecuteSQL('UPDATE TRECRITURE SET TE_CODEFLUX = "' + Rub + '" WHERE ' + Lib);
        Inc(Nb);
      except
        on E : Exception do raise;
      end;
    end;

begin
  {$IFDEF EAGLCLIENT}
  F := THGrid(TFMul(Ecran).FListe);
  {$ELSE}
  F := THDBGrid(TFMul(Ecran).FListe);
  {$ENDIF}

  Nb := 0;

  {Aucune sélection, on sort}
  if (F.NbSelected = 0)
  {$IFNDEF EAGLCLIENT}
  and not F.AllSelected {20/09/04 FQ 10138}
  {$ENDIF}
  then begin
    HShowMessage('3;' + Ecran.Caption + ';Aucun élément n''est sélectionné.;W;O;O;O;', '', '');
    Exit;
  end;

  if HShowMessage('2;' + Ecran.Caption + ';Êtes-vous sûr de vouloir mettre à jour les rubriques;Q;YNC;N;C;', '', '') <> mrYes then Exit;

  Obj := TObjRubrique.Create;
  Obj.Generaux.LoadDetailFromSQL('SELECT DISTINCT TE_CONTREPARTIETR GENERAL FROM TRECRITURE');
  CIB := TObjCIBModePaie.Create;
  try
    {Construction de la liste de correspondance Comptes généraux / Rubriques}
    Obj.SetListeCorrespondance;

    BeginTrans;
    try
      {$IFNDEF EAGLCLIENT}
      TFMul(Ecran).Q.First;
      if F.AllSelected then
        while not TFMul(Ecran).Q.EOF do begin
          ModifierRubrique;
          TFMul(Ecran).Q.Next;
        end
      else
      {$ENDIF}

      for n := 0 to F.nbSelected - 1 do begin
        F.GotoLeBookmark(n);
        {$IFDEF EAGLCLIENT}
        TFMul(Ecran).Q.TQ.Seek(F.Row - 1);
        {$ENDIF}
        ModifierRubrique;
      end;
      CommitTrans;
      if Nb = 0 then s := TraduireMemoire('Aucune écriture n''a été traitée.')
      else if Nb = 1 then s := TraduireMemoire('Une écriture a été traitée.')
      else s := TraduireMemoire(IntToStr(Nb) + ' écritures ont été traitées.');
      HShowMessage('1;' + Ecran.Caption + ';Le traitement s''est correctement effectué.'#13 + s + ';I;O;O;O;', '', '');
    except
      on E : Exception do begin
        RollBack;
        HShowMessage('0;' + Ecran.Caption + '; Traitement interrompu :'#13 + E.Message + ';E;O;O;O;', '', '');
      end;
    end;
  finally
    FreeAndNil(Obj);
    FreeAndNil(Cib);
  end;
  {Raffraîchissement de la liste}
  TToolbarButton97(GetControl('BCHERCHE')).Click;
end;

{FQ 10364 : Branchement du AllSelect en eAgl
{---------------------------------------------------------------------------------------}
procedure TOF_TRMODIFIERUBRIQUE.SlctAllClick(Sender: TObject);
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



initialization
  RegisterClasses([TOF_TRMODIFIERUBRIQUE]);
end.
