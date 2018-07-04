{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 10/08/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGRETROACTIV ()
Mots clefs ... : TOF;PGRETROACTIV
*****************************************************************
PT1  02/07/2007 FC V_72 : FQ 14507  Habilitation et rétroactivité: pb de traitement.
PT2  13/10/2008 JS FQ n°15351 ajout bouton [Ouvrir] pour accéder à la saisie de la rétroactivité
}
Unit UtofPGRETROACTIV;

Interface

Uses
{$IFDEF VER150}
     Variants,
{$ENDIF}
     StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     uTob,
     forms,
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,                       //TToolbarButton97;
     UtofPGRETROSIMULCOT,         //PgRetroSimulcot_LanceFiche    (simulation cotisation)
     UtomPGRETROACTIV,            //PgRetroactiv_LanceFiche       (saisie cotisation/rémunération)
     UtofPGRETROSIMULREM,         //PgRetroSimulrem_LanceFiche    (simulation rémunération)
     UTOF
      ;

Type
  TOF_PGRETROACTIV = Class (TOF)
    private
      procedure btInsertCOTOnClick;
      procedure btInsertREMOnClick;
      procedure btinsertOnClick ( Sender : Tobject );
      procedure LalisteOnDblClick ( Sender : Tobject);
      procedure btSimulOnClick ( Sender : TObject );
      procedure ChangeNature ( Sender : Tobject );

    public
      BTOPEN            : TToolbarButton97;//PT2
      btSimul           : TToolbarButton97;
      btInsert          : TToolbarButton97;
      Laliste           : THGrid;
      Combonature       : THValComboBox;

      procedure OnArgument (Arguments : String ) ; override ;
      procedure OnLoad                   ; override ;
      procedure OnDisplay                ; override ;
      procedure OnNew                    ; override ;
      procedure OnDelete                 ; override ;
      procedure OnUpdate                 ; override ;
      procedure OnClose                  ; override ;
      procedure OnCancel                 ; override ;
  end ;

Implementation

procedure TOF_PGRETROACTIV.OnArgument (Arguments : String ) ;
begin
  Inherited ;
//  V_PGI.SAV             := True;   //PT1

  SetControlVisible('BINSERT', True);
 //deb PT2
  SetControlVisible('BOuvrir', True);
  BTOPEN := TToolbarButton97(GetControl('BOuvrir'));
  BTOPEN.OnClick := LalisteOnDblClick;
  //fin PT2

  Laliste               := THGrid(GetControl('FLISTE'));
  Laliste.OnDblClick    := LalisteOnDblClick;

  btInsert              := TToolbarButton97(GetControl('BINSERT'));
  btInsert.OnClick      := btinsertOnClick;

  btSimul               := TToolbarButton97(GetControl('BSIMUL'));
  btSimul.OnClick       := btSimulOnClick;

  SetControlText('PGR_NATURERUB', 'COT');

  Combonature           := THValComboBox(GetControl('PGR_NATURERUB'));
  Combonature.OnChange  := ChangeNature;



end ;

procedure TOF_PGRETROACTIV.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGRETROACTIV.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGRETROACTIV.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGRETROACTIV.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGRETROACTIV.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGRETROACTIV.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGRETROACTIV.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PGRETROACTIV.LalisteOnDblClick(Sender: Tobject);
var
  FicheRetro            : String;
  StrArg                : String;

begin
  if GetField('PGR_PREDEFINI') = null then
    exit;

  FicheRetro            := GetField('PGR_PREDEFINI') + ';' + GetField('PGR_NODOSSIER') + ';' +
                          GetField('PGR_NATURERUB') + ';' + GetField('PGR_RUBRIQUE') + ';' +
                          DateToText(GetField('PGR_DATEVALIDITE'));

  StrArg                := 'MONOFICHE;ACTION=MODIFICATION;' + Getfield('PGR_NATURERUB');

  PgRetroactiv_LanceFiche(StrArg, FicheRetro);
  TFMul(ecran).BChercheClick(Ecran);

end;

procedure TOF_PGRETROACTIV.btInsertCOTOnClick;
var
  StrArg                : String;

begin
  StrArg                := 'MONOFICHE;ACTION=CREATION;COT';
  PgRetroactiv_LanceFiche(StrArg);
  TFMul(ecran).BChercheClick(Ecran);

end;

procedure TOF_PGRETROACTIV.btInsertREMOnClick;
var
  StrArg                : String;

begin
  StrArg                := 'MONOFICHE;ACTION=CREATION;AAA';
  PgRetroactiv_LanceFiche(StrArg);
  TFMul(ecran).BChercheClick(Ecran);

end;

procedure TOF_PGRETROACTIV.btSimulOnClick(Sender: TObject);
var
  StrArg : String;
  F : TFmul;
  i : integer;
begin
  F := TFmul(Ecran);
  if (F.FListe.nbSelected=0) and (not F.FListe.AllSelected) then
  begin
    PGIInfo('Aucun élément sélectionné.',F.Caption);
    exit;
  end;

  if F.FListe.AllSelected then
  begin
  {$IFDEF EAGLCLIENT}
      if not F.FetchLesTous then
        PGIInfo('Impossible de récupérer tous les enregistrements')
      else
  {$ENDIF}
     F.Q.First;
     while not F.Q.Eof do
     begin
       if GetField('PGR_RUBRIQUE') = null then
          exit;
       StrArg := GetField('PGR_RUBRIQUE') + ';' + GetField('PGR_NODOSSIER') +
                 ';' + GetField('PGR_PREDEFINI') + ';' + DateToText(GetField('PGR_DATEVALIDITE')) +
                 ';' + DateToText(GetField('PGR_DATEDEBUT')) + ';' + DateToText(GetField('PGR_DATEFIN'));
       if GetField('PGR_NATURERUB') = 'COT' then
         PgRetroSimulcot_LanceFiche(StrArg);
       if GetField('PGR_NATURERUB') = 'AAA' then
         PgRetroSimulrem_LanceFiche(StrArg);

       F.Q.Next;
     end;
  end
  else
  begin
     for i := 0 to F.FListe.nbSelected-1 do
     begin
       F.Fliste.GotoLeBookmark(i);
       {$IFDEF EAGLCLIENT}
       F.Q.TQ.Seek(F.FListe.row -1);
       {$ENDIF}

       if GetField('PGR_RUBRIQUE') = null then
          exit;
       StrArg := GetField('PGR_RUBRIQUE') + ';' + GetField('PGR_NODOSSIER') +
                 ';' + GetField('PGR_PREDEFINI') + ';' + DateToText(GetField('PGR_DATEVALIDITE')) +
                 ';' + DateToText(GetField('PGR_DATEDEBUT')) + ';' + DateToText(GetField('PGR_DATEFIN'));
       if GetField('PGR_NATURERUB') = 'COT' then
         PgRetroSimulcot_LanceFiche(StrArg);
       if GetField('PGR_NATURERUB') = 'AAA' then
         PgRetroSimulrem_LanceFiche(StrArg);
      end;
  end;
  F.BChercheClick(Ecran);
  F.FListe.ClearSelected;
end;

procedure TOF_PGRETROACTIV.btinsertOnClick(Sender: Tobject);
begin
  if GetControlText('PGR_NATURERUB') = 'COT' then
    btInsertCOTOnClick;

  if GetControlText('PGR_NATURERUB') = 'AAA' then
    btInsertREMOnClick;

end;

procedure TOF_PGRETROACTIV.ChangeNature(Sender: Tobject);
var
  Rubrique              : THEdit;
  TRubrique             : THLabel;

begin

  Rubrique              := THEdit(GetControl('PGR_RUBRIQUE'));
  Trubrique             := THLabel(GetControl('TPGR_RUBRIQUE'));
  SetControlText('PGR_RUBRIQUE','');

  if GetControlText('PGR_NATURERUB') = 'COT' then
  begin
    Rubrique.DataType       := 'PGCOTIS';
    Rubrique.ElipsisButton  := True;
    Trubrique.Caption       := 'Cotisation';
  end;

  if GetControlText('PGR_NATURERUB') = 'AAA' then
  begin
    Rubrique.DataType       := 'PGREMUNERATION';
    Rubrique.ElipsisButton  := True;
    Trubrique.Caption       := 'Rémunération';
  end;

  if GetControlText('PGR_NATURERUB') = 'BAS' then
  begin
    Rubrique.DataType       := '';
    Rubrique.ElipsisButton  := False;
    Trubrique.Caption       := 'Base';
  end;

  TFMul(Ecran).BChercheClick(Ecran);
end;

Initialization
  registerclasses ( [ TOF_PGRETROACTIV ] ) ;
end.
