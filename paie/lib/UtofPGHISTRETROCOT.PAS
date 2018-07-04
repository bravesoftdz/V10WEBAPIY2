{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 24/08/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGHISTRETRO ()
Mots clefs ... : TOF;PGHISTRETROCOT
*****************************************************************
PT1   24/08/2007 FC  V_72 : FQ 14688 ajouter un message lors de la suppression d'une ligne déjà intégrée dans le bulletin 
}
Unit UtofPGHISTRETROCOT ;

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
     HTB97,
     ed_tools,
     UTOF ;

Type
  TOF_PGHISTRETROCOT = Class (TOF)
    public

    btSupp : TToolbarButton97;


    procedure SupprimeRegul;
    procedure btSuppOnClick ( Sender : Tobject );

    procedure OnArgument (Arguments : String ) ; override ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

procedure TOF_PGHISTRETROCOT.OnArgument (Arguments : String ) ;
var
  Rubrique : String;
  Datevalidite : String;
  LeWhere : String;

begin
  Inherited ;
  if Arguments<>'' then
  begin
    Rubrique := READTOKENST(Arguments);
    Datevalidite := READTOKENST(Arguments);

    LeWhere := 'PGT_NATURERUB="COT" AND PGT_RUBRIQUE="'+Rubrique+
                '" AND PGT_DATEVALIDITE="'+USDATETIME(strtodate(Datevalidite))+'"';
  end
  else
    LeWhere := 'PGT_NATURERUB="COT"';


  SetControlText('XX_WHERE', LeWhere);

  btSupp := TToolbarButton97(GetControl('BSUPPRIMER'));
  btSupp.OnClick := btSuppOnClick;

end ;

procedure TOF_PGHISTRETROCOT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGHISTRETROCOT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGHISTRETROCOT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGHISTRETROCOT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGHISTRETROCOT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGHISTRETROCOT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGHISTRETROCOT.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_PGHISTRETROCOT.btSuppOnClick(Sender: Tobject);
var
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
     InitMoveProgressForm(nil,'Suprression en cours ...','',(F.Q.recordcount),True,True);
     while not F.Q.Eof do
     begin
        SupprimeRegul;
        MoveCurProgressForm();
        F.Q.Next;
     end;
  end
  else
  begin
     InitMoveProgressForm(nil,'Suppression en cours','',(F.FListe.nbSelected-1),True,True);
     for i := 0 to F.FListe.nbSelected-1 do
     begin
        F.Fliste.GotoLeBookmark(i);
        {$IFDEF EAGLCLIENT}
        F.Q.TQ.Seek(F.FListe.row -1);
        {$ENDIF}

        SupprimeRegul;
        MoveCurProgressForm();

     end;
  end;
  FiniMoveProgressForm();
  F.BChercheClick(Ecran);
  F.FListe.ClearSelected;


end;

procedure TOF_PGHISTRETROCOT.SupprimeRegul;
var
  StrSql : String;
  LeWhere : String;

begin
  StrSql := 'SELECT 1 FROM PGHISTRETRO WHERE';
  LeWhere := ' PGT_NATURERUB="COT" AND PGT_RUBRIQUE="'+GetField('PGT_RUBRIQUE')+
             '" AND PGT_SALARIE="'+GetField('PGT_SALARIE')+'" AND PGT_DATEVALIDITE="'+
             USDATETIME(GetField('PGT_DATEVALIDITE'))+'" AND PGT_DATEINTEG=""';
  if ExisteSQL(StrSql+LeWhere) then
  begin
    ExecuteSQL('DELETE FROM PGHISTRETRO WHERE'+LeWhere);
  end
  else
    PGIBox('Vous ne pouvez pas supprimer la régularisation car elle est déjà intégrée dans un bulletin de paie.' + #13 +
      ' Vous devez préalablement supprimer le bulletin de paie du salarié',Ecran.Caption);  //PT1

end;

Initialization
  registerclasses ( [ TOF_PGHISTRETROCOT ] ) ;
end.
