{***********UNITE*************************************************
Auteur  ...... : PL
Créé le ...... : 20/03/2001
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : BLOCAGEAFFAIRE (BLOCAGEAFFAIRE)
Mots clefs ... : TOM;BLOCAGEAFFAIRE
*****************************************************************}
Unit UTOMBLOCAGEAFFAIRE ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
{$IFDEF EAGLCLIENT}
    Maineagl, eFiche,  eFichList,
      UtileAGL,
{$ELSE}
   dbTables, db, dbctrls, FE_Main, Fiche,FichList,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, UTOM,
     UtofAfBlocAffMulti, // à laisser pour obliger projet à avoir cetet tof utilisé depuis script
      UTob, M3FP, DicoAF ;

Type
  TOM_BLOCAGEAFFAIRE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    procedure Duplication;
    end ;


const
	// libellés des messages de la TOM Ressource
	TexteMsgBlocage: array[1..5] of string 	= (
          {1}         'Ce blocage existe déjà'
          {2}        ,'Veuillez choisir un événement de déclenchement'
          {3}        ,'Veuillez choisir un type de blocage'
          {4}        ,'Veuillez choisir un état de déclenchement'
          {5}        ,'Veuillez choisir un opérateur de test pour le déclenchement'
    );
Procedure AFLanceFiche_BlocageAff;
Implementation

procedure TOM_BLOCAGEAFFAIRE.OnNewRecord ;
Var QQ : TQuery ;
    IMax :integer ;
Begin
  inherited;
  QQ := nil;
  try
    QQ:=OpenSQL('SELECT MAX(ABA_NUMBLOCAGE) FROM BLOCAGEAFFAIRE',TRUE);
    if Not QQ.EOF then Imax:=QQ.Fields[0].AsInteger+1 else iMax:=1;
  finally
    Ferme(QQ);
  end;

  SetField('ABA_NUMBLOCAGE',IMax);
  SetControlEnabled('BDUPLIQUE', False);
end;

procedure TOM_BLOCAGEAFFAIRE.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_BLOCAGEAFFAIRE.OnUpdateRecord ;
var
SQL,sLibelle :string;
begin
  Inherited ;
if (GetField ('ABA_TYPEBLOCAGE')='') then
    Begin LastError:=3; LastErrorMsg:=TraduitGa(TexteMsgBlocage[LastError]); SetFocusControl ('ABA_TYPEBLOCAGE'); Exit; End;
if (GetField('ABA_EVENEMENTAFF')='') then
    Begin LastError:=2; LastErrorMsg:=TraduitGa(TexteMsgBlocage[LastError]); SetFocusControl ('ABA_EVENEMENTAFF'); Exit; End;
if (GetField ('ABA_TYPEBLOCAGE')='EAF') and (GetField ('ABA_ETATAFFAIRE')='') then
    Begin LastError:=4; LastErrorMsg:=TraduitGa(TexteMsgBlocage[LastError]); SetFocusControl ('ABA_ETATAFFAIRE'); Exit; End;
if (GetField ('ABA_OPERATEUR')='') then
    Begin LastError:=5; LastErrorMsg:=TraduitGa(TexteMsgBlocage[LastError]); SetFocusControl ('ABA_OPERATEUR'); Exit; End;

if (GetField ('ABA_LIBELLE')='') then
    begin
    if (GetField ('ABA_TYPEBLOCAGE')='EAF') then
        sLibelle := TraduitGA(rechdom('AFTYPEBLOCAGEAFFAIRE',GetField ('ABA_TYPEBLOCAGE'),false) + ' ' + rechdom('AFETATAFFAIRE',GetField ('ABA_ETATAFFAIRE'),false))
    else
        sLibelle := TraduitGA(rechdom('AFOPERATEUR',GetField ('ABA_OPERATEUR'),false) + ' à la '+ rechdom('AFTYPEBLOCAGEAFFAIRE',GetField ('ABA_TYPEBLOCAGE'),false));

    sLibelle := sLibelle + ', ' +TraduitGA(rechdom('AFEVTBLOCAFFAIRE',GetField ('ABA_EVENEMENTAFF'),false));

    if (GetField ('ABA_GROUPE')<>'') then
        sLibelle := sLibelle + ', ' + rechdom('TTUSERGROUPE',GetField ('ABA_GROUPE'),false)
    else
        sLibelle := sLibelle + ', tous groupes';
    SetField('ABA_LIBELLE', sLibelle);
    end;

SQL := 'SELECT ABA_NUMBLOCAGE FROM BLOCAGEAFFAIRE WHERE ABA_NUMBLOCAGE<>"'+inttostr(GetField ('ABA_NUMBLOCAGE'))
        +'" AND ABA_EVENEMENTAFF="'+GetField ('ABA_EVENEMENTAFF')
        +'" AND ABA_TYPEBLOCAGE="' + GetField ('ABA_TYPEBLOCAGE')
        +'" AND ABA_ETATAFFAIRE="' + GetField ('ABA_ETATAFFAIRE')
        +'" AND (ABA_GROUPE="' + GetField ('ABA_GROUPE')+ '" OR ABA_GROUPE="")';
if (ExisteSQL(SQL)) then
    Begin LastError:=1; LastErrorMsg:=TraduitGa(TexteMsgBlocage[LastError]); Exit; End;

end ;

procedure TOM_BLOCAGEAFFAIRE.OnAfterUpdateRecord ;
begin
  Inherited ;
SetControlEnabled('BDUPLIQUE', true);
end ;

procedure TOM_BLOCAGEAFFAIRE.OnLoadRecord ;
begin
  Inherited ;
If (GetField ('ABA_TYPEBLOCAGE') <> 'EAF') then
   Begin
   SetControlEnabled ('ABA_ETATAFFAIRE',False);
   SetControlEnabled ('TABA_ETATAFFAIRE',False);
   SetControlEnabled ('ABA_OPERATEUR',True);
   SetField('ABA_ETATAFFAIRE', '');
   End
Else
   Begin
   SetControlEnabled ('ABA_ETATAFFAIRE',True);
   SetControlEnabled ('TABA_ETATAFFAIRE',True);
   SetControlEnabled ('ABA_OPERATEUR',False);
   End;

end ;

procedure TOM_BLOCAGEAFFAIRE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_BLOCAGEAFFAIRE.OnArgument ( S: String ) ;
begin
  Inherited ;

SetControlProperty('ABA_ETATAFFAIRE','Plus',' AND (CC_LIBRE<>"BTP")');

end ;

procedure TOM_BLOCAGEAFFAIRE.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_BLOCAGEAFFAIRE.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_BLOCAGEAFFAIRE.Duplication;
var TobForm : TOB;
    Q : TQuery ;
    i : integer;
    req:string;
begin
  TobForm := TOB.Create ('BLOCAGEAFFAIRE', Nil, -1);
// SELECT * : on a besoin de tout pour dupliquer
  req := 'SELECT * FROM BLOCAGEAFFAIRE WHERE ABA_NUMBLOCAGE="'+ inttostr(GetField ('ABA_NUMBLOCAGE'))+'"';
  Q := nil;
  try
    Q:=OpenSQL(Req,True);
    If (Not Q.EOF) then TobForm.SelectDB('',Q);
  finally
    Ferme(Q);
  end;

  if TobForm <> Nil then
      begin
      if (DS.State <> dsEdit) and (DS.State <> dsInsert) then
        TFFicheListe(Ecran).Bouge(NbInsert);

      for i := 1 to TobForm.NbChamps do
          begin
          if (TobForm.GetNomChamp(i)<>'ABA_NUMBLOCAGE') then
              begin
              SetField (TobForm.GetNomChamp(i), TobForm.GetValeur (i));
              end;
          end;

      If (GetField('ABA_TYPEBLOCAGE') <> 'EAF') then
          begin
          SetControlEnabled ('ABA_ETATAFFAIRE',False);
          SetControlEnabled ('TABA_ETATAFFAIRE',False);
          SetField('ABA_ETATAFFAIRE', '');
          SetControlEnabled ('ABA_OPERATEUR',True);
          end
      Else
          begin
          SetControlEnabled ('ABA_ETATAFFAIRE',True);
          SetControlEnabled ('TABA_ETATAFFAIRE',True);
          SetControlEnabled ('ABA_OPERATEUR',False);
          end;
      end;

  TobForm.Free;

  SetFocusControl ('ABA_LIBELLE');
end;

procedure AGLDuplicationBlocage( parms: array of variant; nb: integer );
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFicheListe) then OM:=TFFicheListe(F).OM else exit;
if (OM is TOM_BLOCAGEAFFAIRE) then TOM_BLOCAGEAFFAIRE(OM).Duplication else exit;
end;

Procedure AFLanceFiche_BlocageAff;
begin
AGLLanceFiche ('AFF','AFBLOCAGEAFFAIRE','','','');
end;


Initialization
  registerclasses ( [ TOM_BLOCAGEAFFAIRE ] ) ;
 RegisterAglProc ( 'DuplicationBlocAff', TRUE , 0, AGLDuplicationBlocage);
end.
