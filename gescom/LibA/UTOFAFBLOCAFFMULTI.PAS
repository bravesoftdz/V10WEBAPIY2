{***********UNITE*************************************************
Auteur  ...... : PL
Créé le ...... : 22/03/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : AFBLOCAFFMULTI ()
Mots clefs ... : TOF;AFBLOCAFFMULTI
*****************************************************************}
Unit UTOFAFBLOCAFFMULTI ;

Interface

Uses StdCtrls, Controls, Classes,  forms, sysutils,  ComCtrls,
{$IFDEF EAGLCLIENT}

{$ELSE}
     dbTables, db,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, UTOF, utob, DicoAf ;

Type
  TOF_AFBLOCAFFMULTI = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (stArgument : String ) ; override ;
    procedure OnClose                  ; override ;
    private
    TobBlocAff : TOB;
    function DernierNumBlocage:integer ;
  end ;

Implementation


procedure TOF_AFBLOCAFFMULTI.OnArgument (stArgument : String ) ;
var
  req, {st, }Critere,champ, valeur : string;
  NumCourant, x : integer;
  Q : TQuery;

begin
  Inherited;
  NumCourant := 0;
  Critere := (Trim(ReadTokenSt(stArgument)));
  While (Critere <>'') do
  BEGIN
    X:=pos('=',Critere);
    if x<>0 then
    begin
      Champ:=copy(Critere,1,X-1);
      Valeur:=Copy (Critere,X+1,length(Critere)-X);
    end;

    if (Champ='NUM') then NumCourant := strtoint(Valeur);
    Critere:=(Trim(ReadTokenSt(stArgument)));
  END;

  // Recuperation de la tob du blocage à copier
  TobBlocAff := TOB.Create ('BLOCAGEAFFAIRE', Nil, -1);
// SELECT *  : nombre de champ et d'enregistrements limité, on a besoin de l'essentiel
  req := 'SELECT * FROM BLOCAGEAFFAIRE WHERE ABA_NUMBLOCAGE="'+ inttostr(NumCourant)+'"';
  Q := nil;
  try
    Q:=OpenSQL(Req,True);
    If (Not Q.EOF) then
      TobBlocAff.SelectDB('',Q);
  finally
    Ferme(Q);
  end;

  SetControlText('ABA_TYPEBLOCAGE', TobBlocAff.GetValue('ABA_TYPEBLOCAGE'));
  SetControlText('TTABA_TYPEBLOCAGE', rechdom('AFTYPEBLOCAGEAFFAIRE',TobBlocAff.getValue('ABA_TYPEBLOCAGE'),false));
  SetControlText('ABA_ETATAFFAIRE', TobBlocAff.GetValue('ABA_ETATAFFAIRE'));
  SetControlText('TTABA_ETATAFFAIRE', rechdom('AFETATAFFAIRE',TobBlocAff.getValue('ABA_ETATAFFAIRE'),false));
  SetControlText('ABA_OPERATEUR', TobBlocAff.GetValue('ABA_OPERATEUR'));
  if (TobBlocAff.getValue('ABA_OPERATEUR')<>'') then
      SetControlText('TTABA_OPERATEUR', rechdom('AFOPERATEUR',TobBlocAff.getValue('ABA_OPERATEUR'),false));
  SetControlText('ABA_ALERTE',TobBlocAff.GetValue('ABA_ALERTE'));
end ;

procedure TOF_AFBLOCAFFMULTI.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFBLOCAFFMULTI.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFBLOCAFFMULTI.OnUpdate ;
var TobNewBlocs, Tobbloc : TOB;
    //Q : TQuery ;
    i : integer;
    SQL, {req, }sEvenements, sGroupes, SG, sEvt, sGrp, sLibelle:string;
    bTousGroupes:boolean;
begin
  Inherited ;
  bTousGroupes:=false;
  sEvenements:=GetControlText('ABA_EVENEMENTAFF');
  sGroupes:=GetControlText('ABA_GROUPE');
  if (sEvenements='') then
  begin
    LastError:=1; LastErrorMsg:=TraduitGa('Veuillez sélectionner au-moins un événement');
    SetFocusControl('ABA_EVENEMENTAFF');
    exit;
  end
  else
  if (sGroupes='') or (sGroupes='<<Tous>>') then
  begin
    If (PGIAskAf('Veuillez confirmer la génération des blocages affaires définis sur tous les groupes',Ecran.caption)<> mrYes) then
    begin
      LastError:=2;
      SetFocusControl('ABA_GROUPE');
      exit;
    end;
    bTousGroupes:=True;
  end
  else If (PGIAskAf('Veuillez confirmer la génération des blocages affaires définis',Ecran.caption)<> mrYes) then
  begin
    LastError:=3;
    SetFocusControl('ABA_EVENEMENTAFF');
    exit;
  end;

  // Génération des multiples blocages à générer
  i:=DernierNumBlocage;
  if (i=0) then exit;
  TobNewBlocs := TOB.Create ('New blocages', Nil, -1);
  try
    sEvt:=(Trim(ReadTokenSt(sEvenements)));
    While (sEvt <>'') do
    BEGIN
      if bTousGroupes then SG := '' else SG := sGroupes;
      Repeat
        sGrp:=(Trim(ReadTokenSt(SG)));
        if (sGrp='') and (Not bTousGroupes) then continue;
        SQL := 'SELECT ABA_NUMBLOCAGE FROM BLOCAGEAFFAIRE WHERE ABA_EVENEMENTAFF="'+sEvt
            +'" AND ABA_TYPEBLOCAGE="' + TobBlocAff.GetValue('ABA_TYPEBLOCAGE')
            +'" AND ABA_ETATAFFAIRE="' + TobBlocAff.GetValue('ABA_ETATAFFAIRE');
        if bTousGroupes then
            SQL := SQL +'" AND ABA_GROUPE=""'
        else
            SQL := SQL +'" AND (ABA_GROUPE="' + sGrp + '" OR ABA_GROUPE="")';

        if Not (ExisteSQL(SQL)) then
            begin
            Tobbloc := TOB.Create('BLOCAGEAFFAIRE', TobNewBlocs, -1);
            Tobbloc.Dupliquer(TobBlocAff,True,True,False);
            Tobbloc.PutValue('ABA_NUMBLOCAGE', i);
            Tobbloc.PutValue('ABA_EVENEMENTAFF', sEvt);
            if (Tobbloc.GetValue('ABA_TYPEBLOCAGE')='EAF') then
                sLibelle := TraduitGA(rechdom('AFTYPEBLOCAGEAFFAIRE',Tobbloc.GetValue('ABA_TYPEBLOCAGE'),false) + ' ' + rechdom('AFETATAFFAIRE',Tobbloc.GetValue('ABA_ETATAFFAIRE'),false))
            else
                sLibelle := TraduitGA(rechdom('AFOPERATEUR',Tobbloc.GetValue('ABA_OPERATEUR'),false) + ' à la '+ rechdom('AFTYPEBLOCAGEAFFAIRE',Tobbloc.GetValue('ABA_TYPEBLOCAGE'),false));

            sLibelle := sLibelle + ', ' +TraduitGA(rechdom('AFEVTBLOCAFFAIRE',sEvt,false));
            if bTousGroupes or (sGrp='') then
                sLibelle := sLibelle + ', tous groupes'
            else
                sLibelle := sLibelle + ', ' + rechdom('TTUSERGROUPE',sGrp,false);
            Tobbloc.PutValue('ABA_LIBELLE', sLibelle);
            if (bTousGroupes) then
                Tobbloc.PutValue('ABA_GROUPE', '')
            else
                Tobbloc.PutValue('ABA_GROUPE', sGrp);

            i:=i+1;
            end;
        Until (sGrp='');

      sEvt:=(Trim(ReadTokenSt(sEvenements)));
    END;

    TobNewBlocs.InsertDB(nil);

  finally
    TobNewBlocs.Free;
    Ecran.close;
  end;
end;

procedure TOF_AFBLOCAFFMULTI.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFBLOCAFFMULTI.OnClose ;
begin
  Inherited ;
TobBlocAff.Free;
end ;


function TOF_AFBLOCAFFMULTI.DernierNumBlocage:integer ;
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

  Result:=iMax;
end;

Initialization
  registerclasses ( [ TOF_AFBLOCAFFMULTI ] ) ;
end.
