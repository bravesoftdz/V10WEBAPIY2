{ Unité : Source TOF de la FICHE : CPMULRUB ()
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
4.3.1.101.001    18/06/03    JP     Création de l'unité

--------------------------------------------------------------------------------------}
unit CPQRTRUBRIQUE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     maineagl,eQRS1,HPdfPrev,UtileAGL,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_main,QRS1,EdtREtat,
{$ENDIF}
     UTOB, Forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, HTB97 ;

type
  TResCpt = record
    L1 : string;
    L2 : string;
    L3 : string;
    L4 : string;
  end;

  TResCod = record
    C1 : string;
    C2 : string;
  end;

  TOF_CPQRTRUBRIQUE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    cbFamille : THMultiValComboBox;
    cbRubDe   : THEdit;
    cbRubA    : THEdit;
    cbSigne   : THValComboBox;
    cbTypeRub : THValComboBox;
    rbType    : THRadioGroup;
    rbTri     : THRadioGroup;
    DetailRub : TCheckBox;

    procedure RB_OrdreTriChange  (Sender : TObject);
//    procedure BImprimerClick     (Sender : TObject);
    procedure TypeEditionClick   (Sender : TObject);

    function  RenseigneCritere      : string;
    procedure ConstituerClauseSQL(var Requete : string);
  public
    class procedure RemplirTob   (var MaTob : TOB; Clause : string; Detail, Bud : Boolean);
    class procedure PresenterTob (var MaTob : TOB);

    class function  RemplitDetail(TypeRub : string) : TResCpt;
    class function  RemplitMemo  (var Ch  : string; QuelCas : Integer; TT : string): TResCod;
    class function  TransBud     (St      : string) : string;

  end ;

procedure ImpRubrique(UnCpte : string ; DuBouton, Bud : Boolean);


Implementation

uses
    {$IFDEF eAGLCLIENT}
    MenuOLX
    {$ELSE}
    MenuOLG
    {$ENDIF eAGLCLIENT}
    ;

var
  Budget : Boolean;

{---------------------------------------------------------------------------------------}
procedure ImpRubrique(UnCpte : string ; DuBouton, Bud : Boolean);
{---------------------------------------------------------------------------------------}
begin
  Budget := Bud;

  if DuBouton then
  // appeler les fonctions ad hoc
    LanceEtat('E', 'RUB', 'FRU', True, False, False, nil, 'SELECT * FROM RUBRIQUE WHERE RB_RUBRIQUE = "' + UnCpte + '"', 'Impression de la rubrique : ' + UnCpte, False, 0)
  else
    AGLLanceFiche('CP', 'CPQRTRUBRIQUE', '', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPQRTRUBRIQUE.OnNew ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPQRTRUBRIQUE.OnDelete ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPQRTRUBRIQUE.OnUpdate ;
{---------------------------------------------------------------------------------------}
var
  ClauseWhere : string;
begin
  Inherited ;
  if rbType.ItemIndex = 0 then begin
    ConstituerClauseSQL(ClauseWhere);
    TFQRS1(Ecran).WhereSQL := ClauseWhere;
  end
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPQRTRUBRIQUE.OnLoad ;
{---------------------------------------------------------------------------------------}
var
  ClauseWhere : string;
  MaTob       : TOB;
begin
 TFQRS1(Ecran).NatureEtat := 'RUB';
 TFQRS1(Ecran).TypeEtat := 'E';
 TFQRS1(Ecran).CodeEtat := 'LRU';
  if rbType.ItemIndex = 0 then begin
    SetControlText('MUL', '');

//    LanceEtat('E', 'RUB', 'LRU', True, False, False, TPageControl(GetControl('PAGES')), ClauseWhere , 'Impression des rubriques', False, 0, RenseigneCritere);
  end
  else begin
    ConstituerClauseSQL(ClauseWhere);
    RemplirTob(MaTob, ClauseWhere , DetailRub.Checked, Budget);
    LanceEtatTOB('E', 'RUB', 'IRU', MaTob, True, False, False, nil, '', 'Liste des rubriques', False, 0, RenseigneCritere);
    FreeAndNil(MaTob);
  end;
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPQRTRUBRIQUE.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
  cbRubDe   := THEdit(GetControl('RUB'));
  cbRubA    := THEdit(GetControl('RUB1'));
  cbSigne   := THValComboBox(GetControl('SIGNERUB'));
  cbTypeRub := THValComboBox(GetControl('TYPERUB'));
  cbFamille := THMultiValComboBox(GetControl('FAMILLE'));

  rbType    := THRadioGroup(GetControl('FTYPEIMPRESSION'));
  rbTri     := THRadioGroup(GetControl('FTRI'));
  DetailRub := TCheckBox(GetControl('DETAILRUB'));


//  TToolbarButton97(GetControl('BIMPRIMER')).OnClick := BImprimerClick;

  rbType.OnClick := TypeEditionClick;
  rbTri .OnClick := RB_OrdreTriChange;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPQRTRUBRIQUE.OnClose ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPQRTRUBRIQUE.OnDisplay () ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPQRTRUBRIQUE.OnCancel () ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPQRTRUBRIQUE.RB_OrdreTriChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  DetailRub.Enabled := (rbType.ItemIndex = 1);
  DetailRub.Checked := (rbType.ItemIndex = 1);

  if rbType.ItemIndex = 0 then
    TFQRS1(Ecran).Caption := 'Impression des rubriques en fiche'
  else
    TFQRS1(Ecran).Caption := 'Impression des rubriques en liste';
  UpdateCaption(Ecran);
end;

{---------------------------------------------------------------------------------------}
(*procedure TOF_CPQRTRUBRIQUE.BImprimerClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  ClauseWhere : string;
  MaTob       : TOB;
begin
  if rbType.ItemIndex = 0 then begin
    ConstituerClauseSQL(ClauseWhere);
    LanceEtat('E', 'RUB', 'LRU', True, False, False, TPageControl(GetControl('PAGES')), ClauseWhere , 'Impression des rubriques', False, 0, RenseigneCritere);
  end
  else begin
    ConstituerClauseSQL(ClauseWhere);
    RemplirTob(MaTob, ClauseWhere, DetailRub.Checked, Budget);
    LanceEtatTOB('E', 'RUB', 'IRU', MaTob, True, False, False, nil, '', 'Liste des rubriques', False, 0, RenseigneCritere);
    FreeAndNil(MaTob);
  end;
end;*)

{---------------------------------------------------------------------------------------}
procedure TOF_CPQRTRUBRIQUE.TypeEditionClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  DetailRub.Enabled := (rbType.ItemIndex = 1);
  DetailRub.Checked := (rbType.ItemIndex = 1);
end;

{---------------------------------------------------------------------------------------}
class procedure TOF_CPQRTRUBRIQUE.RemplirTob(var MaTob : TOB; Clause : string; Detail, Bud : Boolean);
{---------------------------------------------------------------------------------------}
var
  Cpt1, Cpt2,
  Exc1, Exc2,
  Rub, Lib,
  Sig, Typ,
  Axe, Fam,
  ttRub, C2,
  E1, E2 : string;
  n, nbmax : Integer;
  TTmp   : TOB;
  Q      : TQuery;
  LibCpt : TResCpt;
  CodFam,
  CodCpt : TResCod;
  Double : Byte; {0 : pas de détail, 1 : détail simple, 2 : détail double}
begin
  n := 0;
  nbmax:=6000; // YMO FQ17724 11/04/2006 Nombre max d'enreg avec détail porté à 1800
  if Bud then ttRub := 'ttRubTypeBud'
         else ttRub := 'ttRubFamille';

  MaTob := TOB.Create('§LISTERUBRIQUE', nil, -1);
  Q := OpenSQL(Clause, False);

  while not Q.EOF do begin
    Rub  := Q.FindField('RB_RUBRIQUE').AsString;
    Lib  := Q.FindField('RB_LIBELLE' ).AsString;
    Sig  := Q.FindField('RB_SIGNERUB').AsString;
    Typ  := Q.FindField('RB_TYPERUB' ).AsString;
    Axe  := Q.FindField('RB_AXE'     ).AsString;
    Fam  := Q.FindField('RB_FAMILLES').AsString;
    Cpt1 := Q.FindField('RB_COMPTE1' ).AsString;
    Cpt2 := Q.FindField('RB_COMPTE2' ).AsString;
    Exc1 := Q.FindField('RB_EXCLUSION1').AsString;
    Exc2 := Q.FindField('RB_EXCLUSION2').AsString;

    Double := 0;
    {Gestion des libellés du détail des comptes}
    if Detail then begin
      nbmax:=1800;
      LibCpt := RemplitDetail(Typ);
      Double := 2;
      {Si détail simple, on permute L2 et L3 por centrer l'affichage}
      if (Trim(LibCpt.L3) = '') and (Trim(LibCpt.L4) = '') then begin
        LibCpt.L3 := LibCpt.L2;
        LibCpt.L2 := '';
        Double := 1;
      end;
    end
    else
      Cpt1 := '';

    if (Cpt1 <> '') and (Cpt1[Length(Cpt1)] <> ';') then Cpt1 := Cpt1 + ';';
    if (Fam  <> '') and (Fam [Length(Fam) ] <> ';') then Fam  := Fam  + ';';

    while (Cpt1 <> '') or (Fam <> '') do begin
      {On décompose les "Champs Memo" (suite de valeurs séparées de ';')}
      {Récupération des familles de rubriques et de leur libellé}
      CodFam := RemplitMemo(Fam , 0, ttRub);
      {Récupération du détail des comptes}
      if Double > 0 then begin
        CodCpt := RemplitMemo(Cpt1, 2, 'ttRubCalcul');
        E1     := RemplitMemo(Exc1, 1, '').C1;
        if Double > 1  then begin
          C2   := RemplitMemo(Cpt2, 2, '').C1;
          E2   := RemplitMemo(Exc2, 1, '').C1;
        end
        else begin
          C2 := E1;
          E1 := '';
          E2 := '';
        end;
      end
      else begin
        E1 := '';
        CodCpt.C1 := '';
        CodCpt.C2 := '';
      end;

      if (Trim(CodCpt.C1) <> '') or (Trim(CodFam.C1) <> '') then begin
        TTmp := TOB.Create('$Fille', MaTob, -1);
        PresenterTob(TTmp);

        TTmp.PutValue('RB_RUBRIQUE', Rub);
        TTmp.PutValue('RB_LIBELLE' , Lib);
        TTmp.PutValue('RB_SIGNERUB', Sig);
        TTmp.PutValue('RB_TYPERUB' , Typ);
        TTmp.PutValue('RB_AXE'     , Axe);
        TTmp.PutValue('RB_CODEFAM' , CodFam.C1);
        TTmp.PutValue('RB_LIBFAM'  , CodFam.C2);
        TTmp.PutValue('RB_CPTGEN1' , CodCpt.C1);
        TTmp.PutValue('RB_CPTEXC1' , E2);
        TTmp.PutValue('RB_CPTGEN2' , C2);
        TTmp.PutValue('RB_CPTEXC2' , E2);
        TTmp.PutValue('RB_SOLDE'   , CodCpt.C2);
        TTmp.PutValue('RB_LIB1'    , LibCpt.L1);
        TTmp.PutValue('RB_LIB2'    , LibCpt.L2);
        TTmp.PutValue('RB_LIB3'    , LibCpt.L3);
        TTmp.PutValue('RB_LIB4'    , LibCpt.L4);

        {Pour limiter à 6000 enregistrements}
        Inc(n);
        if n > nbmax then begin
          HShowMessage('0;Impression des rubriques;Les critères de sélection correspondent à plus de '+inttostr(nbmax)+' enregistrements !'#10#13 +
                       'Seuls les '+inttostr(nbmax)+' premiers seront affichés;Q;O;N;N;', '', '');
          Ferme(Q);
          Exit;
        end;
      end;
    end;
    Q.Next;
  end;
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPQRTRUBRIQUE.ConstituerClauseSQL(var Requete : string);
{---------------------------------------------------------------------------------------}
var
  Traduction,
  St, StOr : string;
begin
  Requete := 'SELECT * FROM RUBRIQUE ';
  Requete := Requete + 'WHERE RB_RUBRIQUE <> "' + W_W+ '" ';
  if Budget then Requete := Requete + 'AND RB_NATRUB =  "BUD" AND RB_BUDJAL <> "" '
            else Requete := Requete + 'AND RB_NATRUB <> "BUD" AND RB_BUDJAL =  "" ';

  if cbRubDe.Text <> '' then Requete := Requete + 'AND RB_RUBRIQUE >= "' + cbRubDe.Text + '" ';
  if cbRubA .Text <> '' then Requete := Requete + 'AND RB_RUBRIQUE <= "' + cbRubA .Text + '" ';

  if GetControlText('LIBELLE') <> '' then begin
    Traduction := TraduitJoker(GetControlText('LIBELLE'));
    Requete := Requete + 'AND UPPER(RB_LIBELLE) LIKE "' + Traduction + '" ' ;
  end;

  if cbSigne  .ItemIndex >= 0 then Requete := Requete + 'AND RB_SIGNERUB = "' + cbSigne  .Value + '" ';
  if cbTypeRub.ItemIndex >= 0 then Requete := Requete + 'AND RB_TYPERUB  = "' + cbTypeRub.Value + '" ';

  if not cbFamille.Tous  then begin
    St := cbFamille.Text;
    if St[Length(St)] <> ';' then St := St + ';';
    StOr := '';
    while St <> '' do StOr := StOr + 'OR RB_FAMILLES LIKE "%' + ReadTokenSt(St) + '%" ';
    if StOR <> '' then Requete := Requete + 'AND (' + StOr + ') ';
  end ;
  case rbTri.ItemIndex of
    0 : Requete := Requete + 'ORDER BY RB_RUBRIQUE';
    1 : Requete := Requete + 'ORDER BY RB_LIBELLE';
  end;
//  end
  //else
    //QRub.SQL.Add(' And RB_RUBRIQUE>="'+FRub1.Value+'"') ; QRub.SQL.Add(' And RB_RUBRIQUE<="'+FRub2.Value+'"') ;
end;

{---------------------------------------------------------------------------------------}
class procedure TOF_CPQRTRUBRIQUE.PresenterTob(var MaTob : TOB);
{---------------------------------------------------------------------------------------}
begin
  MaTob.AddChampSup('RB_RUBRIQUE', True);
  MaTob.AddChampSup('RB_LIBELLE' , True);
  MaTob.AddChampSup('RB_SIGNERUB', True);
  MaTob.AddChampSup('RB_TYPERUB' , True);
  MaTob.AddChampSup('RB_AXE'     , True);
  MaTob.AddChampSup('RB_CODEFAM' , True);
  MaTob.AddChampSup('RB_LIBFAM'  , True);
  MaTob.AddChampSup('RB_CPTGEN1' , True);
  MaTob.AddChampSup('RB_CPTEXC1' , True);
  MaTob.AddChampSup('RB_CPTGEN2' , True);
  MaTob.AddChampSup('RB_CPTEXC2' , True);
  MaTob.AddChampSup('RB_SOLDE'   , True);
  MaTob.AddChampSup('RB_LIB1'    , True);
  MaTob.AddChampSup('RB_LIB2'    , True);
  MaTob.AddChampSup('RB_LIB3'    , True);
  MaTob.AddChampSup('RB_LIB4'    , True);
end;

{---------------------------------------------------------------------------------------}
class function TOF_CPQRTRUBRIQUE.TransBud(St : string) : string;
{---------------------------------------------------------------------------------------}
begin
  if St = 'CBG' then Result := 'GEN' else
  if St = 'CBS' then Result := 'ANA' else
  if St = 'G/S' then Result := 'G/A' else
  if St = 'S/G' then Result := 'A/G';
end ;

{Retourne les critères de l'état
{---------------------------------------------------------------------------------------}
function TOF_CPQRTRUBRIQUE.RenseigneCritere : string;
{---------------------------------------------------------------------------------------}
begin
  {Detail des comptes des rubriques ?}
  if (rbType.ItemIndex = 0) and DetailRub.Checked then Result := 'DETAIL=NON`';
  {Critères de l'édition}
  Result := Result + 'RUB1=' + cbRubDe.Text + '`RUB2=' + cbRubA.Text + '`SIGNE=' + cbSigne.Value +
            '`FAMILLE=' + cbFamille.Text + '`LIBELLE=' + GetControlText('LIBELLE') +
            '`TYPE=' + cbTypeRub.Value;// + '`MUL=';
end;

{---------------------------------------------------------------------------------------}
class function TOF_CPQRTRUBRIQUE.RemplitMemo(var Ch : string; QuelCas : Integer; TT : string): TResCod;
{---------------------------------------------------------------------------------------}
var
  STC,
  STB,
  STL : string;
  j   : Integer;
begin
  Result.C1 := '';
  Result.C2 := '';

  if Ch <> '' then begin
   STC := ReadTokenSt(Ch);
   if STC = '' then Exit;
   case QuelCas of
     {Gestion des familles}
     0 : begin
           STB := STC;
           if Budget   then STB := TransBud(STC);
           if TT <> '' then STL := RechDom(TT, STB, False);
           Result.C1 := STC;
           Result.C2 := STL;
         end;
     {Gestion des comptes exclus}
     1 : Result.C1 := STC;
     {Gestion des comptes et des modes de calcul}
     2 : begin
           j := Pos('(', STC);
           if j > 0 then begin
             {Recherche du mode de calcul}
             STL := Copy(STC, j + 1, 2);
             if TT <> '' then begin
               STL := RechDom(TT, STL, False);
               Result.C2 := STL;
             end;
             {Recherche du compte}
             STC := Copy(STC, 1, j - 1);
             Result.C1 := STC;
           end;
         end;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
class function TOF_CPQRTRUBRIQUE.RemplitDetail(TypeRub : string) : TResCpt;
{---------------------------------------------------------------------------------------}
var
  C1,C2  : Char;
  Double : Boolean;
  MsgBox : THMsgBox;
begin
  MsgBox := THMsgBox.create(FMenuG);
  MsgBox.Mess.Add('Sections');
  MsgBox.Mess.Add('Sect. Except.');
  MsgBox.Mess.Add('Cptes Budgétaires');
  MsgBox.Mess.Add('Bud. Except.');
  MsgBox.Mess.Add('Cptes Auxiliaires');
  MsgBox.Mess.Add('Auxi. Except.');
  MsgBox.Mess.Add('Cptes Généraux');
  MsgBox.Mess.Add('Géné. Except.');
  MsgBox.Mess.Add('Calcul Rub.');

  C1 := TypeRub[1];
  C2 := TypeRub[3];
  Double := (TypeRub[2] = '/');

  case C1 of
    'A' : begin Result.L1 := MsgBox.Mess[0]; Result.L2 := MsgBox.Mess[1]; end;
    'B' : begin Result.L1 := MsgBox.Mess[2]; Result.L2 := MsgBox.Mess[3]; end;
    'T' : begin Result.L1 := MsgBox.Mess[4]; Result.L2 := MsgBox.Mess[5]; end;
    'G' : begin Result.L1 := MsgBox.Mess[6]; Result.L2 := MsgBox.Mess[7]; end;
   end;

  if Double then begin
    case C2 of
      'A' : begin Result.L3 := MsgBox.Mess[0]; Result.L4 := MsgBox.Mess[1]; end;
      'B' : begin Result.L3 := MsgBox.Mess[2]; Result.L4 := MsgBox.Mess[3]; end;
      'T' : begin Result.L3 := MsgBox.Mess[4]; Result.L4 := MsgBox.Mess[5]; end;
      'G' : begin Result.L3 := MsgBox.Mess[6]; Result.L4 := MsgBox.Mess[7]; end;
    end ;
  end;
  FreeAndNil(MsgBox);
end;

Initialization
  registerclasses ( [ TOF_CPQRTRUBRIQUE ] );

end.


