unit UtilMulTrt;

interface

uses  HMsgBox,HEnt1,HCtrls,HQry,SysUtils,
{$IFDEF EAGLCLIENT}
      MaineAGL,eMul,
{$ELSE}
      Mul, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,
{$ENDIF}
      UTOB,Dicobtp;


Type T_ErrTraiteEnreg = (tteAucunEnreg,tteTropEnreg,tteOK) ;

Const
TexteMessage : array[1..2] of string 	= (
          {1}  'Aucun élément sélectionné',
          {2}  'Trop d''éléments sélectionnés. Veuiller réduire la sélection'
              );
FUNCTION TraiteEnregMulListe (Mul : TFMul; ChampCpt,TableCpt : string; TobEnregs : TOB; Parle : Boolean):T_ErrTraiteEnreg;
FUNCTION TraiteEnregMulTable (Mul : TFMul; stSelect,ArgWhere, NomTob, ChampCpt, TableCpt: string;
                              TobEnregs : TOB; Parle : Boolean ; nbEnregMax : integer = 200000) : T_ErrTraiteEnreg;


implementation


FUNCTION TestEnregATraiter (Mul : TFMul; Parle : Boolean) : Boolean;
begin
result := True;
//Test si des enreg sont sélectionnés
if (Mul.FListe.NbSelected=0) and (not Mul.FListe.AllSelected) then
   begin
   if Parle then PGIBoxAf(TexteMessage[1],TitreHalley);
   Result := False;
   end;
end;

FUNCTION VerifNbrPourBlocage ( ChampCpt,TableCpt, stWhere : string;  Parle : Boolean ; nbEnregMax : integer) : Boolean;
Var   QQ : TQuery;
			SQl : string;
begin
Result := True;
if (ChampCpt = '') or (TableCpt = '') then Exit;
	Sql := 'SELECT count(' + ChampCpt + ') as NbrEnreg FROM '+TableCpt+ ' ' + stWhere;
QQ := OpenSQL (Sql, True);
if (Not QQ.EOF) and (QQ.FindField ('NbrEnreg').AsInteger > nbEnregMax) then
   begin
   Result := False;
   if Parle then PGIBoxAf(TexteMessage[2],TitreHalley);
   end;
Ferme(QQ);
end;

Procedure FinTraiteMul (Mul : TFMul);
begin
if Mul.FListe.AllSelected then Mul.FListe.AllSelected:=False
                          else Mul.FListe.ClearSelected;
Mul.bSelectAll.Down := False ;
end;

procedure ChargeUnEnreg (Mul : TFMul; stSelect,ArgWhere, NomTob : String; TobEnregs : TOB);
Var stWhere,Champ,TypeChamp : string;
    QQ : TQuery;
    TobDet : TOB;
begin
  stWhere := '';

  // Fabrication du where
  Champ:=(Trim(ReadTokenSt(ArgWhere)));
  While (Champ <>'') do
  begin
    if stWhere <> '' then stWhere := stWhere + ' AND ';
    TypeChamp := ChampToType (Champ);
    if (TypeChamp = 'INTEGER') or (TypeChamp = 'DOUBLE') or (TypeChamp = 'RATE') then
    begin
      // C.B 01/08/2003
      // GL_NUMERO est dans la vue AFREVMODIFCOEF et pas AFR_NUMERO
      // obligatoire pour renvoyer les données
      // on est donc obligé de faire ce qui suis
      // onyx 
      if (NomTob = 'AFREVISION') and
         (Champ = 'GL_NUMERO') then
      begin
        if (Mul.Q.FindField(Champ).AsString <> '') then
          stWhere := stWhere + 'AFR_NUMERO = ' + Mul.Q.FindField(Champ).AsString
        else
          stWhere := stWhere + 'AFR_NUMERO = 0'
      end
      // cas normal
      else
       stWhere := stWhere + Champ + '=' + Mul.Q.FindField(Champ).AsString;
    end
    else if (TypeChamp = 'DATE') then
       stWhere := stWhere + Champ + '="' + usDatetime(StrToDate(Mul.Q.FindField(Champ).AsString)) + '"'

      // C.B 07/08/2003
      // on ne peut pas faire de select sur plusieurs tables avec jointure
      // on est donc obligé de faire ce qui suis
      // onyx                         
    else if (NomTob = 'AFPARAMFORMULE') and
            (Champ = 'AFC_AFFAIRE') then
      stWhere := stWhere + 'AFC_AFFAIRE = "' + Mul.Q.FindField('AFF_AFFAIRE').AsString+ '"'
    else
     // string par défaut
      stWhere := stWhere + Champ + '="' + Mul.Q.FindField(Champ).AsString + '"';

    Champ:=(Trim(ReadTokenSt(ArgWhere)));
  end;

  if stwhere <> '' then stWhere := ' WHERE ' + stWhere;
  QQ := OpenSql (stSelect + stWhere, True);
  if not QQ.EOF then
  begin
    TobDet := Tob.Create (NomTob,TobEnregs,-1);
    TobDet.SelectDB ('',QQ);
  end;
  Ferme (QQ);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 29/08/2001
Modifié le ... : 08/01/2002
Description .. : fonction de traitement des enregistrements d'un mul de
Suite ........ : traitement. Retourne une tob de tous les enregs sur la
Suite ........ : base de la liste du mul
Suite ........ : ChampCpt et TableCpt utilisés pour comptage des enreg
Suite ........ : attention limite actuelle à 2000 enreg
Suite ........ : Vori aide de EnregMulTable pour la différence entre les 2 fct
Mots clefs ... : MUL;
*****************************************************************}
FUNCTION TraiteEnregMulListe (Mul : TFMul; ChampCpt,TableCpt : string; TobEnregs : TOB; Parle : Boolean):T_ErrTraiteEnreg;
Var stWhere : string;
    i :integer;
begin
Result := tteOK;

if Mul = Nil then Exit;
if (TobEnregs = Nil) then Exit;

if Not(TestEnregAtraiter(Mul,Parle)) then
   begin Result := tteAucunEnreg; Exit; end;
{$IFDEF EAGLCLIENT}
if Mul.bSelectAll.Down then
   if not Mul.Fetchlestous then
     begin
     Mul.bSelectAllClick(Nil);
     Mul.bSelectAll.Down := False;
     exit;
     end;
{$ENDIF}

if Mul.FListe.AllSelected then
   begin
   stWhere := ' ' + RecupWhereCritere(Mul.Pages);
   if Not (VerifNbrPourBlocage ( ChampCpt,TableCpt, stWhere , Parle, 200000 )) then
         begin Result := tteTropEnreg; Exit; end;
{$IFDEF EAGLCLIENT}
   TOBEnregs.LoadDetailDB('LISTE ENREG', '', '', Mul.Q.TQ,false)
{$ELSE}
   try
   THDBGrid(TFMUL(Mul).FListe).datasource.dataset.disablecontrols;
   if not Mul.Q.BOF then Mul.Q.First; // PL : Obligé de remettre au début car le loaddetaildb teste la position du curseur pas EOF
   TOBEnregs.LoadDetailDB('LISTE ENREG', '', '', Mul.Q,false);
   finally
   THDBGrid(TFMUL(Mul).FListe).datasource.dataset.enablecontrols;
   end;
{$ENDIF}
   end
else
   begin
   for i := 0 to Mul.FListe.NbSelected-1 do
      begin
      Mul.FListe.GotoLeBookMark(i);
{$IFDEF EAGLCLIENT}
      Mul.Q.TQ.Seek(Mul.FListe.Row-1) ;
      TOB.Create('Liste Enreg', TOBEnregs, -1).SelectDB('',Mul.Q.TQ, true);
{$ELSE}
      TOB.Create('Liste Enreg', TOBEnregs, -1).SelectDB('', Mul.Q, true);
{$ENDIF}
      end;
    end;
//******* fin de traitement déselection *******
// PL le 10/02/03 : Repositionne les flags de modification à False pour que les putvalue éventuels les positionnent à true
// et que si un update suit, ne soient updatés que les champs réellement modifiés par les put !  enfin,... j'me comprends
TOBEnregs.SetAllModifie (false);

// Gestion de la déselection des lignes et des boutons allSelect...
FinTraiteMul (Mul);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 29/08/2001
Modifié le ... : 08/01/2002
Description .. : fonction de traitement des enregistrements d'un mul de
Suite ........ : traitement. Rend une tob (NomTable de la Tob = NomTob)
Suite ........ : basée sur le select stSelect
Suite ........ : le where est fait à partir de la liste des champs de ArgWhere
Suite ........ : ChampCpt et TableCpt utilisés pour comptage des enreg
Suite ........ : attention limite actuelle à 2000 enreg
Suite ........ : Diffère de EnregMulListe car dans ce cas, on attaque tous 
Suite ........ : les champs de la table pour charger la tob , au lieu d'avoir 
Suite ........ : une tob renseignée (dans MulListe) uniquement avec les 
Suite ........ : champs de la liste (les colonnes).
Suite ........ : Utiliser l'une ou l'autre en fct du traitement à faire.
Mots clefs ... : MUL
*****************************************************************}
FUNCTION TraiteEnregMulTable (Mul : TFMul; stSelect,ArgWhere, NomTob, ChampCpt, TableCpt: string;
                              TobEnregs : TOB; Parle : Boolean ; nbEnregMax : integer = 200000) : T_ErrTraiteEnreg;
Var stWhere : string;
    QQ : TQuery;
    i : integer;
begin
Result := tteOK;
if Mul = Nil then Exit;
if (TobEnregs = Nil) then Exit;

if Not(TestEnregAtraiter(Mul,Parle)) then
   begin Result := tteAucunEnreg; Exit; end;

{$IFDEF EAGLCLIENT}
if Mul.bSelectAll.Down then
   if not Mul.Fetchlestous then
     begin
     Mul.bSelectAllClick(Nil);
     Mul.bSelectAll.Down := False;
     exit;
     end;
{$ENDIF}
                                
if Mul.FListe.AllSelected then
   begin
   stWhere := ' ' + RecupWhereCritere(Mul.Pages);
   // Comptage pour blocage
   if Not (VerifNbrPourBlocage ( ChampCpt,TableCpt, stWhere , Parle ,nbEnregMax )) then
         begin Result := tteTropEnreg; Exit; end;

   QQ := OpenSQL (stSelect + stWhere, True);
   if not QQ.EOF then TobEnregs.LoadDetailDB( NomTob, '', '', QQ,False);
   Ferme(QQ) ;
   end
else
   begin
   for i := 0 to Mul.FListe.NbSelected-1 do
      begin
      Mul.FListe.GotoLeBookMark(i);
{$IFDEF EAGLCLIENT}
      Mul.Q.TQ.Seek (Mul.FListe.Row-1) ;
{$ENDIF}
      ChargeUnEnreg (Mul, stSelect,ArgWhere, NomTob, TobEnregs);
      end;
   end;

//******* fin de traitement déselection *******
// PL le 10/02/03 : Repositionne les flags de modification à False pour que les putvalue éventuels les positionnent à true
// et que si un update suit, ne soient updatés que les champs réellement modifiés par les put ! enfin,... j'me comprends
TOBEnregs.SetAllModifie (false);

// Gestion de la déselection des lignes et des boutons allSelect...
FinTraiteMul (Mul);
end;


end.
