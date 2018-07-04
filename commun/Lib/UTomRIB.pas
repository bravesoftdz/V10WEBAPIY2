{$A+,B-,C-,D+,E-,F-,G+,H+,I+,J+,K-,L+,M-,N+,O-,P+,Q+,R+,S-,T-,U-,V+,W+,X+,Y+,Z1}
{$MINSTACKSIZE $00004000}
{$MAXSTACKSIZE $00100000}
{$IMAGEBASE $00400000}
{$APPTYPE GUI}
unit UTomRIB;

interface

uses
  Classes,
  UTOM,UTOB,
  {$IFDEF EAGLCLIENT}
  eFichList,
  messages,
  windows,
  Forms,
  {$ELSE}
  db,
  {$IFNDEF DBXPRESS}dbtables {BDE}, {$ELSE}uDbxDataSet, {$ENDIF}
  Fiche,
  FichList,
  Fe_Main,
  {$ENDIF}
  HCtrls,
  Ent1,
  HMsgBox,
  UtilPGI,
  Controls,
  HEnt1,
  SysUtils,
  HTB97,
  ULibIdentBancaire,
  wCommuns,
  Hqry,
  HDB,uEntCommun
  ;

type
  Tom_RIB = class(TOM)
  private
    //CodePays : THDBEdit;
{$IFDEF EAGLCLIENT}
   CodePays : THValComboBox;
   Fliste   : THGrid;
{$ELSE}
   CodePays :  THDBValComboBox ;
   Fliste   : THDBGrid;
{$ENDIF}
    codeISO, codeTypePays: string;
    Auxiliaire: string;
    StRIB: string;
    Origine: string;
    FromSaisie: boolean;
    FirstTime: boolean;
    gbNoIban: boolean;
    gbChargement: boolean;
    cbo_Pays: THValComboBox;
    {$IFDEF AFFAIRE} // BDU - 19/04/07 - PDev : 5063
    fboAvecMaj: Boolean;
    ReferenceMajAuto: String;
    {$ENDIF}
    procedure R_PAYSChange(Sender: TObject);
    procedure R_TYPEPAYSChange(Sender : TObject);
    procedure R_NUMEROCOMPTEKeyPress(Sender: TObject; var Key: Char);
    procedure OnKeyPressR_CODEBIC(Sender: TObject; var Key: Char);
    procedure AfficheSpecifES;
    function RIBRenseigner: Boolean;
    {$IFDEF AFFAIRE} // BDU - 19/04/07 - PDev : 5063
    procedure SauveRib(Vide: Boolean);
    {$ENDIF}
  	procedure FlisteDblClick (Sender : TObject);
  public
    procedure OnNewRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
    procedure AfficheEcran;
    procedure VerifIBAN;
    function VerifiRibBic(TypeCompte : string): Boolean;
    function PrincipalUnique: Boolean;
    function TypeRibUnique(TypeRib: string): Boolean; // SB 17/09/2003
  end;

const
  // libellés des messages
  TexteMessage: array[1..21] of string = (
    {1}'Vous devez renseigner la domiciliation.',
    {2}'La domiciliation comporte des caractères interdits.',
    {3}'3',
    {4}'4',
    {5}'5',
    {6}'Vous devez renseigner la banque.',
    {7}'Vous devez renseigner le guichet.',
    {8}'Vous devez renseigner le numéro de compte.',
    {9}'Vous devez renseigner la clé du RIB.',
    {10}'Vous devez renseigner le code BIC.',
    {11}'Ce numero de RIB ne peut pas être le principal car il n''est pas unique.',
    {12}'Confirmez-vous le nouveau RIB principal ?',
    {13}'La clé RIB est erronnée. Souhaitez-vous la recalculer ?',
    {14}'Confirmez-vous le nouveau RIB salaire ?', // SB 17/09/2003
    {15}'Confirmez-vous le nouveau RIB acompte ?',
    {16}'Confirmez-vous le nouveau RIB frais professionnel ?',
    {17}'17',
    {18}'La clé du code IBAN est erronée. Voulez-vous la recalculer ?',
    {19}'Vous devez renseigner le code IBAN.',
    {20}'Vous devez renseigner le type d''identification.',
    {21}'La longueur minimale du code IBAN est de 10 caractères.'
    );

implementation

uses
{$IFDEF MODENT1}
  CPTypeCons,
  CPProcGen,
{$ENDIF MODENT1}
{$IFDEF AFFAIRE} // BDU - 19/04/07 - PDev : 5063
  UtofAfMajChampsAuto, // dans \gescom\liba
{$ENDIF}
  uTobDebug, TntDBCtrls;

{$IFDEF AFFAIRE} // BDU - 19/04/07 - PDev : 5063
type
  TSauveRib = record
    Principal,
    Banque,
    Guichet,
    Compte,
    Cle,
    Domiciliation: String;
  end;

var
  VSauveRib: TSauveRib;
{$ENDIF}

{$IFDEF AFFAIRE}
// BDU - 19/04/07 - PDev : 5063
procedure Tom_RIB.SauveRib(Vide: Boolean);
begin
  with VSauveRib do
  begin
    // Nettoi les variables de sauvegarde
    if Vide then
    begin
      Principal := '';
      Banque := '';
      Guichet := '';
      Compte := '';
      Cle := '';
      Domiciliation := '';
    end
    else
    begin
      if FromSaisie then
      begin
        Banque := GetControlText('R_ETABBQ');
        Guichet := GetControlText('R_GUICHET');
        Compte := GetControlText('R_NUMEROCOMPTE');
        Cle := GetControlText('R_CLERIB');
        Domiciliation := GetControlText('R_DOMICILIATION');
        Principal := GetControlText('R_PRINCIPAL');
      end
      else
      begin
        Banque := GetField('R_ETABBQ');
        Guichet := GetField('R_GUICHET');
        Compte := GetField('R_NUMEROCOMPTE');
        Cle := GetField('R_CLERIB');
        Domiciliation := GetField('R_DOMICILIATION');
        Principal := GetField('R_PRINCIPAL');
      end;
    end;
  end;
end;
{$ENDIF}

procedure Tom_RIB.OnCLose;
begin
  inherited;
  DestroySaveAffichage ;
end;

procedure Tom_RIB.OnArgument(Arguments: string);
var
  x: integer;
  critere: string;
  Arg, Val: string;
{$IFNDEF EAGLCLIENT}
  txt_Iban, txt_Num : THDBEdit;
  R_TYPEPAYS : THDBValComboBox ;
{$ELSE}
  txt_Iban, txt_Num : THEdit;
  R_TYPEPAYS : THValComboBox ;
{$ENDIF}
begin
  inherited;
  {$IFDEF AFFAIRE} // BDU - 19/04/07 - PDev : 5063
  fboAvecMaj := False;
  {$ENDIF}
  // On sauvegarde les positions des THDBEdit par defaut afin de les réutiliser plus tard.
  SaveAffichageDefaut(ecran,'R_');
  FromSaisie := True;
  repeat
    gbNoIban := True;
    Critere := uppercase(Trim(ReadTokenSt(Arguments)));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        Arg := copy(Critere, 1, x - 1);
        Val := copy(Critere, x + 1, length(Critere));
        if Arg = 'NUMAUX' then Auxiliaire := Val;
        if Arg = 'ORIGINE' then Origine := Val;
        if Arg = 'FROMSAISIE' then
        begin
          if Val = 'X' then FromSaisie := True else FromSaisie := False;
        end;
        if Arg = 'STRIB' then StRIB := Val;
           //if Arg='ISAUX' then
        {$IFDEF AFFAIRE}
        if CompareText(Arg, 'OBJ') = 0 then
          ReferenceMajAuto := Val;
        {$ENDIF}
      end
      {$IFDEF AFFAIRE} // BDU - 19/04/07 - PDev : 5063
      else if CompareText(Critere, 'MAJ') = 0 then
        fboAvecMaj := True;
      {$ENDIF}
    end;
  until Critere = '';

//  BBI correction bug sur mapping fichier txt
    if Origine = 'PGISIDE' then
      FromSaisie := False;
//  BBI Fin correction bug sur mapping fichier txt
  // Pour gerer de l'elipsis du type de pays
  //CodePays := THDBEdit(GetControl('R_TYPEPAYS', true));
  //SDA le 14/03/2007
{$IFNDEF EAGLCLIENT}
    CodePays := THDBValComboBox(GetControl('R_TYPEPAYS'));
    cbo_Pays := THDBValComboBox(GetControl('R_PAYS'));
    Fliste   := THDBGrid(GetControl('FLISTE'));
{$ELSE}
   CodePays := THValComboBox(GetControl('R_TYPEPAYS'));
   cbo_Pays := THValComboBox(GetControl('R_PAYS'));
   Fliste   := THGrid(GetControl('FLISTE'));
{$ENDIF}
  //Fin SDA le 14/03/2007

  if FromSaisie then SetControlChecked('FROMSAISIE', true);
    FirstTime := true;
  if ctxPaie in V_PGI.PGIContexte then
  begin
    SetControlVisible('R_SALAIRE', True);
    SetControlVisible('R_ACOMPTE', True);
    //VERSION PAIE S3
    {$IFNDEF CCS3}
    SetControlVisible('R_FRAISPROF', True);
    {$ENDIF}
    //FIN VERSION PAIE S3
  end;

  // Code IBAN

{$IFNDEF EAGLCLIENT}
  txt_Iban := THDBEdit(GetControl('R_CODEIBAN'));
{$ELSE}
  txt_Iban := THEdit(GetControl('R_CODEIBAN'));
{$ENDIF}

  if (txt_Iban <> nil) then
  begin
    txt_Iban.OnKeyPress := R_NUMEROCOMPTEKeyPress;
    // txt_Iban.OnExit := R_CODEIBANOnExit ;   // On ne verifie plus l'Iban
  end;

{$IFNDEF EAGLCLIENT}
  txt_Num := THDBEdit(GetControl('R_NUMEROCOMPTE'));
{$ELSE}
  txt_Num := THEdit(GetControl('R_NUMEROCOMPTE'));
{$ENDIF}

  if (txt_Num <> nil) then txt_Num.OnKeyPress := R_NUMEROCOMPTEKeyPress;

  {$IFNDEF EAGLCLIENT}
    R_TYPEPAYS := THDBValComboBox(GetControl('R_TYPEPAYS'));
  {$ELSE}
    R_TYPEPAYS := THValComboBox(GetControl('R_TYPEPAYS'));
  {$ENDIF}
  if (R_TYPEPAYS <> nil) then R_TYPEPAYS.OnChange := R_TYPEPAYSChange;

  {FQ 19473 BV 11/01/2007
  if VH^.PaysLocalisation<>CodeISOES then
  begin}
 //   cbo_Pays := THValComboBox(GetControl('R_PAYS'));
 //   if (cbo_Pays <> nil) then cbo_Pays.OnChange := R_PAYSChange;
{  End ; //XVI 24/02/2005 }

  if Origine = 'PGISIDE' then Exit;
{$IFNDEF EAGLCLIENT}
       THDBEdit(GetControl('R_CODEBIC', true)).OnKeyPress := OnKeyPressR_CODEBIC;
{$ELSE}
       THEdit(GetControl('R_CODEBIC', true)).OnKeyPress := OnKeyPressR_CODEBIC;
{$ENDIF}


  //SG6 07.02.05 FQ 12454
  SetFocusControl('R_PRINCIPAL');
	Fliste.OnDblClick := FlisteDblClick;
end;

procedure Tom_RIB.OnLoadRecord;
var
    Etab, Guichet, NumCompte, Cle, Dom, codePaysLocal : string;
    TRIB : THTable;
begin

  gbChargement := True;

  if ds.State <> dsBrowse then
  begin

  if FromSaisie and FirstTime then
  begin
    FirstTime := false;
    DecodeRIB(Etab, Guichet, NumCompte, Cle, Dom, StRIB,codeISODuPAys(GetField('R_PAYS'))); //Attention au Code ISOPays //XVI 24/02/2005
    NumCompte := Trim(NumCompte); //Attention au Code ISOPays
    TRIB := TFFicheliste(Ecran).Ta;
    if TRIB.State = dsBrowse then
    begin
      //TRIB.Locate('R_NUMEROCOMPTE',NumCompte,[]) ;
      while not TRIB.EOF do if uppercase(TRIB.FindField('R_NUMEROCOMPTE').AsString) = uppercase(NumCompte) then Break else TRIB.Next;
    end;
  end;

  end;

  if (GetControlText('R_PAYS') = '') then
    if RibRenseigner then
    begin
      { ajout mng 28/03/2003 }
      if ds.State = dsBrowse then
        DS.Edit;
      codePaysLocal := CodePaysDeIso(VH^.PaysLocalisation); //XVI 24/02/2005
      SetField('R_PAYS', codePaysLocal);
      if VH^.PaysLocalisation <> CodeISOES then
        cbo_Pays.Value := codePaysLocal; //XVI 24/02/2005
      PGIBox('Votre code pays a été modifié. Mise à jour effectué', 'Attention');
    end;
  // Se positionne sur PRINCIPAL
  //SetFocusControl('R_PRINCIPAL');

  // Permet le non calcul de l'iban
  gbNoIban := True;
  gbChargement := False;


  // Affichage :
  // Gestion differente selon si on change de pays ou non.


  if (codeISO <> codeISODuPAys(GetControlText('R_PAYS'))) then
  begin
    R_PAYSChange(nil);
    AfficheEcran; // SDA le 15/03/2007 mauvais affichage en fonction type identification
  end
  else
  begin
    R_PAYSChange(nil); //SDA le 15/03/2007 mauvais affichage en fonction type identification
    AfficheEcran;
  end;


  {$IFDEF EAGLCLIENT}
    sendmessage(Fliste.Handle, WM_LBUTTONUP, 0,0);
  {$ENDIF}

  {$IFDEF AFFAIRE} // BDU - 19/04/07 - PDev : 5063
  if fboAvecMaj then
    SauveRib(False);
  {$ENDIF}


end;

procedure Tom_RIB.OnUpdateRecord;

var
  RibCalcul: string;
  StDom, stParam, sType, szIBAN, strType : string;
  R_ETABBQ, R_GUICHET, R_NUMEROCOMPTE, R_CLERIB: string;

  function IsPrincipal: boolean;
  begin
    if FromSaisie then
      Result := (GetControlText('R_PRINCIPAL') = 'X')
    else
      Result := (GetField('R_PRINCIPAL') = 'X');
  end;

begin
  inherited;
  sType := quelType(codeISO,codeTypePays);
  // Vérifie si un RIB "principal" existe : Si oui demande confirmation
  // BPY le 04/03/2004 => fiche n°13245
  if ((not PrincipalUnique) and IsPrincipal) then
  begin
    if FromSaisie then
    begin
      if PGIAsk(TexteMessage[12], 'RIB') = mrYes then
        // Rend l'autre RIB non principal
        ExecuteSQL('UPDATE RIB SET R_PRINCIPAL="-" WHERE R_AUXILIAIRE="' + Auxiliaire + '" AND R_PRINCIPAL="X" AND R_NUMERORIB<>' + GetControlText('R_NUMERORIB'))
      else
      begin
        // Décoche PRINCIPAL
        SetControlText('R_PRINCIPAL', '-');
        SetField('R_PRINCIPAL', '-');
      end;
    end
    else
    begin
      if IsPrincipal then
        // Rend l'autre RIB non principal
        ExecuteSQL('UPDATE RIB SET R_PRINCIPAL="-" WHERE R_AUXILIAIRE="' + Auxiliaire + '" AND R_PRINCIPAL="X" AND R_NUMERORIB<>' + GetControlText('R_NUMERORIB'));
    end;
  end;
  // fin BPY

  //SDA le 14/03/2007
  if GetControlVisible('R_TYPEPAYS') = true then
  begin
   if FromSaisie  then
    StrType := GetControlText('R_TYPEPAYS')
   else
    StrType := GetField('R_TYPEPAYS');
   if StrType = '' then
   begin
    SetControlText('R_TYPEPAYS', CodePays.Values[0]);
    SetFocusControl('R_TYPEPAYS');
    LastError := 20;
    LastErrorMsg := TexteMessage[LastError];
    exit;
   end;
   //SetControlText('R_TYPEPAYS', CodePays.Values[0]);
   if strType <> '' then
   SetControlText('R_TYPEPAYS', CodePays.Values[CodePays.itemindex]);
   //SetField('R_TYPEPAYS',CodePays.Values[CodePays.itemindex]);
  end;

  //Fin SDA le 14/03/2007
  if FromSaisie then
    StDom := GetControlText('R_DOMICILIATION')
  else
    StDom := GetField('R_DOMICILIATION');
  if StDom = '' then
  begin
    SetFocusControl('R_DOMICILIATION');
    LastError := 1;
    LastErrorMsg := TexteMessage[LastError];
    exit;
  end;

  if ExisteCarInter(StDom) then
  begin
    SetFocusControl('R_DOMICILIATION');
    LastError := 2;
    LastErrorMsg := TexteMessage[LastError];
        exit;
  end;
  //ecran.FormState ;
  if not VerifiRibBic(sType) then exit ;

  // BPY le 04/03/2004 => fiche n°13245
  if ((not PrincipalUnique) and IsPrincipal) then
  begin
    SetFocusControl('R_PRINCIPAL');
    LastError := 11;
    LastErrorMsg := TexteMessage[LastError];
    exit;
  end;
  // Fin BPY

  {JP 17/01/006 : FQ 17310 : ajout du test sur VH^.PaysLocalisation, car si le pays est l'Espagne,
                  mais qu'il ne s'agit pas d'une compta espagnole, les zones R_ETABBQ, R_GUICHET ...
                  sont désactivées}
  //RR if (VH^.CtrlRIB) and (RibRenseigner) and
  if (VH^.CtrlRIB) and
    ((codeIso = CodeISOFR) or ((codeIso = CodeISOES) and
    (VH^.PaysLocalisation = CodeISOES))) then
  begin
    if FromSaisie then
    begin
      R_ETABBQ := GetControlText('R_ETABBQ');
      R_GUICHET := GetControlText('R_GUICHET');
      R_NUMEROCOMPTE := GetControlText('R_NUMEROCOMPTE');
      R_CLERIB := GetControlText('R_CLERIB');
    end
    else
    begin
      R_ETABBQ := Trim(GetField('R_ETABBQ'));
      R_GUICHET := Trim(GetField('R_GUICHET'));
      R_NUMEROCOMPTE := Trim(GetField('R_NUMEROCOMPTE'));
      R_CLERIB := Trim(GetField('R_CLERIB'));
    end;

    RibCalcul := VerifRib(R_ETABBQ, R_GUICHET, R_NUMEROCOMPTE, CodeISO); //XVI 24/02/2005
    if (RibCalcul <> Trim(R_CLERIB)) then
    begin
      if FromSaisie then
      begin
        if PGIAsk(TexteMessage[13], 'RIB') <> mrYes then
        begin
          SetFocusControl('R_CLERIB');
          SetField('R_CODEIBAN', ''); //SetControlText('R_CODEIBAN','');
          LastError := 3;
          Exit;
        end
        else
        begin
          { ajout mng 28/03/2003 }
          if ds.State = dsBrowse then
            DS.Edit;
          SetField('R_CLERIB', RibCalcul);
          if VH^.PaysLocalisation <> CodeISOES then
          begin
            SetField('R_CODEIBAN', '');
          end; //XVI 24/02/2005
        end;
      end
      else
     begin
        {b thl FQ 17466 22/05/2006}
        if PGIAsk(TexteMessage[13], 'RIB') = mrYes then
        begin
          SetField('R_CLERIB', RibCalcul);
          R_CLERIB := Trim(GetField('R_CLERIB'));
          LastErrorMsg := '';
        end;
        {e thl FQ 17466 22/05/2006}
        SetFocusControl('R_CLERIB');
        SetField('R_CODEIBAN', '');
        LastError := 3;
        Exit;
      end;
    end;
  end;
  // SB 17/09/2003
  {$IFDEF PAIEGRH}
  { Traitement du Rib salaire }
  if not TypeRibUnique('R_SALAIRE') then
    if PGIAsk(TexteMessage[14], Ecran.Caption) = mrYes then
      // Rend l'autre RIB non salaire
      ExecuteSQL('UPDATE RIB SET R_SALAIRE="-" WHERE R_AUXILIAIRE="' + Auxiliaire + '" ' +
        'AND R_SALAIRE="X" AND R_NUMERORIB<>' + GetControlText('R_NUMERORIB'))
    else
      SetField('R_SALAIRE', '-'); { Décoche Salaire }
  { Traitement du Rib Acompte  }
  if not TypeRibUnique('R_ACOMPTE') then
    if PGIAsk(TexteMessage[15], Ecran.Caption) = mrYes then
      // Rend l'autre RIB non acompte
      ExecuteSQL('UPDATE RIB SET R_ACOMPTE="-" WHERE R_AUXILIAIRE="' + Auxiliaire + '" ' +
        'AND R_ACOMPTE="X" AND R_NUMERORIB<>' + GetControlText('R_NUMERORIB'))
    else
      SetField('R_ACOMPTE', '-'); { Décoche acompte }
  { Traitement du Rib Frais professionnel }
  if not TypeRibUnique('R_FRAISPROF') then
    if PGIAsk(TexteMessage[16], Ecran.Caption) = mrYes then
      // Rend l'autre RIB non Frais professionnel
      ExecuteSQL('UPDATE RIB SET R_FRAISPROF="-" WHERE R_AUXILIAIRE="' + Auxiliaire + '" ' +
        'AND R_FRAISPROF="X" AND R_NUMERORIB<>' + GetControlText('R_NUMERORIB'))
    else
      SetField('R_FRAISPROF', '-'); { Décoche Frais professionnel }
  {$ENDIF}

  // Verification du RIB
  if (sType = 'RIB') or
  ((sType = '') and (codeISO = CodeISOFR) and (VH^.PaysLocalisation = CodeISOFR) ) or
  ((sType = '') and (codeISO = CodeISOES) and (VH^.PaysLocalisation = CodeISOES) ) then
  begin
  	// FQ 19489 BV  11/01/2007
    // On verifie l'IBAN car un des codes du RIB peut avoir été modifié sans qu'on modifie l'IBAN
	  stParam := calcRIB(codeISO, GetField('R_ETABBQ'), GetField('R_GUICHET'), GetField('R_NUMEROCOMPTE'), GetField('R_CLERIB'));
    if stParam <> '' then
    	szIban := calcIBAN(codeISO, stParam);
    stParam := Trim(THDBEdit(getcontrol('R_CODEIBAN')).Text);
    if (stParam <> szIban) then
  	 begin
    	SetControlText('R_CODEIBAN','');
      VerifIBAN;                
      stParam := Trim(THDBEdit(getcontrol('R_CODEIBAN')).Text);
    end;
    // END FQ 19489
    if ErreurDansIban(stParam)then
      if PGIAsk(TraduireMemoire(TexteMessage[18]), 'RIB') = mrYes then
      begin
        SetControlText('R_CODEIBAN','');
        VerifIBAN;
      end;
  end
  else if (sType = 'IBA') or (sType = '') then
  begin
    // verification IBAN
    stParam := Trim(THDBEdit(getcontrol('R_CODEIBAN')).Text);
    { FQ BV 01/02/2007 }
    if Trim(stParam) = '' then
    begin
       LastError := 19;
       LastErrorMsg := TexteMessage[LastError];
       SetFocusControl('R_CODEIBAN');
       Exit;
    end
    //SDA le 14/03/2007
    else
    if length(stParam) < 10 then
    begin
       LastError := 21;
       LastErrorMsg := TexteMessage[LastError];
       SetFocusControl('R_CODEIBAN');
       Exit;
    end
    //Fin SDA le 14/03/2007
    { END FQ }
    else if ErreurDansIban(stParam) then
      if PGIAsk(TraduireMemoire(TexteMessage[18]), 'RIB') = mrYes then
      begin
        VerifIBAN;
      end;
  end;

end;

function Tom_RIB.VerifiRibBic(TypeCompte : string) : Boolean;
var R_ETABBQ, R_GUICHET, R_NUMEROCOMPTE, R_CLERIB, R_CODEBIC : string;
begin
	Result := False;
	if FromSaisie then
  begin
    R_ETABBQ := GetControlText('R_ETABBQ');
    R_GUICHET := GetControlText('R_GUICHET');
    R_NUMEROCOMPTE := GetControlText('R_NUMEROCOMPTE');
    R_CLERIB := GetControlText('R_CLERIB');
    R_CODEBIC := GetControlText('R_CODEBIC');
  end
  else
  begin
    R_ETABBQ := Trim(GetField('R_ETABBQ'));
    R_GUICHET := Trim(GetField('R_GUICHET'));
    R_NUMEROCOMPTE := Trim(GetField('R_NUMEROCOMPTE'));
    R_CLERIB := Trim(GetField('R_CLERIB'));
    R_CODEBIC := Trim(GetField('R_CODEBIC'));
  end;

  {JP 17/01/006 : FQ 17310 : ajout du test sur VH^.PaysLocalisation, car si le pays est l'Espagne,
                  mais qu'il ne s'agit pas d'une compta espagnole, les zones R_ETABBQ, R_GUICHET ...
                  sont désactivées : on ne teste donc pas les zones du RIB}
    // Cas RIB FQ 19583 BV 01/02/2007
    if (TypeCompte = 'RIB') or
       ((TypeCompte = '') and (codeISO = CodeISOFR) and (VH^.PaysLocalisation = CodeISOFR)) or
       ((TypeCompte = '') and (codeISO = CodeISOES) and (VH^.PaysLocalisation = CodeISOES)) then
  begin
    if (R_ETABBQ = '') and (chercheLONG('ETABBQ',codeISO,GetControlText('R_TYPEPAYS')) <> 0) then
    begin
      SetFocusControl('R_ETABBQ');
      LastError := 6;
      LastErrorMsg := TexteMessage[LastError];
      exit;
    end;
    if (R_GUICHET = '') and (chercheLONG('GUICHET',codeISO,GetControlText('R_TYPEPAYS')) <> 0) then
    begin
      SetFocusControl('R_GUICHET');
      LastError := 7;
      LastErrorMsg := TexteMessage[LastError];
      exit;
    end;
    if (R_NUMEROCOMPTE = '') and (chercheLONG('NUMEROCOMPTE',codeISO,GetControlText('R_TYPEPAYS')) <> 0) then
    begin
      SetFocusControl('R_NUMEROCOMPTE');
      LastError := 8;
      LastErrorMsg := TexteMessage[LastError];
      exit;
    end;
    if (R_CLERIB = '') and (chercheLONG('CLERIB',codeISO,GetControlText('R_TYPEPAYS')) <> 0) then
    begin
      R_CLERIB := VerifRib(R_ETABBQ, R_GUICHET, R_NUMEROCOMPTE, CodeISO);
      SetFocusControl('R_CLERIB');
      LastError := 9;
      LastErrorMsg := TexteMessage[LastError];
      exit;
    end;
  end
  else
  begin
    if R_CODEBIC <> '' then
    begin
    	Result := True;
      Exit;
    end
    else if (codeISO <> VH^.PaysLocalisation) then // FQ 19617 BV 01/02/2007
    begin
     SetFocusControl('R_CODEBIC');
     LastError := 10;
     LastErrorMsg := TexteMessage[LastError];
     exit;
    end;
  end;
  Result := True;
end;

function Tom_RIB.PrincipalUnique: Boolean;
var
  Q: TQuery;
begin
    Result := True;
    if (GetControlText('R_NUMERORIB') = '') then exit;
    Q := OpenSql('Select R_PRINCIPAL from RIB Where R_AUXILIAIRE="' + Auxiliaire + '" And ' +
        'R_PRINCIPAL="X" And R_NUMERORIB<>' + GetControlText('R_NUMERORIB') + '', True);
    if not Q.Eof then Result := False;
    Ferme(Q);
end;

procedure Tom_RIB.OnNewRecord;
var
  Q: TQuery;
  LastNum: Longint;
  strDevise : string; //SDA le 27/12/2007
begin
  inherited;

  // BPY le 04/03/2004 => fiche n°13245
  if (PrincipalUnique) then
  begin
    // Rend le RIB "principal" par défaut
    SetControlChecked('R_PRINCIPAL', True);
    SetField('R_PRINCIPAL', 'X');
  end
  else
  begin
    SetControlChecked('R_PRINCIPAL', false);
    SetField('R_PRINCIPAL', '-');
  end;
    // fin BPY

    Q := OpenSQL('SELECT Max(R_NumeroRib) FROM RIB WHERE R_AUXILIAIRE ="' + Auxiliaire + '"', TRUE);
    if not Q.EOF then LastNum := Q.Fields[0].AsInteger
    else LastNum := 0;
    ferme(Q);

    SetField('R_AUXILIAIRE', Auxiliaire);
    SetField('R_NUMERORIB', LastNum + 1);
    SetField('R_PAYS', CodePaysDeIso(VH^.PaysLocalisation)); //XVI 24/02/2005
    SetField('R_TYPEPAYS', '');

{$IFNDEF EAGLCLIENT}
  //SDA le 27/12/2007 SetFocusControl('R_DOMICILIATION');
{$ENDIF}
  // Affichage :
  //SDA le 14/03/2007 AfficheEcran;
  R_PAYSChange(nil);
  {$IFDEF AFFAIRE} // BDU - 19/04/07 - PDev : 5063
  if fboAvecMaj then
    SauveRib(True);
  {$ENDIF}

  //SDA le 27/12/2007  - devise positionnée par défaut
  strDevise := V_PGI.DevisePivot;
  if strDevise <> '' then
    SetField ('R_DEVISE', strDevise);
  SetFocusControl('R_VILLE');
  {$IFDEF EAGLCLIENT}
  postmessage(THEdit(GetControl('R_PRINCIPAL')).Handle,WM_SETFOCUS,0,0);
  postmessage(THEdit(GetControl('R_PRINCIPAL')).Handle,WM_KEYDOWN, VK_TAB, 0) ;
  postmessage(THEdit(GetControl('R_DEVISE')).Handle,WM_KEYDOWN, VK_TAB, 0) ;
  //PostMessage(THEdit(GetControl('R_VILLE')).Handle, WM_KEYDOWN, VK_TAB, 0) ;
  //PostMessage(THEdit(GetControl('R_VILLE')).Handle, WM_KEYDOWN, VK_LEFT, 0) ;
  {$ENDIF EAGLCLIENT}
  //Fin SDA le 27/12/2007
end;


procedure Tom_RIB.VerifIBAN;
var
  szIban, sParam, sIBAN, sType: string;
begin
    if FromSaisie then
    begin
    	R_PAYSChange(nil);
      codeTypePays := GetControlText('R_TYPEPAYS') ;
      sIban := Trim(THDBEdit(getcontrol('R_CODEIBAN')).Text);
    end
    else
    begin
      codeISO := codeISODuPays(GetField('R_PAYS'));
      codeTypePays := GetField('R_TYPEPAYS');
      sIban := GetField('R_CODEIBAN');
    end;
    // Pas de code iban ou iban incorrect : Le calcul
    // Gestion selon les differents type de compte :
    sType := quelType(codeISO,codeTypePays);
    if (sType = 'RIB') or
      ((sType = '') and (codeISO = CodeISOFR) and (VH^.PaysLocalisation = CodeISOFR) ) or
      ((sType = '') and (codeISO = CodeISOES) and (VH^.PaysLocalisation = CodeISOES) ) then
      // RIB
      sParam := calcRIB(codeISO, GetField('R_ETABBQ'), GetField('R_GUICHET'), GetField('R_NUMEROCOMPTE'), GetField('R_CLERIB'))
    else if (sType = 'IBA') or (sType = '') then
    begin
      // IBAN
      if length(sIban) > 5 then
        sParam := copy(sIban ,5,length(sIban) - 4);
    end;
    if sParam <> '' then
      szIban := calcIBAN(codeISO, sParam);
    SetField('R_CODEIBAN', szIban);

end;

procedure Tom_RIB.OnChangeField(F: TField);
begin
  inherited;
  if F.FieldName = 'R_PAYS' then
     R_PAYSChange(nil)
  else if F.FieldName = 'R_TYPEPAYS' then
     R_TYPEPAYSChange(nil); 

end;

procedure Tom_RIB.OnAfterUpdateRecord;
{$IFDEF AFFAIRE}
var
  R_Banque,
  R_Guichet,
  R_Compte,
  R_Cle,
  R_Domiciliation,
  R_Principal: String;  
{$ENDIF}
begin
  inherited;

  {$IFDEF AFFAIRE} // BDU - 19/04/07 - PDev : 5063
  if fboAvecMaj then
  begin
    if FromSaisie then
    begin
      R_Banque := GetControlText('R_ETABBQ');
      R_Guichet := GetControlText('R_GUICHET');
      R_Compte := GetControlText('R_NUMEROCOMPTE');
      R_Cle := GetControlText('R_CLERIB');
      R_Domiciliation := GetControlText('R_DOMICILIATION');
      R_Principal := GetControlText('R_PRINCIPAL');
    end
    else
    begin
      R_Banque := GetField('R_ETABBQ');
      R_Guichet := GetField('R_GUICHET');
      R_Compte := GetField('R_NUMEROCOMPTE');
      R_Cle := GetField('R_CLERIB');
      R_Domiciliation := GetField('R_DOMICILIATION');
      R_Principal := GetField('R_PRINCIPAL');
    end;
    if (VSauveRib.Compte <> R_Compte) or (VSauveRib.Guichet <> R_Guichet) or
      (VSauveRib.Banque <> R_Banque) or (VSauveRib.Cle <> R_Cle) or
      (VSauveRib.Domiciliation <> R_Domiciliation) or (VSauveRib.Principal <> R_Principal) then
    MajAutoChangeValeur('T_RIB', 'X', ReferenceMajAuto);
  end;
  {$ENDIF}
  if VH^.PaysLocalisation <> CodeISOES then
    gbChargement := True;
  RefreshDB;
  if VH^.PaysLocalisation <> CodeISOES then //XVI 24/02/2005
    gbChargement := False;
end;

procedure Tom_RIB.R_NUMEROCOMPTEKeyPress(Sender : TObject; var Key : Char);
begin
    // Authorise uniquement les caractères A-Z et 0-9 et Backspace
    if not (Key in ['a'..'z', 'A'..'Z', '0'..'9', #8]) then Key := #0;
end;


function Tom_RIB.RIBRenseigner : Boolean;
begin
  if FromSaisie then
    Result := ((Trim(GetControlText('R_ETABBQ')) <> '') and
      (Trim(GetControlText('R_GUICHET')) <> '') and
      (Trim(GetControlText('R_NUMEROCOMPTE')) <> ''))
  else
    Result := ((Trim(GetField('R_ETABBQ')) <> '') and
      (Trim(GetField('R_GUICHET')) <> '') and
               (Trim(GetField('R_NUMEROCOMPTE')) <> ''));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/10/2003
Modifié le ... :   /  /
Description .. : empeche de saisie des espace dans la zone BIC du RIB
Suite ........ : fiche de bug n°12959
Mots clefs ... :
*****************************************************************}

procedure Tom_RIB.OnKeyPressR_CODEBIC(Sender : TObject; var Key : Char);
begin
    if (Key = ' ') then Key := #0;
end;

procedure Tom_RIB.R_TYPEPAYSChange(Sender: TObject);
var stTypePays : string;
begin

  if FromSaisie then
  begin
    stTypePays := GetControlText('R_TYPEPAYS');
    if codeIso <> codeISODuPays(GetControlText('R_PAYS')) then
    begin
    	R_PAYSChange(nil);
      exit;
    end;
  end
  else
  begin
    stTypePays := GetField('R_TYPEPAYS');
    if codeIso <> codeISODuPays(GetField('R_PAYS')) then
    begin
    	R_PAYSChange(nil);
      exit;
    end;
  end;
  
  // Si le pays n'ai pas renseigné pas la peine de continuer.
  if codeIso = '' then exit;

  if CodeTypePays <> stTypePays then
  begin
    CodeTypePays := stTypePays ;
    if ds.State in [dsInsert, dsEdit] then
    begin
      // Affichage :
      AfficheEcran;
      SetField('R_ETABBQ', '');
      SetField('R_GUICHET', '');
      SetField('R_NUMEROCOMPTE', '');
      SetField('R_CLERIB', '');
      //if (((codeISO = codeISOES) and (VH^.PaysLocalisation = CodeISOES))) or (codeISO = CodeISOFR) or (quelType(codeISO,GetControlText('R_TYPEPAYS')) = 'BBA') then
      //RR 15/03/2007 FQ 19833 .... à contre coeur...
      if (quelType(codeISO,GetControlText('R_TYPEPAYS')) = 'BBA') then
        SetField('R_CODEIBAN', '')
      else
        SetField('R_CODEIBAN', codeISO);
    end;
  end;
  // MAJ du libelle
  //THLabel(ecran.FindComponent('LR_TYPEPAYS')).Caption := GetLibelleParamPays(codeIso,CodeTypePays);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/11/2006
Modifié le ... :   /  /
Description .. : Affichage Specifique pour l'espagne
Mots clefs ... : 
*****************************************************************}
procedure Tom_RIB.AfficheSpecifES;
var
  ZoneRIB, ZoneCP : THDBEdit;
  LblZoneRib : THLabel;
  ZoneLeft, OrdreTab : Integer;
begin
  // On fait un affichage defaut pour s'assurer que les champs ont bien le formatage attendu.
 AfficheDefault(ecran,'R_',codeISO, true) ;
  // On desactive l'IBAN FQ 19465 BV 11/01/2007
  SetControlEnabled('R_CODEIBAN',false);
  if ds.State in [dsInsert, dsEdit] then
    begin
      SetField('R_ETABBQ', '');
      SetField('R_GUICHET', '');
      SetField('R_NUMEROCOMPTE', '');
      SetField('R_CLERIB', '');
    end;

    if (codeISO = CodeISOFR) or (CodeISO='') or (CodeISO = CodeISOES) then
    begin
      ZoneRIB := THDBEdit(GetControl('R_ETABBQ'));
      if Assigned(ZoneRIB) then
      begin
        ZoneRib.Enabled := (TFFicheListe(Ecran).TypeAction <> taConsult);
        ZoneRib.MaxLength := 4 ;
        ZoneRib.EditMask := FormatZonesRib('FR', 'BQ');
        {$IFNDEF EAGLCLIENT}
          ZoneRib.Field.EditMask := FormatZonesRib(CodeISO, 'BQ');
        {$ENDIF}
      end;

      ZoneRIB := THDBEdit(GetControl('R_GUICHET'));
      if Assigned(ZoneRIB) then
      begin
        ZoneRib.Enabled := (TFFicheListe(Ecran).TypeAction <> taConsult);
        ZoneRib.MaxLength := 4 ;
        ZoneRib.EditMask := FormatZonesRib(CodeISO, 'GU');
        {$IFNDEF EAGLCLIENT}
          ZoneRib.Field.EditMask := FormatZonesRib(CodeISO, 'GU');
        {$ENDIF}
      end;

      ZoneRIB := THDBEdit(GetControl('R_NUMEROCOMPTE'));
      if Assigned(ZoneRIB) then
      begin
        ZoneRib.Enabled := (TFFicheListe(Ecran).TypeAction <> taConsult);
        ZoneRib.MaxLength := 10 ;
        ZoneRib.EditMask := FormatZonesRib(CodeISO, 'CP');
        {$IFNDEF EAGLCLIENT}
          ZoneRib.Field.EditMask := FormatZonesRib(CodeISO, 'CP');
        {$ENDIF}
      end;

      ZoneRIB := THDBEdit(GetControl('R_CLERIB'));
      if Assigned(ZoneRIB) then
      begin
        ZoneRib.Enabled := (TFFicheListe(Ecran).TypeAction <> taConsult);
        ZoneRib.MaxLength := 2 ;
        ZoneRib.EditMask := FormatZonesRib(CodeISO, 'DC');
        {$IFNDEF EAGLCLIENT}
          ZoneRib.Field.EditMask := FormatZonesRib(CodeISO, 'DC');
        {$ENDIF}
      end;

      SetControlEnabled('R_CODEIBAN', False);
      if (ds.state in [dsInsert, dsEdit]) then
        if quelType(codeISO,GetControlText('R_TYPEPAYS')) <> 'BBA' then
           SetField('R_CODEIBAN', CodeISO); //XMG cela force le recalcul de l'IBAN
    end
    else
    begin
      SetControlEnabled('R_ETABBQ', False);
      SetControlEnabled('R_GUICHET', False);
      SetControlEnabled('R_NUMEROCOMPTE', False);
      SetControlEnabled('R_CLERIB', False);
      SetControlEnabled('R_CODEIBAN', True);
      if (ds.state in [dsInsert, dsEdit]) then
        if quelType(codeISO,GetControlText('R_TYPEPAYS')) <> 'BBA' then
          SetField('R_CODEIBAN', codeISO);
    end;

    ZoneCP := THDBEdit(GetControl('R_NUMEROCOMPTE'));
    ZoneRIB := THDBEdit(GetControl('R_CLERIB'));
    if (Assigned(ZoneCP)) and (Assigned(ZOneRIB)) then
    begin
      with THDBEdit(GetControl('R_GUICHET')) do
      begin
        ZoneLeft := Left + Width + 20;
        OrdreTab := TabOrder;
      end;

      if (CodeISO = CodeISOES) then
      begin //Si Espagne, on tourne la position des champs....
        ZoneRIB.Left := ZoneLeft;
        ZoneRIB.TabOrder := OrdreTab + 1;
        ZoneCP.Left := ZoneRIB.Left + ZoneRIB.Width + 20;
        ZoneCP.TabOrder := ZoneRIB.TabOrder + 1;
      end
      else
      begin //if CodeISO=CodeISOFR then //Pour tout autre Pays on re-établir l'affichage Français....
        ZoneCP.Left := ZoneLeft;
        ZoneCP.TabOrder := OrdreTab + 1;
        ZoneRIB.Left := ZoneCP.Left + ZoneCP.Width + 20;
        ZoneRIB.TabOrder := ZoneCP.TabOrder + 1;
      end;

      lblZoneRib := THLabel(GetControl('TR_NUMEROCOMPTE'));
      lblZoneRIB.Left := LblZoneRIB.FocusControl.Left + ((LblZoneRIB.FocusControl.Width - lblZoneRib.Width) div 2);
      lblZoneRib := THLabel(GetControl('TR_CLERIB'));
      lblZoneRIB.Left := LblZoneRIB.FocusControl.Left + ((LblZoneRIB.FocusControl.Width - lblZoneRib.Width) div 2);
    end;
end;

procedure Tom_RIB.AfficheEcran;
begin

  // Permet d'activer ou non le type pays
  AfficheTYPEPAYS(ecran,ds,'R_',codeISO) ;

  if (codeISO = CodeISOES) and (VH^.PaysLocalisation = CodeISOES) then
    // Affichage Specifique pour l'Espagne
    AfficheSpecifES
  else
    // Affichage Standard
    Affiche(ecran,codeISO,codeTypePays,'R_');
end;


procedure Tom_RIB.R_PAYSChange(Sender: TObject);
var stCodeISOPAYS, stPlus : string;
begin

  if FromSaisie then
    stCodeISOPAYS := codeISODuPays(GetControlText('R_PAYS'))
  else
    stCodeISOPAYS := codeISODuPays(GetField('R_PAYS'));

  if codeISO <> stCodeISOPAYS then
  begin
    codeISO := stCodeISOPAYS;

    stPlus := '' ;
    if Length(trim(codeISO))>0 then stPlus := ' AND YIB_PAYSISO="' + codeISO + '"' ;
    CodePays.Plus := stPlus ;

    // Affichage :
    AfficheEcran;
    if ds.State in [dsInsert, dsEdit] then
    begin
      SetField('R_ETABBQ', '');
      SetField('R_GUICHET', '');
      SetField('R_NUMEROCOMPTE', '');
      SetField('R_CLERIB', '');
      SetField('R_TYPEPAYS', '');

      if CodePays.Items.Count > 0 then
      begin        if Codepays.ItemIndex > -1 then

          SetField('R_TYPEPAYS',CodePays.Values[CodePays.itemindex])
        else
          SetField('R_TYPEPAYS',CodePays.Values[0]);
      end;
     // FQ 19469 BV 01/02/2007
     // if (((codeISO = codeISOES) and (VH^.PaysLocalisation = CodeISOES)) or
     //    ((codeISO = CodeISOFR) and not(existeParamPays(codeISO,codeTypePays)))) or
     //    (quelType(codeISO,GetControlText('R_TYPEPAYS')) = 'BBA') then
     //   SetField('R_CODEIBAN', '')
     // else
     //   SetField('R_CODEIBAN', codeISO);
      if (quelType(codeISO,GetControlText('R_TYPEPAYS')) = 'IBA') or
         ((quelType(codeISO,GetControlText('R_TYPEPAYS')) = '') and
         (((codeISO <> codeISOES) or (VH^.PaysLocalisation <> CodeISOES)) and ((codeISO <> codeISOFR) or (VH^.PaysLocalisation <> CodeISOFR)))) then
         SetField('R_CODEIBAN', codeISO)
      else
         SetField('R_CODEIBAN', '');
    end;
  end;
  R_TYPEPAYSChange(nil);
{$IFDEF EAGLCLIENT}
  if (ds.State in [dsInsert]) or
   (THVALComboBox(GetControl('R_TYPEPAYS')).text = '') then
  begin
  if THVALComboBox(GetControl('R_TYPEPAYS')).items.count > 0 then
  begin
   if (THVALComboBox(GetControl('R_TYPEPAYS')).text = '') then
      THVALComboBox(GetControl('R_TYPEPAYS')).ItemIndex := 0;
   SetField('R_TYPEPAYS',THVALComboBox(GetControl('R_TYPEPAYS')).Values[THVALComboBox(GetControl('R_TYPEPAYS')).itemindex]);
  end
  end;
{$ELSE}
  if (ds.State in [dsInsert]) or
     (THDBVALComboBox(GetControl('R_TYPEPAYS')).text = '') then
  begin
  if THDBVALComboBox(GetControl('R_TYPEPAYS')).items.count > 0 then
  begin
   if (THDBVALComboBox(GetControl('R_TYPEPAYS')).text = '') then
      THDBVALComboBox(GetControl('R_TYPEPAYS')).ItemIndex := 0;
   SetField('R_TYPEPAYS',THDBVALComboBox(GetControl('R_TYPEPAYS')).Values[THDBVALComboBox(GetControl('R_TYPEPAYS')).itemindex]);
  end
  end;
{$ENDIF}

end;

// SB 17/09/2003
function Tom_RIB.TypeRibUnique(TypeRib : string) : Boolean;
var
    Q : TQuery;
begin
    Result := True;
    if (GetControlText(TypeRib) <> 'X') then Exit;
    Q := OpenSql('Select ' + TypeRib + ' from RIB Where R_AUXILIAIRE="' + Auxiliaire + '" And ' +
        TypeRib + '="X" And R_NUMERORIB<>' + GetControlText('R_NUMERORIB') + '', True);
    if not Q.Eof then Result := False;
    Ferme(Q);
end;



procedure Tom_RIB.FlisteDblClick(Sender: TObject);
begin
//  TFFicheListe(Ecran).STa.
  TFFicheListe(Ecran).Retour := THDBEdit(getControl('R_NUMERORIB')).Text;
  TFFicheListe(Ecran).Close;
end;

initialization
  registerclasses([Tom_RIB]);
end.





