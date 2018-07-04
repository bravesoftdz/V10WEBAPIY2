unit UTofProposToAffaire;

interface
uses  StdCtrls,Controls,Classes,Forms,sysutils,ComCtrls,Vierge, Dicobtp,
{$IFDEF EAGLCLIENT}
     Maineagl,
{$ELSE}
	 {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db, FE_Main,
{$ENDIF}
{$IFDEF BTP}
	  CalcOleGenericBTP,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOF, EntGC,AffaireUtil,FactUtil,UtofAfBaseCodeAffaire;


Type
     TOF_PROPOSTOAFFAIRE = Class (TOF_AFBASECODEAFFAIRE)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnUpdate ; override ;
        procedure OnClose ; override ;
     private
        bSortieOK,bAffModele,bRecodif,bProposition: boolean;
        Part0, Part1, Part2, Part3,Avenant, Affaire, Tiers: THEDIT;
        AffairePropos : string;
     END ;

Type
     TOF_PARAMPROTOAFF = Class (TOF_AFBASECODEAFFAIRE)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnUpdate ; override ;
        procedure OnClose ; override ;
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
        procedure TypeAffaire_Click(Sender : TObject);
     private
        bMajAff,bSortieOK  :boolean;
        {CodeaffMaj, }TiersRef : string;
     END ;
Function AFLanceFiche_ParamProtoAff(Argument:string):variant;
Function AFLanceFiche_ProposToAffaire(Argument:string):variant;

implementation

procedure TOF_PROPOSTOAFFAIRE.OnArgument(stArgument : String );
var
EditProp, EditAff:THLabel;
critere, champ, valeur, AffNew0, AffNew1, AffNew2, AffNew3, AffNew4,CodeOrigine :string;
x:integer;
Begin
Inherited;
bSortieOK:=true; bAffModele := false;
EditProp:=THLabel(GetControl('LIBPROP'));
EditAff:=THLabel(GetControl('LIBAFFAIRE'));
Part0:=THEDIT(GetControl('ACT_AFFAIRE0'));
Part1:=THEDIT(GetControl('ACT_AFFAIRE1'));
Part2:=THEDIT(GetControl('ACT_AFFAIRE2'));
Part3:=THEDIT(GetControl('ACT_AFFAIRE3'));
Avenant:=THEDIT(GetControl('ACT_AVENANT'));
Affaire:=THEDIT(GetControl('ACT_AFFAIRE'));
Tiers:=THEDIT(GetControl('ACT_TIERS'));

bRecodif := False;
bProposition := False;
// on lit dans l''argument, le code proposition et le code affaire passés
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
  BEGIN
  X:=pos(':',Critere);
  if x<>0 then
    begin
    Champ:=copy(Critere,1,X-1);
    Valeur:=Copy (Critere,X+1,length(Critere)-X);
    end
  else Champ := Critere;

  if (Champ='AFFAIRE') then AffairePropos := Valeur
  else if (Champ='RECODIFICATION') then bRecodif:= True
  else if (Champ='PROPOSITION') then CodeOrigine := Valeur
  else if (Champ='TIERS') then Tiers.Text := Valeur
  else if (Champ='LIBTIERS') then SetControlText('LIBTIERS', Valeur)
  else if (Champ='TYPE') and (valeur = 'AFFAIREMODELE') then  bAffModele := true;
  Critere:=(Trim(ReadTokenSt(stArgument)));
  END;

  if bAffModele then
  begin
    if CodeOrigine[1] ='P' then
    begin
      bProposition := true;
      Ecran.Caption := TraduitGA('Saisie d''un code proposition en duplication de modèle');
      UpdateCaption(Ecran);
      {$IFDEF BTP}
      EditProp.Caption := TraduitGA('Code proposition modèle  : ') + BTPCodeAffaireAffiche(CodeOrigine  ,' ');
      EditAff.caption  := TraduitGA('Code proposition proposé : ') + BTPCodeAffaireAffiche(AffairePropos,' ');
      {$ELSE}
      EditProp.Caption := TraduitGA('Code proposition modèle  : ') + CodeAffaireAffiche(CodeOrigine  ,' ');
      EditAff.caption  := TraduitGA('Code proposition proposé : ') + CodeAffaireAffiche(AffairePropos,' ');
      {$ENDIF}
      SetControlText('TITRE1', TraduitGA('Vous pouvez saisir un code proposition différent du code proposé :'));
      SetControlText('TITRE2', TraduitGA('Attention : si le code saisi est vide, la proposition ne sera pas générée.'));
    end else
    begin
      Ecran.Caption := TraduitGA('Saisie d''un code affaire en duplication de modèle');
      UpdateCaption(Ecran);
      {$IFDEF BTP}
      EditProp.Caption := TraduitGA('Code affaire modèle  : ') + BTPCodeAffaireAffiche(CodeOrigine  ,' ');
      EditAff.caption  := TraduitGA('Code affaire proposé : ') + BTPCodeAffaireAffiche(AffairePropos,' ');
      {$ELSE}
      EditProp.Caption := TraduitGA('Code affaire modèle  : ') + CodeAffaireAffiche(CodeOrigine  ,' ');
      EditAff.caption  := TraduitGA('Code affaire proposé : ') + CodeAffaireAffiche(AffairePropos,' ');
      {$ENDIF}
      SetControlText('TITRE1', TraduitGA('Vous pouvez saisir un code affaire différent du code proposé :'));
      SetControlText('TITRE2', TraduitGA('Attention : si le code saisi est vide, l''affaire ne sera pas générée.'));
    end;
  end else
  begin
	{$IFDEF BTP}
    EditProp.Caption := EditProp.Caption+ ' ' + BTPCodeAffaireAffiche(CodeOrigine,' ');
    if bRecodif then
    EditAff.Caption := TraduitGA('Code affaire proposé : ') + BTPCodeAffaireAffiche(AffairePropos,' ')
    else
    EditAff.Caption := EditAff.Caption + ' ' + BTPCodeAffaireAffiche(AffairePropos,' ');
    {$ELSE}
    EditProp.Caption := EditProp.Caption+ ' ' + CodeAffaireAffiche(CodeOrigine,' ');
    if bRecodif then
    EditAff.Caption := TraduitGA('Code affaire proposé : ') + CodeAffaireAffiche(AffairePropos,' ')
    else
    EditAff.Caption := EditAff.Caption + ' ' + CodeAffaireAffiche(AffairePropos,' ');
    {$ENDIF}
  end;

  if bRecodif then
  begin
    //EditAff.Caption:='Code affaire proposé '+ CodeAffaireAffiche(AffairePropos,' ');
    // repris par défaut dans le nouveau code
    {$IFDEF BTP}
    BTPCodeAffaireDecoupe(AffairePropos,AffNew0, AffNew1, AffNew2, AffNew3, AffNew4, taCreat, false);
    {$ELSE}
    CodeAffaireDecoupe(AffairePropos,AffNew0, AffNew1, AffNew2, AffNew3, AffNew4, taCreat, false);
    {$ENDIF}
    Part0.Text:=AffNew0;  Part1.Text:=AffNew1;
    if VH_GC.CleAffaire.Co2Visible then Part2.Text := AffNew2;
    if VH_GC.CleAffaire.Co3Visible then Part3.Text := AffNew3;
  end;
End;


procedure TOF_PROPOSTOAFFAIRE.OnUpdate;
var
CodeAffaireRet:string;
iPartErreur:integer;
Begin
inherited;
bSortieOK:=true;

CodeAffaireRet := DechargeCleAffaire(Part0, Part1, Part2, Part3, Avenant, Tiers.Text,
                                                taCreat, True, True, bProposition, iPartErreur);
if (iPartErreur=0) then
    TFVierge(ecran).Retour := CodeAffaireRet
else
    begin
    if (Part1.CanFocus) then Part1.SetFocus else
    if (Part2.CanFocus) then Part2.SetFocus else
    if (Part3.CanFocus) then Part3.SetFocus;
    bSortieOK:=false;
    end;
End;

procedure TOF_PROPOSTOAFFAIRE.OnClose;
begin
if Not bSortieOK then
    LastError:=1;
end;

/// ****************** Paramétrage TOF PARAMPROTOAFF ***************************
procedure TOF_PARAMPROTOAFF.OnArgument(stArgument : String );
var
EditTiers,EditPro:THEdit;
Rb : TRadioButton;
critere, champ, valeur, CodeProposition:string;
x:integer;
bCpt : Boolean;
Begin
Inherited;
bSortieOK:=true;  bMajAff := False;
EditPro:=THEdit(GetControl('PROPOSITION'));
EditTiers:=THEdit(GetControl('AFF_TIERS'));
Rb := TRadioButton(Getcontrol('RBCREATEAFFAIRE'));
if Rb <> Nil then Rb.Onclick := TypeAffaire_Click;
Rb := TRadioButton(Getcontrol('RBMAJAFFAIRE'));
if Rb <> Nil then Rb.Onclick := TypeAffaire_Click;
// CBGARDERCOMPTEUR est-il visible en fonction du param du code affaire
// Affiché uniquement si compteur dans le découpage + même compteur proposition / affaire
bCpt := False;
if VH_GC.CleAffaire.ProDifferent = False then
   begin
   if (Pos(VH_GC.CleAffaire.Co1Type,'CPT;CPM')>0) then bCpt := true;
   if (Pos(VH_GC.CleAffaire.Co2Type,'CPT;CPM')>0) And (VH_GC.CleAffaire.Co2Visible) then bCpt := true;
   if (Pos(VH_GC.CleAffaire.Co3Type,'CPT;CPM')>0) And (VH_GC.CleAffaire.Co3Visible) then bCpt := true;
   end;
SetControlVisible('CBGARDERCOMPTEUR',bCpt);
SetControlChecked('CBGARDERCOMPTEUR',bCpt);

// on lit dans l''argument, le code proposition
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
    X:=pos(':',Critere);
    if x<>0 then begin Champ:=copy(Critere,1,X-1); Valeur:=Copy (Critere,X+1,length(Critere)-X); end;
    if (Champ='PROPOSITION') then
      begin
      CodeProposition := Valeur;
      {$IFDEF BTP}
      EditPro.Text:=BTPCodeAffaireAffiche (Valeur,' ');
      {$ELSE}
      EditPro.Text:=CodeAffaireAffiche (Valeur,' ');
      {$ENDIF}
      end
    else
    if (Champ='TIERS') then begin EditTiers.Text := Valeur; TiersRef := Valeur; end;

    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;
// Affichage du transfert des lignes d'activite si elles existent
if ExisteSQL('SELECT ACT_AFFAIRE FROM ACTIVITE WHERE ACT_TYPEACTIVITE="REA" AND ACT_AFFAIRE="'+CodeProposition+'"') then
   SetControlVisible ('ACTIVITEPROTOAFF',True)
else
   SetControlvisible ('ACTIVITEPROTOAFF',False);
End;


procedure TOF_PARAMPROTOAFF.OnUpdate;
var
CodeAffaireRet, stRet :string;
Part0, Part1, Part2, Part3,Avenant, Affaire, Tiers: THEDIT;
iPartErreur:integer;
Begin
inherited;
stRet:='';
if (TRadioButton(GetControl('RBCREATEAFFAIRE')).checked=true) then
    stRet:= 'PROP;';
bSortieOK:=true;  iPartErreur := 0;
if bMajAff then
   begin
   Part0:=THEDIT(GetControl('AFF_AFFAIRE0')); Part1:=THEDIT(GetControl('AFF_AFFAIRE1'));
   Part2:=THEDIT(GetControl('AFF_AFFAIRE2')); Part3:=THEDIT(GetControl('AFF_AFFAIRE3'));
   Avenant:=THEDIT(GetControl('AFF_AVENANT'));
   Affaire:=THEDIT(GetControl('AFF_AFFAIRE')); Tiers:=THEDIT(GetControl('AFF_TIERS'));
    // dans le cas où frappe de code, le champ affaire n'est pas renseigné,
    // il faut voir si les parties du code tapé existe en fichier
   if Affaire.text='' then  begin
      Affaire.text:=DechargeCleAffaire(Part0, Part1, Part2, Part3, Avenant, Tiers.Text,taModif, False, True, False, iPartErreur);
      ipartErreur:=TestCleAffaire(Affaire,part1,part2,part3,avenant,tiers,part0.Text, True, True,False);
      if ipartErreur<=0 then begin
         IpartErreur:=-1;
         affaire.text:='';
         end;
      end;
   if Tiers.Text <> Tiersref then
      begin
      PGIBoxAf ('Le tiers associé à l''affaire est différent de la proposition','Acceptation d''affaires');
      iPartErreur := -1;
      end;
   if (iPartErreur=0) then
      CodeAffaireRet:=Affaire.text; // mcd 08/02/02 les parties ne sont pas renseignées si pas visible
      {CodeAffaireRet := DechargeCleAffaire(Part0, Part1, Part2, Part3, Avenant, Tiers.Text,
                                                taModif, False, True, False, iPartErreur); }
   if (iPartErreur=0) then
      begin
      stRet := stRet + 'AFFAIRE:' + CodeAffaireRet;
      end
   else
       begin
       if (Part1.CanFocus) then Part1.SetFocus else
       if (Part2.CanFocus) then Part2.SetFocus else
       if (Part3.CanFocus) then Part3.SetFocus;
       bSortieOK:=false;
       end;
   end;
if (GetControlText('CBGARDERCOMPTEUR')='X') then stRet:=stRet+'GARDERCOMPTEUR:X;';
if (GetControlText('ACTIVITEPROTOAFF')='X') then stRet:=stRet+'ACTIVITEPROTOAFF:X;';

TFVierge(ecran).Retour := stRet ;
End;

procedure TOF_PARAMPROTOAFF.OnClose;
begin
if Not bSortieOK then
    begin
    LastError:=1;
    BSortieOk :=True; // poour repartir dans bonne condition mcd 17/07/03
    end;
end;

procedure TOF_PARAMPROTOAFF.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('AFF_AFFAIRE'));
Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
Tiers:=THEdit(GetControl('AFF_TIERS'));
end;

procedure TOF_PARAMPROTOAFF.TypeAffaire_Click(Sender : TObject);
begin
If TRadioButton(Sender).Tag = 1 then bMajAff:=Not(TRadioButton(Sender).Checked)
                                else bMajAff:=TRadioButton(Sender).Checked;
SetControlVisible('GBMAJAFFAIRE',bmajAff);
end;

Function AFLanceFiche_ParamProtoAff(Argument:string):variant;
begin
result:=AGLLanceFiche ('AFF','AFPARAMPROTOAFF','','',Argument);
end;
Function AFLanceFiche_ProposToAffaire(Argument:string):variant;
begin
result:=AGLLanceFiche ('AFF','AFPROPOSTOAFFAIRE','','',Argument);
end;


Initialization
registerclasses([TOF_PROPOSTOAFFAIRE,TOF_PARAMPROTOAFF]);


end.
