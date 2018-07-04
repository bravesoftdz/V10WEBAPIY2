unit UTofAFValidePropos;

interface                   
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
{$IFDEF EAGLCLIENT}
   Maineagl, eMul,
{$ELSE}
   dbTables, db, FE_Main, Mul,HDB,DBCtrls,
{$ENDIF}
      HCtrls,Ent1,HMsgBox,UTOF, AffaireUtil,DicoAF,SaisUtil,EntGC,utofAfBaseCodeAffaire,HTB97,
      Buttons, ExtCtrls, HEnt1, UTob, M3FP, factGrp, factutil,
      ActiviteUtil, AffaireDuplic, UtilMulTrt
      ,UtofChoixtrans_Act,
      UtofProposToAffaire;

Type
     TOF_AFVALIDEPROPOS = Class (TOF_AFBASECODEAFFAIRE)
         procedure OnArgument(stArgument : String ) ; override ;
         procedure OnUpdate ; override ;
         procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
         procedure AFF_ETATAFFAIREChange(Sender: TObject);
         Function  ValideLaProposition(CodeProposition, CodeTiers : string) : string;
         procedure RefuseLaProposition(CodeAffaire:string; TOBL,TobAffaire:TOB);
// PL le 16/05/02 : Pb, processus n'est pas annulé si on sort avec la croix
         function RecupParamValidation(CodeProposition, CodeTiers : string):string;
     public
      EtatAffaire :  THValComboBox;
      bOuvrir     :  TToolBarButton97;
      bRefuser    :  TToolBarButton97;
      bReactiver  :  TToolBarButton97;
      stAffMaj    :  String;
      bGarderCpt, bMajAff, bStop, bGardeLaProp, ActProToAff  : Boolean;
      procedure TraitementValidePropositions;
      procedure RefusePropositions;
      procedure ReactivePropositions;
     END ;
const
	// libellés des messages
	TexteMessage: array[1..6] of string 	= (
          {1}        'Confirmez-vous la validation des propositions sélectionnées ?'
          {2}       ,'Confirmez-vous le refus des propositions sélectionnées ?'
          {3}       ,'Confirmez-vous la réactivation des propositions sélectionnées ?'
          {4}       ,'Voulez-vous transférer les lignes d''activité de la proposition '
          {5}       ,' sur l''affaire '
          {6}       ,'Le transfert des lignes d''activité a échoué, voulez-vous annuler le refus ?'
                       );

Procedure AFLanceFiche_ValidePropos(Range,Argument:string);

implementation


procedure TOF_AFVALIDEPROPOS.OnArgument(stArgument : String );
var  Critere : string;
Begin
Inherited;
Critere:=(Trim(ReadTokenSt(stArgument)));
EtatAffaire :=  THValComboBox(GetControl('AFF_ETATAFFAIRE'));
EtatAffaire.Plus :=GetPlusEtatAffaire (False); //mcd 19/06/03
bOuvrir     :=  TToolBarButton97(GetControl('BOUVRIR'));
bRefuser    :=  TToolBarButton97(GetControl('BREFUSER'));
bReactiver  :=  TToolBarButton97(GetControl('BREACTIVER'));
EtatAffaire.OnChange:=AFF_ETATAFFAIREChange;
end;

procedure TOF_AFVALIDEPROPOS.AFF_ETATAFFAIREChange(Sender: TObject);
begin
inherited;
if (EtatAffaire.Value='ENC') then
   begin
   bOuvrir.Enabled := true;
   bRefuser.Enabled := true;
   bReactiver.Enabled := false;
   end
else
if (EtatAffaire.Value='REF') then
   begin
   bReactiver.Enabled := true;
   bOuvrir.Enabled := false;
   bRefuser.Enabled := false;
   end
else
if (EtatAffaire.Value='ACC') then
   begin
   bReactiver.Enabled := false;
   bOuvrir.Enabled := false;
   bRefuser.Enabled := false;
   end;
end;

procedure TOF_AFVALIDEPROPOS.OnUpdate;
Begin
inherited;
if Not (VH_GC.CleAffaire.Co2Visible) then SetControlText('AFF_AFFAIRE2','');
if Not (VH_GC.CleAffaire.Co3Visible) then SetControlText('AFF_AFFAIRE3','');
End;

procedure TOF_AFVALIDEPROPOS.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('AFF_AFFAIRE'));
Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
Tiers:=THEdit(GetControl('AFF_TIERS'));
end;

// PL le 16/05/02 : Pb, processus n'est pas annulé si on sort avec la croix
// + on gère la sortie par le bouton STOP
Function TOF_AFVALIDEPROPOS.ValideLaProposition(CodeProposition, CodeTiers : string) : string;
var
   Rep,arg : string;
begin
Arg := ''; Result :=''; bStop:=false; bGardeLaProp:=false;
// appel Ecran de paramétrage de l'affaire générée
Rep:=RecupParamValidation(CodeProposition, CodeTiers);
if ((bGardeLaProp=false) and (Rep='')) or bStop then exit;

// passage d'arguments
if bGarderCpt  then arg := arg + 'GARDERCPTPRO:X;' ;
if ActProToAff then arg := arg + 'ACTPROTOAFF:X;' ;
if (bMajAff) and (stAffMaj <> '') then
   begin
   arg := arg + 'AFFAIREMAJBYPRO:'+ stAffMaj+ ';' ;
   end;
Result := DuplicationAffaire (tdaProToAff, CodeProposition, Arg, Nil,True, False, False);
// Bascule d'un prospect en client
if (trim(Result) <> '') and  (GetInfoParPiece(VH_GC.AFNatAffaire,'GPP_PROCLI') = 'X') then
  AFMajProspectClient(CodeTiers);
end;
// Fin PL le 16/05/02

// PL le 16/05/02 : Pb, processus n'est pas annulé si on sort avec la croix
// + on gère la sortie par le bouton STOP
function TOF_AFVALIDEPROPOS.RecupParamValidation(CodeProposition, CodeTiers : string):string;
Var
Retour,Critere,Champ,Valeur : string;
X : integer;
begin
Retour:='';

//Initialisation par défaut
bGarderCpt := False; // nouvelle num de l'affaire par défaut
ActProToAff:= False; // Transfert lignes d'activité en proposition
bMajAff := false; stAffMaj := ''; // création d'une nouvelle affaire par défaut

Retour:=AFLanceFiche_ParamProtoAff('PROPOSITION:'+ CodeProposition+'; TIERS:'+ CodeTiers+';') ;
Result:=Retour;

Critere:=(Trim(ReadTokenSt(Retour)));
While (Critere <>'') do
    BEGIN
    X:=pos(':',Critere);
    if x<>0 then begin Champ:=copy(Critere,1,X-1); Valeur:=Copy (Critere,X+1,length(Critere)-X); end
        else Champ:=Critere;
    if (Champ='AFFAIRE') And (Valeur <> '') then begin bMajAff := true; stAffMaj := Valeur; end;
    if (Champ='GARDERCOMPTEUR') And (Valeur ='X') then bGarderCpt := True;
    if (Champ='ACTIVITEPROTOAFF') And (Valeur ='X') then ActProToAff := True;
    if (Champ='STOP') then bStop := True;
    if (Champ='PROP') then bGardeLaProp := True;

    Critere:=(Trim(ReadTokenSt(Retour)));
    END;
end;
// Fin PL le 16/05/02

procedure TOF_AFVALIDEPROPOS.TraitementValidePropositions;
var   CodeProposition, CodeTiers, {Req, }LastAff : string;
      i,NbAff  : Integer;
      TobMAff : TOB;
begin
  NbAff := 0; LastAff := '';
  if (PGIAskAF (TexteMessage[1], Ecran.Caption)<>mrYes) then exit;
  TobMAff := Tob.Create('les affaires',NIL,-1);
  try
  SourisSablier;

  TraiteEnregMulListe (TFMul(Ecran), 'AFF_AFFAIRE','AFFAIRE', TobMAff, True);
  For i:=0 to TobMAff.Detail.Count-1 do
     begin
     if TobMAff.detail[i].GetValue('AFF_AFFAIRE') <> '' then
        begin
        Inc(NbAff);
        CodeProposition:= TobMAff.detail[i].GetValue('AFF_AFFAIRE');
        CodeTiers:= TobMAff.detail[i].GetValue('AFF_TIERS');
        LastAff := ValideLaProposition(CodeProposition, CodeTiers);
        end;
     end;

  Finally
  SourisNormale ;
  TobMaff.Free;
  // Si une seule affaire est crée => ouverte en auto. depuis le script
  if (NbAff = 1 ) and (LastAff <> '') then
     begin SetControlText('LASTAFFAIRE',LastAff); SetControlText('LASTTIERS',CodeTiers); end
  else
     begin SetControlText('LASTAFFAIRE',''); SetControlText('LASTTIERS',''); end;
  end;
end;

procedure TOF_AFVALIDEPROPOS.RefuseLaProposition(CodeAffaire:string; TOBL,TobAffaire:TOB);
var
CodeAffaireAdm, sReq:string;
begin
CodeAffaireAdm:='';

   TOBL.InitValeurs ;
   TOBL.SelectDB('"'+CodeAffaire +'"',Nil) ;
   // Modif sur la proposition à refuser
   TOBL.PutValue('AFF_ETATAFFAIRE', 'REF');
   TOBL.PutValue('AFF_DATESIGNE', 0);

   // Transfert des lignes d'activité sur affaire administrative ?
   if ExisteSQL('SELECT ACT_AFFAIRE FROM ACTIVITE WHERE ACT_AFFAIRE="'+TOBL.GetValue('AFF_AFFAIRE')+'"') then
      begin
      CodeAffaireAdm:=AFLanceFiche_ChoixTransF_ACt('PROPOSITION:'+TOBL.GetValue('AFF_AFFAIRE')) ;

        //  pas de saisie            action annulée           sur proposition
      if (CodeAffaireAdm<>'') and (CodeAffaireAdm<>'0') and (CodeAffaireAdm<>'1') then
         begin
         if AFRemplirTOBAffaire(CodeAffaireAdm, TobAffaire, nil) then
         // on a selectionne un code affaire administrative
            begin
            try
            sReq := 'UPDATE ACTIVITE SET ACT_AFFAIRE="'+TobAffaire.GetValue('AFF_AFFAIRE')
                                       +'", ACT_AFFAIRE0="'+TobAffaire.GetValue('AFF_AFFAIRE0')
                                       +'", ACT_AFFAIRE1="'+TobAffaire.GetValue('AFF_AFFAIRE1')
                                       +'", ACT_AFFAIRE2="'+TobAffaire.GetValue('AFF_AFFAIRE2')
                                       +'", ACT_AFFAIRE3="'+TobAffaire.GetValue('AFF_AFFAIRE3')
                                       +'", ACT_AVENANT="'+TobAffaire.GetValue('AFF_AVENANT')
                                       +'", ACT_TIERS="'+TobAffaire.GetValue('AFF_TIERS')
                                       +'" WHERE ACT_TYPEACTIVITE="REA" AND ACT_AFFAIRE="'+TOBL.GetValue('AFF_AFFAIRE')+'"';
            ExecuteSQL(sReq);
            except
            if (PGIAskAF (TexteMessage[6], Ecran.Caption)=mrYes) then CodeAffaireAdm:='0';
            end;
            end;
         end;
      end;

   // si l'action n'a pas ete annulée
   if (CodeAffaireAdm<>'0') then TOBL.UpdateDB(false);
end;

procedure TOF_AFVALIDEPROPOS.RefusePropositions;
var   i:integer;
      TOBL, TobAffaire, TobMAff:TOB;
begin
if (PGIAskAF (TexteMessage[2], Ecran.Caption)<>mrYes) then exit;
TOBL:=TOB.Create('AFFAIRE',Nil,-1) ;
TobAffaire:=TOB.Create('AFFAIRE',Nil,-1) ;
TobMAff := Tob.Create('les affaires',NIL,-1);
try
SourisSablier;
TraiteEnregMulListe (TFMul(Ecran), 'AFF_AFFAIRE','AFFAIRE', TobMAff, True);
For i:=0 to TobMAff.Detail.Count-1 do
   begin
   if TobMAff.detail[i].GetValue('AFF_AFFAIRE') <> '' then
      RefuseLaProposition(TobMAff.detail[i].GetValue('AFF_AFFAIRE'), TOBL,TobAffaire);
   end;

finally
TOBL.Free; TOBAffaire.Free;  TobMAff.Free;
SourisNormale ;
end;
end;

procedure TOF_AFVALIDEPROPOS.ReactivePropositions;
var   i:integer;
      TOBL, TobMAff:TOB;
begin
  if (PGIAskAF (TexteMessage[3], Ecran.Caption)<>mrYes) then exit;
  TOBL:=TOB.Create('AFFAIRE',Nil,-1) ;
  TobMAff := Tob.Create('les affaires',NIL,-1);
  try
    SourisSablier;
    TraiteEnregMulListe (TFMul(Ecran), 'AFF_AFFAIRE','AFFAIRE', TobMAff, True);
    For i:=0 to TobMAff.Detail.Count-1 do
       begin
       if TobMAff.detail[i].GetValue('AFF_AFFAIRE') <> '' then
          begin
          TOBL.InitValeurs ;
          TOBL.SelectDB('"'+ TobMAff.detail[i].GetValue('AFF_AFFAIRE') +'"',Nil) ;
          // Modif sur la proposition en cours
          TOBL.PutValue('AFF_ETATAFFAIRE', 'ENC'); TOBL.PutValue('AFF_DATESIGNE', 0);
          TOBL.UpdateDB(false);
          end;
       end;
  finally
    TOBL.Free; TobMAff.Free;
    SourisNormale ;
  end;
end;

procedure AGLTraitementValidePropositions( parms: array of variant; nb: integer ) ;
var  F : TForm ;
      TOFV:TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFMUL) then TOFV:=TFMUL(F).LaTOF else exit;
Transactions(TOF_AFVALIDEPROPOS(TOFV).TraitementValidePropositions,2) ;
end;

procedure AGLRefusePropositions( parms: array of variant; nb: integer ) ;
var  F : TForm ;
      TOFV:TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFMUL) then TOFV:=TFMUL(F).LaTOF else exit;
TOF_AFVALIDEPROPOS(TOFV).RefusePropositions;
end;

procedure AGLReactivePropositions( parms: array of variant; nb: integer ) ;
var  F : TForm ;
      TOFV:TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFMUL) then TOFV:=TFMUL(F).LaTOF else exit;
TOF_AFVALIDEPROPOS(TOFV).ReactivePropositions;
end;
Procedure AFLanceFiche_ValidePropos(Range,Argument:string);
begin
AGLLanceFiche ('AFF','AFVALIDEPROPOS',Range,'',Argument);
end;


Initialization
registerclasses([TOF_AFVALIDEPROPOS]);
RegisterAglProc( 'TraitValPropositions', true , 0, AGLTraitementValidePropositions);
RegisterAglProc( 'RefusePropositions', true , 0, AGLRefusePropositions);
RegisterAglProc( 'ReactivePropositions', true , 0, AGLReactivePropositions);
end.
