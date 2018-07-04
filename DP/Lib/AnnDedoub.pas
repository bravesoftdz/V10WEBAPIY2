unit AnnDedoub;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  assist, ComCtrls, HSysMenu, hmsgbox, StdCtrls, ExtCtrls, HTB97, Hctrls,
  Vierge, UtilTOM, AnnOutils, HPanel,Utob;

type
  TFAnnDedoub = class(TFAssist)
    PDetail1: TTabSheet;
    PDetail2: TTabSheet;
    lFiche1: THLabel;
    MemoDetail1: TMemo;
    MemoDetail2: TMemo;
    lFiche2: THLabel;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Déclarations privées }
    function RecupCocheFrm(nomzone: String): Boolean;
    function RecupTextFrm(nomzone: String): String;
    {function RecupCaptionFrm(nomzone: String): String;}
    procedure AjouteDetail(num, txt: String; saufblancs: Boolean=False);
  public
    { Déclarations publiques }
    Frm: TForm; // fiche appelante (pour lecture du contenu)
    NoSuppr: String; // no du bouton de la fiche appelante (AFSUPRDOUBLONANNU)
    ObjSuppr: TPerSup; // profite des tests de suppression de AnnOutils
  end;

procedure FicheAssistantDedoublonnage(F: TForm; numforce: String);


///////////////// IMPLEMENTATION /////////////////
implementation

{$R *.DFM}

procedure FicheAssistantDedoublonnage(F: TForm; numforce: String);
var FAnnDedoub: TFAnnDedoub;
begin
  If F=Nil then exit;
  // ouvre assistant
  FAnnDedoub:=TFAnnDedoub.Create(Application) ;
  Try
    FAnnDedoub.Frm := F ;
    FAnnDedoub.NoSuppr := numforce;
    FAnnDedoub.ShowModal ;
  Finally
    FAnnDedoub.Free ;
  end ;
end;


procedure TFAnnDedoub.FormShow(Sender: TObject);
var i, cpteur, nb: Integer;
    no, Msg, tmp: String;
    QQ: Tquery;
begin
  inherited;
  // récup des données affichées dans AFSUPRDOUBLONANU
  for i:=1 to 2 do
    BEGIN
    no := IntToStr(i);
    ObjSuppr.GuidPer := RecupTextFrm('AN'+no+'_GUIDPER');
    ObjSuppr.Bavard := False;

    // Récup des liens et commentaires détaillés
    cpteur := 1;
    Msg := '1) Identification de la fiche annuaire n° '+ObjSuppr.GuidPer+' :'+#13+#10;
    tmp := RecupTextFrm('AN'+no+'_NOM1');
    if tmp<>'' then Msg := Msg + tmp + #13+#10;
    tmp := RecupTextFrm('AN'+no+'_NOM2');
    if tmp<>'' then Msg := Msg + tmp + #13+#10;
    tmp := RecupTextFrm('AN'+no+'_ALRUE1');
    if tmp<>'' then Msg := Msg + tmp + #13+#10;
    tmp := RecupTextFrm('AN'+no+'_ALCP')+' '+RecupTextFrm('AN'+no+'_ALVILLE');
    if tmp<>'' then Msg := Msg + tmp + #13+#10;
    Msg := Msg + #13+#10;

    // GI / Compta
    Msg := Msg + '2) Liens Tiers Compta/GI'+#13+#10;
    // permissif : toujours False
    // if Not ObjSuppr.TestCompteTiers(Msg) then Msg := Msg+'=> aucuns.'+#13+#10;
    tmp := ''; ObjSuppr.TestCompteTiers(tmp);
    if Trim(tmp)='' then Msg := Msg+'=> aucuns.'+#13+#10
                    else Msg := Msg+tmp+#13+#10;
    Msg := Msg + #13+#10;

    // JURI
    Msg := Msg + '3) Dossiers juridiques :'+#13+#10;
    if Not ObjSuppr.TestJuriDosPer(Msg) then Msg := Msg+'=> aucuns.'+#13+#10;
    Msg := Msg + #13+#10;

    Msg := Msg + '4) Liens vers des dossiers juridiques :'+#13+#10;
    if Not ObjSuppr.TestJuriLienPer(Msg) then Msg := Msg+'=> aucuns.'+#13+#10;
    Msg := Msg + #13+#10;

    Msg := Msg + '5) Informations de dossiers juridiques :'+#13+#10;
    if Not ObjSuppr.TestJuriDosInfoPer(Msg) then Msg := Msg+'=> aucunes.'+#13+#10;
    Msg := Msg + #13+#10;

    Msg := Msg + '6) Groupes de sociétés juridiques :'+#13+#10;
    if Not ObjSuppr.TestJuriGroupeSocPer(Msg) then Msg := Msg+'=> aucunes.'+#13+#10;
    Msg := Msg + #13+#10;

    // permissif : en mode non bavard, force la suppression
    Msg := Msg + '7) Evènements juridiques :'+#13+#10;
    tmp := ''; ObjSuppr.TestJuriEvenementPer(tmp);
    if Trim(tmp)='' then Msg := Msg+'=> aucuns.'+#13+#10
                    else Msg := Msg+tmp+#13+#10;
    Msg := Msg + #13+#10;

    // LIENS DIVERS
    Msg := Msg + '8) Contacts :'+#13+#10;
    tmp := '';
    QQ := OpenSQL('select count(c_typecontact) from CONTACT where c_typecontact="ANN" and C_GUIDPER="'+ObjSuppr.GuidPer+'"', true);
    if not(QQ.EOF)  then
      begin
      nb := QQ.Fields[0].AsInteger;
      if nb=1 then tmp := IntToStr(nb)+' contact est rattaché à cette personne.'+#13+#10
      else tmp := IntToStr(nb)+' contacts sont rattachés à cette personne.'+#13+#10;
      end;
    // permissif : en mode non bavard, force la suppression
    if Trim(tmp)='' then Msg := Msg+'=> aucuns.'+#13+#10
                    else Msg := Msg+tmp+#13+#10;
    Msg := Msg + #13+#10;

    Msg := Msg + '9) Liens annuaire du DP :'+#13+#10;
    if Not ObjSuppr.TestLienPer(Msg) then Msg := Msg+'=> aucuns.'+#13+#10;
    Msg := Msg + #13+#10;

    Msg := Msg + '10) Liens avec les plaquettes :'+#13+#10;
    if Not ObjSuppr.TestLienPlaq(Msg) then Msg := Msg+'=> aucuns.'+#13+#10;
    Msg := Msg + #13+#10;

    // Existence dossier de production
    Msg := Msg + '11) Dossier de production :'+#13+#10;
    if RecupCocheFrm('CBDOSSIER'+no) then
      begin
      Msg := Msg + '=> un dossier est rattaché à cette fiche.'+#13+#10;
      if Not ObjSuppr.TestDossierParti(Msg) then Msg := Msg+'=> pas de dossier en déplacement.'+#13+#10;
      end
    else
      Msg := Msg+'=> aucun.'+#13+#10;
    Msg := Msg + #13+#10;

    Msg := Msg + '12) Fiscalité personnelle :'+#13+#10;
    if Not ObjSuppr.TestFisPer(Msg) then Msg := Msg+'=> néant'+#13+#10;
    Msg := Msg + #13+#10;

{
    // Récup des liens et commentaires détaillés
    cpteur := 1;
    AjouteDetail(no, '1) Identification de la fiche annuaire n° '+ObjSuppr.GuidPer+' :');
    AjouteDetail(no, RecupTextFrm('AN'+no+'_NOM1'), True);
    AjouteDetail(no, RecupTextFrm('AN'+no+'_NOM2'), True);
    AjouteDetail(no, RecupTextFrm('AN'+no+'_ALRUE1'), True);
    AjouteDetail(no, RecupTextFrm('AN'+no+'_ALCP')+' '+RecupTextFrm('AN'+no+'_ALVILLE'), True);
    AjouteDetail(no, ' ');

    // Liens Tiers
    if RecupCocheFrm('CBTIERS'+no) then
      begin
      Inc(cpteur);
      AjouteDetail(no, IntToStr(cpteur)+') Cette fiche est liée au tiers numéro '+ RecupTextFrm('AN'+no+'_TIERS'));
      AjouteDetail(no, '=> après suppression, le lien sera redemandé à la prochaine entrée dans la GI sur le tiers concerné.');
      AjouteDetail(no, ' ');
      end;
    // Liens juridiques
    if RecupCocheFrm('CBJURIDIQUE'+no) then
      begin
      Inc(cpteur);
      Msg := IntToStr(cpteur)+') Juridique :'+#13+#10;
      if ObjSuppr.TestJuriDosPer(tmp) then
        begin
        Msg := Msg+tmp;
        Msg := Msg+'=> supprimer d''abord le dossier juridique pour pouvoir supprimer la fiche.'+#13+#10;
        end;
      if ObjSuppr.TestJuriLienPer(tmp) then
        begin
        Msg := Msg+tmp;
        Msg := Msg+'=> modifier d''abord les dossiers cités.'+#13+#10;
        end;
      if ObjSuppr.TestJuriDosInfoPer(tmp) then
        begin;
        Msg := Msg+tmp;
        Msg := Msg+'=> modifier d''abord les dossiers cités.'+#13+#10;
        end;
      if ObjSuppr.TestJuriGroupeSocPer(tmp) then
        begin
        Msg := Msg+tmp;
        Msg := Msg + '=> modifier auparavant ce(s) groupe(s) de sociétés.'+#13+#10;
        end;
      AjouteDetail(no, Msg);
      end;
    // Liens DP
    if RecupCocheFrm('CBLIENS'+no) then
      begin
      Inc(cpteur);
      ObjSuppr.TestLienPer(Msg);
      Msg := IntToStr(cpteur)+') '+Msg;
      Msg := Copy(Msg,1,Length(Msg)-1); // enlève retour chariot
      AjouteDetail(no, Msg);
      AjouteDetail(no, '=> supprimer d''abord ces liens dans le dossier permanent pour pouvoir supprimer la fiche.');
      AjouteDetail(no, ' ');
      end;
    // Liens fiscalité personnelle (pas de case à cocher affichée)
    If ObjSuppr.TestFisPer(Msg) then
      begin
      Inc(cpteur);
      Msg := IntToStr(cpteur)+') '+Msg;
      Msg := Copy(Msg,1,Length(Msg)-1); // enlève retour chariot
      AjouteDetail(no, Msg);
      AjouteDetail(no, '=> supprimer d''abord ces liens dans le logiciel concerné pour pouvoir supprimer la fiche.');
      AjouteDetail(no, ' ');
      end;
    // Dossier
    if RecupCocheFrm('CBDOSSIER'+no) then
      begin
      Inc(cpteur);
      AjouteDetail(no, IntToStr(cpteur)+') Cette fiche est utilisée dans le dossier '+ RecupCaptionFrm('LDOSSIER'+no));
      AjouteDetail(no, '=> supprimer d''abord le dossier de production pour pouvoir supprimer la fiche.');
      AjouteDetail(no, ' ');
      end;
    // Plaquettes
    if RecupCocheFrm('CBIFU'+no) then
      begin
      Inc(cpteur);
      ObjSuppr.TestLienPlaq(Msg);
      Msg := IntToStr(cpteur)+') '+Msg;
      Msg := Copy(Msg,1,Length(Msg)-1); // enlève retour chariot
      AjouteDetail(no, Msg);
      AjouteDetail(no, '=> supprimer d''abord les liens dans le logiciel IFU pour pouvoir supprimer la fiche.');
      AjouteDetail(no, ' ');
      end;
}
    AjouteDetail(no, Msg);
    END;

  // affiche directement la page 2 si clic sur bouton 2
  if NoSuppr='2' then bSuivantClick(Self);
end;


function TFAnnDedoub.RecupCocheFrm(nomzone: String): Boolean;
// fait l'affectation ET retourne l'état de la coche
Var C : TComponent ;
begin
  Result := False;
  C:=Frm.FindComponent(nomzone) ;
  if C=Nil then exit;
  if (C is TCustomCheckBox) then
    Result := TCheckBox(C).Checked
  else if (C is TRadioButton) then
    Result := TRadioButton(C).Checked;
end;


function TFAnnDedoub.RecupTextFrm(nomzone: String): String;
begin
  Result := TOMTOFGetControlText(Frm, nomzone);
  // #### si zone vide, Result vaut #$D#$A !!
  if Result=#13#10 then Result := '';
end;


{function TFAnnDedoub.RecupCaptionFrm(nomzone: String): String;
var c: TComponent;
begin
  Result := '';
  c := Frm.FindComponent(nomzone);
  if c=Nil then exit;
  Result := THLabel(c).Caption;
end; }


procedure TFAnnDedoub.FormCreate(Sender: TObject);
begin
  inherited;
  ObjSuppr := TPerSup.Create;
end;


procedure TFAnnDedoub.AjouteDetail(num, txt: String; saufblancs: Boolean=False);
var memo: TMemo;
begin
  if (saufblancs) and (Trim(txt)='') then exit;
  memo := TMemo(Self.FindComponent('MemoDetail'+num));
  if memo=Nil then exit;
  with memo.Lines do Add(txt);
end;

procedure TFAnnDedoub.bFinClick(Sender: TObject);
begin
  inherited;
  bAnnulerClick(Sender);
end;

procedure TFAnnDedoub.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  ObjSuppr.Free;
  inherited;
end;

end.
