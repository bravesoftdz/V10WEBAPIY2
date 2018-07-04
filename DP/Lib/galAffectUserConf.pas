unit galAffectUserConf;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, StdCtrls, Hctrls, HSysMenu, HTB97, UTob, HQry, HDB,
{$IFDEF EAGLCLIENT}
  EMul,
{$ELSE}
  Mul,
{$ENDIF}
  HMsgBox, HEnt1;

type
  TFAffectUserConf = class(TFVierge)
    mvNewGroupe: THMultiValComboBox;
    lblNewGroupe: THLabel;
    procedure BValiderClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    { Déclarations privées }
    Frm          : TForm ;             // form appelant
    Lst          : THDBGrid ;          // Liste du mul appelant
    Qry          : THQuery ;           // Query du mul appelant
    procedure InsereOuMajDansUserConf(coduser: String; T: TOB);
  public
    { Déclarations publiques }
  end;

procedure FicheAffectUserConf(F: TForm);

var FAffectUserConf: TFAffectUserConf;
//////////// IMPLEMENTATION //////////////
implementation

uses UtilMulTraitmt;

{$R *.DFM}

procedure FicheAffectUserConf(F: TForm);
// Param facultatif
begin
  if F is TFMul then
    begin
    // ouvre assistant
    FAffectUserConf:=TFAffectUserConf.Create(Application) ;
    Try
      FAffectUserConf.Frm := F;
      FAffectUserConf.Lst := TFMul(F).FListe ;
      FAffectUserConf.Qry := TFMul(F).Q;
      FAffectUserConf.mvNewGroupe.Text := '<<Tous>>'; // par défaut !
      FAffectUserConf.ShowModal ;
    Finally
      FAffectUserConf.Free ;
    end ;
    end;
end;


procedure TFAffectUserConf.BValiderClick(Sender: TObject);
var i, nb : Integer;
    sNewGroupeText, tmp, lesgroupes, coduser : String;
    T : TOB;
begin
  inherited;
  lesgroupes := '';
  // <<Tous>> sélectionné
  if mvNewGroupe.Tous then
    begin
    // items contient les libellés du groupe de conf.
    for i:=0 to mvNewGroupe.Items.Count-1 do
      lesgroupes := lesgroupes + '- '+mvNewGroupe.Items[i]+#13#10;
    end
  else
    begin
    sNewGroupeText := mvNewGroupe.Text;
    tmp := ReadTokenSt(sNewGroupeText);
    while tmp<>'' do
      begin
      // tmp contient le Code du groupe de conf.
      for i:=0 to mvNewGroupe.Items.Count-1 do
        begin
        if mvNewGroupe.Values[i]=tmp then
          begin
          lesgroupes := lesgroupes + '- '+mvNewGroupe.Items[i]+#13#10;
          break;
          end;
        end;
      tmp := ReadTokenSt(sNewGroupeText);
      end;
    end;

  if PGIAsk('Vous allez rattacher tous les utilisateurs sélectionnés '+#13#10
   + ' au(x) groupe(s) de travail suivant(s) :'+#13#10
   + lesgroupes
   + ' Confirmez-vous ?', TitreHalley)=mrNo then exit;

  T := TOB.Create('USERCONF', Nil, -1);

  // ---- début traitement sur mul
  if Lst.AllSelected then
    BEGIN
{$IFDEF EAGLCLIENT}
    if not TFMul(Frm).Fetchlestous then
      PGIInfo('Impossible de récupérer tous les enregistrements')
    else
{$ENDIF}
      begin
      // le query contient tous les enreg
      Qry.First;
      while Not Qry.EOF do
        begin
        coduser := Qry.FindField('US_UTILISATEUR').AsString;
        InsereOuMajDansUserConf(coduser, T);
        Qry.Next;
        end;
      end;
    END
  else
    BEGIN
    nb := Lst.NbSelected;
    for i:=0 to nb-1 do
      begin
      Lst.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
      Qry.TQ.Seek(Lst.Row-1);
{$ENDIF}
      coduser := Qry.FindField('US_UTILISATEUR').AsString;
      InsereOuMajDansUserConf(coduser, T);
      end;
    END;
  // déselectionne
  FinTraitmtMul(TFMul(Frm));

  T.ClearDetail;
  T.Free;
end;


procedure TFAffectUserConf.InsereOuMajDansUserConf(coduser: String; T: TOB);
var TEnreg: TOB;
    sNewGroupeText, tmp : String;
    i : Integer;
    procedure InsereUserConf;
    begin
      if Not T.SelectDB('"'+tmp+'";"'+coduser+'"', Nil, True) then
        begin
        TEnreg := TOB.Create('USERCONF', T, -1);
        TEnreg.PutValue('UCO_GROUPECONF', tmp);
        TEnreg.PutValue('UCO_USER', coduser);
        // maj dans base
        TEnreg.InsertDB(Nil);
        end;
    end;
begin
  // <<Tous>> sélectionné
  if mvNewGroupe.Tous then
    begin
    // value contient le Code du groupe de conf.
    for i:=0 to mvNewGroupe.Items.Count-1 do
      begin
      // tmp contient le Code du groupe de conf.
      tmp := mvNewGroupe.Values[i];
      InsereUserConf;
      end;
    end
  else
    begin
    sNewGroupeText := mvNewGroupe.Text;
    tmp := ReadTokenSt(sNewGroupeText);
    while tmp<>'' do
      begin
      // tmp contient le Code du groupe de conf.
      InsereUserConf;
      tmp := ReadTokenSt(sNewGroupeText);
      end;
    end;
end;

procedure TFAffectUserConf.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
  case Key of
   VK_ESCAPE : FAffectUserConf.ModalResult:=1;
  end;
end;

end.
