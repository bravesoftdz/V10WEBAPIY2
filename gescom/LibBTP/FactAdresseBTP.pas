unit FactAdresseBTP;

interface

uses HEnt1, UTOB, HCtrls, SysUtils, ParamSoc,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     EntGC, TiersUtil, FactTOB, UtilPGI ;

procedure LivAffaireVersAdresses(TOBAffaire, TOBAdresses, TOBPiece: TOB;MiseAjour : boolean = true);

implementation

procedure LivAffaireVersAdresses(TOBAffaire, TOBAdresses, TOBPiece: TOB;MiseAjour : boolean = true);
var Q: TQuery;
  TOBAdrLiv, TOBAdr, TheAffaire: TOB;
  created : boolean;
begin
	created := false;

  if TobAffaire = nil then
  begin
  	if TOBPiece = nil then exit;
    if TOBPiece.GetValue('GP_AFFAIRE')='' then exit;
  	created := True;
  	TheAffaire := TOB.Create ('AFFAIRE',nil,-1);
    Q := OpenSql ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+TOBPiece.getValue('GP_AFFAIRE')+'"',true,-1, '', True);
    if not Q.eof then TheAffaire.selectDb ('',Q,true) else freeAndNil(theAffaire);
    ferme (Q);
  end else
  begin
  	TheAffaire := TOBAffaire;                  
  end;

  if TheAffaire = nil then exit;
  if TheAffaire.getValue('AFF_AFFAIRE')='' then exit;
  Q := OpenSQL('SELECT * FROM ADRESSES WHERE ADR_TYPEADRESSE="INT" AND ADR_REFCODE="' +
    TheAffaire.GetValue('AFF_AFFAIRE') + '"', True,-1, '', True);
  TOBAdrLiv := TOB.Create('LES ADRESSES', nil, -1);
  try
    if not Q.EOF then
    begin
      TOBAdrLiv.LoadDetailDb('ADRESSES','','', Q,false);
      TOBAdr := TOBAdresses.Detail[0];
      if GetParamSoc('SO_GCPIECEADRESSE') then
      begin
        TOBAdr.PutValue('GPA_LIBELLE', TOBAdrLiv.detail[0].GetValue('ADR_LIBELLE'));
        TOBAdr.PutValue('GPA_LIBELLE2', TOBAdrLiv.detail[0].GetValue('ADR_LIBELLE2'));
        TOBAdr.PutValue('GPA_JURIDIQUE', TOBAdrLiv.detail[0].GetValue('ADR_JURIDIQUE'));
        TOBAdr.PutValue('GPA_ADRESSE1', TOBAdrLiv.detail[0].GetValue('ADR_ADRESSE1'));
        TOBAdr.PutValue('GPA_ADRESSE2', TOBAdrLiv.detail[0].GetValue('ADR_ADRESSE2'));
        TOBAdr.PutValue('GPA_ADRESSE3', TOBAdrLiv.detail[0].GetValue('ADR_ADRESSE3'));
        TOBAdr.PutValue('GPA_CODEPOSTAL', TOBAdrLiv.detail[0].GetValue('ADR_CODEPOSTAL'));
        TOBAdr.PutValue('GPA_VILLE', TOBAdrLiv.detail[0].GetValue('ADR_VILLE'));
        TOBAdr.PutValue('GPA_PAYS', TOBAdrLiv.detail[0].GetValue('ADR_PAYS'));
        TOBAdr.PutValue('GPA_NUMEROCONTACT', TOBAdrLiv.detail[0].GetValue('ADR_NUMEROCONTACT'));
        TobPiece.Putvalue('GP_NUMADRESSELIVR',1);
        TobPiece.Putvalue('GP_NUMADRESSEFACT',2);
        (*
        if MiseAJour then
        begin
        	If TobPiece.getvalue('GP_NUMADRESSELIVR') = TobPiece.getvalue('GP_NUMADRESSEFACT') then TOBPiece.PutValue('GP_NUMADRESSELIVR',-1); //mcd 02/10/2003 il faut ds cecas indiquer que les adresse st #
        end;
        *)
      end else
      begin
        TOBAdr.Dupliquer(TOBAdrLiv.detail[0], False, True);
        // reinitialisation de la clé
        TOBAdr.PutValue('ADR_NUMEROADRESSE', 0);
        TOBAdr.PutValue('ADR_REFCODE', '');
        TOBAdr.PutValue('ADR_TYPEADRESSE', '');
        TOBAdr.PutValue('ADR_NUMADRESSEREF', 0);
        TobPiece.Putvalue('GP_NUMADRESSELIVR',1);
        TobPiece.Putvalue('GP_NUMADRESSEFACT',2);

//        if MiseAjour then TOBPiece.PutValue('GP_NUMADRESSELIVR', -1);
      end;
    end else
    begin
        TobPiece.Putvalue('GP_NUMADRESSELIVR',1);
        TobPiece.Putvalue('GP_NUMADRESSEFACT',2);
//    	If TobPiece.getvalue('GP_NUMADRESSELIVR') = TobPiece.getvalue('GP_NUMADRESSEFACT') then
//      	TOBPiece.PutValue('GP_NUMADRESSELIVR',-1); //mcd 02/10/2003 il faut ds cecas indiquer que les adresse st #
    end;
  finally
    Ferme(Q);
    TOBAdrLiv.Free;
    if created then TheAffaire.free;
  end;
end;

end.
