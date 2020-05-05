IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(N'dbo.fnViaCep') AND xtype IN (N'FN', N'IF', N'TF'))
    DROP FUNCTION [dbo].[fnViaCep];
GO

CREATE FUNCTION dbo.fnViaCep(@CEP varchar(10)) 
RETURNS @T TABLE (
    cep varchar(10),
    logradouro varchar(100),
    complemento varchar(100),
    bairro varchar(100),
    localidade varchar(100),
    unidade varchar(100),
    uf char(2),
    ibge char(7),
    gia varchar(10)
) AS BEGIN

    DECLARE @authHeader VARCHAR(64),
            @contentType VARCHAR(64),
            @postData VARCHAR(2000),
            @responseText VARCHAR(2000),
            @responseXML VARCHAR(2000),
            @ret INT,
            @status VARCHAR(32),
            @statusText VARCHAR(32),
            @token INT,
            @url VARCHAR(256),
            @xml XML;

    SET @authHeader = 'BASIC 0123456789ABCDEF0123456789ABCDEF';
    SET @contentType = 'application/x-www-form-urlencoded';
    SET @url = 'http://viacep.com.br/ws/' + @CEP + '/xml/';

    -- Open the connection.
    EXEC @ret = sp_OACreate 'MSXML2.ServerXMLHTTP', @token OUT;
    -- IF @ret <> 0 PRINT (CAST('Unable to open HTTP connection.' AS INT));

    -- Send the request.
    EXEC @ret = sp_OAMethod @token, 'open', NULL, 'GET', @url, 'false';
    EXEC @ret = sp_OAMethod @token, 'setRequestHeader', NULL, 'Authentication', @authHeader;
    EXEC @ret = sp_OAMethod @token, 'setRequestHeader', NULL, 'Content-type', @contentType;
    EXEC @ret = sp_OAMethod @token, 'send', NULL

    -- Handle the response.
    EXEC @ret = sp_OAGetProperty @token, 'status', @status OUT;
    EXEC @ret = sp_OAGetProperty @token, 'statusText', @statusText OUT;
    EXEC @ret = sp_OAGetProperty @token, 'responseText', @responseText OUT;

    SET @xml = CONVERT(XML, replace(@responseText,'<?xml version="1.0" encoding="UTF-8"?>',''), 1);

    INSERT INTO @T
    SELECT
        t.c.value('cep[1]', 'varchar(10)') as cep,  
        t.c.value('logradouro[1]', 'varchar(100)') as logradouro,  
        t.c.value('complemento[1]', 'varchar(100)') as complemento,
        t.c.value('bairro[1]', 'varchar(100)') as bairro,
        t.c.value('localidade[1]', 'varchar(100)') as localidade,
        t.c.value('unidade[1]', 'varchar(100)') as unidade,
        t.c.value('uf[1]', 'char(2)') as uf,
        t.c.value('ibge[1]', 'char(7)') as ibge
        t.c.value('gia[1]', 'varchar(10)') as gia,
    FROM @xml.nodes('//xmlcep') t(c)  

    -- Close the connection.
    EXEC @ret = sp_OADestroy @token;
    -- IF @ret <> 0 RAISERROR('Unable to close HTTP connection.', 10, 1);

RETURN
END
GO
