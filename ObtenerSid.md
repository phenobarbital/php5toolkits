# Introduccion #

Usando la librería ldap-toolkit, podemos conectarnos a un controlador Active Directory, para extraer información, por ejemplo, el Domain SID.


# Configuracion #

Para configurar una conexion a un MS Active Directory, configuramos:

```
  #una conexion usando un Active Directory Service
  $database['ads']['adapter'] 		= 'ads';
  $database['ads']['host'] 		= '10.1.1.1';
  $database['ads']['basedn'] 		= 'DC=test,DC=com,DC=ve';
  $database['ads']['domain']		= 'test.com.ve';
  $database['ads']['netbios_name']	= 'TEST';
  $database['ads']['username'] 		= 'administrator';
  $database['ads']['password'] 		= 'mipassword';
  $database['ads']['port'] 		= 389;
  $options = array('LDAP_OPT_PROTOCOL_VERSION' => 3, 'LDAP_OPT_REFERRALS'=>0, 'LDAP_OPT_SIZELIMIT'=>5000, 'LDAP_OPT_TIMELIMIT'=>300);
  $database['ads']['options'] 		= $options;
```

  * Importante el usuario y el netbios\_name (nombre del dominio netbios) para el login

# Snippet de Codigo #

```
include "conf/base.inc.php";
include_once BASE_DIR . "conf/include_ldap.inc.php";

#obtenemos la configuracion AD
$ldap = ldap::load('ad');

#buscar una entrada basica, para extraer su SID
$entry = "(&(objectClass=user)(samaccounttype=". ADS_NORMAL_ACCOUNT .")(samaccountname=jesuslara))";

#conectamos al AD
$ldap->open();

$entry = $ldap->query($entry);
echo "SID usuario {$user}: " . $entry->bin_to_str_sid('objectSid');

$ldap->close();
```

Ese echo, mostrará un valor semejante a: 'S-1-5-21-2102913520-367280043-1452191782-14800'
De esa:
**S-1-5-21-2102913520-367280043-1452191782**-_14800_

Lo que está en negrillas es el SID del dominio, lo que es itálica, es el UidNumber del usuario.