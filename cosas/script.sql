CREATE TABLE rubros (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  descripcion VARCHAR(50) NOT NULL
);

INSERT INTO rubros (id, descripcion) VALUES (1, 'General');

CREATE TABLE productos (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  codigo VARCHAR(50) NOT NULL,
  descripcion VARCHAR(50) NOT NULL,
  idrubro INTEGER NOT NULL DEFAULT 0,
  stock_minimo FLOAT NULL NOT NULL DEFAULT 0,
  stock FLOAT NULL NOT NULL DEFAULT 0,
  costo FLOAT NOT NULL DEFAULT 0,
  precio FLOAT NOT NULL DEFAULT 0,
  UNIQUE(codigo)
);

CREATE TABLE usuarios (
  usuario VARCHAR(8) NOT NULL,
  contrasena VARCHAR(8) NOT NULL,
  es_admin TINYINT(1) NOT NULL DEFAULT 0,
  UNIQUE (contrasena),
  PRIMARY KEY (usuario)
);

INSERT INTO usuarios (usuario, contrasena, es_admin) VALUES ('admin', 'admin', 1);

CREATE TABLE movimientos (
	id INTEGER PRIMARY KEY AUTOINCREMENT,
	fecha DATE NOT NULL,
	tipo CHAR(1) NOT NULL DEFAULT 'E'
);

CREATE TABLE movimientos_det (
	id INTEGER PRIMARY KEY AUTOINCREMENT,
	idmovimiento NOT NULL,
	idproducto INTEGER NOT NULL,
	cantidad FLOAT NOT NULL DEFAULT 0,
	precio FLOAT NOT NULL DEFAULT 0,
	UNIQUE(idmovimiento, idproducto)
	FOREIGN KEY(idmovimiento) REFERENCES movimientos(id),
	FOREIGN KEY(idproducto) REFERENCES productos(id)
);

CREATE VIEW vw_movimientos_det AS 
SELECT d.idmovimiento as idmovimiento, d.id as iddetalle, d.idproducto as idproducto, p.codigo as codigo, p.descripcion as descripcion, d.precio as precio, d.cantidad as cantidad, d.precio * d.cantidad as importe 
FROM movimientos_det d 
INNER JOIN productos p ON d.idproducto = p.id;