USE IV;
CREATE TABLE ManufactureData
(
serialnumber int NOT NULL auto_increment primary key,
manufacturer varchar(60),
modeltype varchar(60),
modulearea varchar(60),
Isc double,
Voc double,
Imp double,
Vmp double,
Pm double,
Date datetime NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE measuredAtSTC
(
Isc double,
Voc double,
Imp double,
Vmp double,
FF double,
Pm double,
serialnumber int NOT NULL AUTO_INCREMENT Primary key,
 FOREIGN KEY (serialnumber) REFERENCES ManufactureData(serialnumber)
 
);

CREATE TABLE temperatureCoefficient
(
Isc double,
Voc double,
Imp double,
Vmp double,
FF double,
Pm double,
serialnumber int NOT NULL AUTO_INCREMENT Primary key,
 FOREIGN KEY (serialnumber) REFERENCES ManufactureData(serialnumber)
); 

CREATE TABLE deltaMeasured
(
Isc double,
Voc double,
Imp double,
Vmp double,
FF double,
Pm double,
serialnumber int NOT NULL AUTO_INCREMENT Primary key,
 FOREIGN KEY (serialnumber) REFERENCES ManufactureData(serialnumber)
); 
