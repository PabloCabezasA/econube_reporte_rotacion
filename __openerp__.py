# -*- coding: utf-8 -*-
##############################################################################
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) Econube LTDA (<http://www.econube.cl>).
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Affero General Public License as
#    published by the Free Software Foundation, either version 3 of the
#    License, or (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU Affero General Public License for more details.
#
#    You should have received a copy of the GNU Affero General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
##############################################################################
{
  "name" : "Reporte rotacion de productos",
  "version" : "1.0",
  "author" : "econube | Pablo Cabezas, Jose Pinto",
  "website" : "http://econube.cl",
  "category" : "reportes",
  "description": """
                -Reporte para ver la rotacion de los productos
                 """,
  "depends" : ['base','product','econube_modificacion_puntarenas_v2','purchase','sale', 'stock'],
  "init_xml" : [ ],
  "demo_xml" : [ ],
  "data" : ['views/report_view.xml','views/view_product_form.xml'],
  "installable": True,
  "application": True
}
