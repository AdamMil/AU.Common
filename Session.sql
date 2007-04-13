DROP PROCEDURE sess_Clear
CREATE PROCEDURE sess_Clear
@session_id int
AS

/* 
   This procedure is part of the AU.Common library, a set of ActiveX
   controls and C++ classes used to aid in COM and Web development.
   Copyright (C) 2002 Adam Milazzo

   This library is free software; you can redistribute it and/or
   modify it under the terms of the GNU Lesser General Public
   License as published by the Free Software Foundation; either
   version 2.1 of the License, or (at your option) any later version.

   This library is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
   Lesser General Public License for more details.

   You should have received a copy of the GNU Lesser General Public
   License along with this library; if not, write to the Free Software
   Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
*/

SET NOCOUNT ON
DELETE session_data WHERE session_id=@session_id

DROP PROCEDURE sess_LoadSession
CREATE PROCEDURE sess_LoadSession
@session_id int
AS

/* 
   This procedure is part of the AU.Common library, a set of ActiveX
   controls and C++ classes used to aid in COM and Web development.
   Copyright (C) 2002 Adam Milazzo

   This library is free software; you can redistribute it and/or
   modify it under the terms of the GNU Lesser General Public
   License as published by the Free Software Foundation; either
   version 2.1 of the License, or (at your option) any later version.

   This library is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
   Lesser General Public License for more details.

   You should have received a copy of the GNU Lesser General Public
   License along with this library; if not, write to the Free Software
   Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
*/

SET NOCOUNT ON
SELECT bindata,textdata,item_id,item_name FROM session_data WHERE session_id=@session_id

DROP PROCEDURE sess_CreateSession
CREATE PROCEDURE sess_CreateSession
@session_key char(20),
@timeout_val int,
@persists bit=0,
@session_id int=0 OUTPUT
AS
SET NOCOUNT ON

/* 
   This procedure is part of the AU.Common library, a set of ActiveX
   controls and C++ classes used to aid in COM and Web development.
   Copyright (C) 2002 Adam Milazzo

   This library is free software; you can redistribute it and/or
   modify it under the terms of the GNU Lesser General Public
   License as published by the Free Software Foundation; either
   version 2.1 of the License, or (at your option) any later version.

   This library is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
   Lesser General Public License for more details.

   You should have received a copy of the GNU Lesser General Public
   License along with this library; if not, write to the Free Software
   Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
*/

IF NOT EXISTS (SELECT 1 FROM sessions WHERE session_key=@session_key)
BEGIN
  INSERT INTO sessions (session_key,persists,timeout_val,timeout_date) values
    (@session_key,@persists,@timeout_val,DATEADD(s, @timeout_val, GETDATE()))
  SELECT @session_id=@@IDENTITY
END
ELSE SELECT @session_id=0

DROP PROCEDURE sess_DeleteSection
CREATE PROCEDURE sess_DeleteSection
@session_id int,
@section_name varchar(256)
AS

/* 
   This procedure is part of the AU.Common library, a set of ActiveX
   controls and C++ classes used to aid in COM and Web development.
   Copyright (C) 2002 Adam Milazzo

   This library is free software; you can redistribute it and/or
   modify it under the terms of the GNU Lesser General Public
   License as published by the Free Software Foundation; either
   version 2.1 of the License, or (at your option) any later version.

   This library is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
   Lesser General Public License for more details.

   You should have received a copy of the GNU Lesser General Public
   License along with this library; if not, write to the Free Software
   Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
*/

SET NOCOUNT ON
DELETE session_data WHERE session_id=@session_id AND item_name LIKE @section_name+'.%'

DROP PROCEDURE sess_DeleteSession
CREATE PROCEDURE sess_DeleteSession
@session_id int
AS

/* 
   This procedure is part of the AU.Common library, a set of ActiveX
   controls and C++ classes used to aid in COM and Web development.
   Copyright (C) 2002 Adam Milazzo

   This library is free software; you can redistribute it and/or
   modify it under the terms of the GNU Lesser General Public
   License as published by the Free Software Foundation; either
   version 2.1 of the License, or (at your option) any later version.

   This library is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
   Lesser General Public License for more details.

   You should have received a copy of the GNU Lesser General Public
   License along with this library; if not, write to the Free Software
   Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
*/

SET NOCOUNT ON
DELETE session_data WHERE session_id=@session_id
DELETE sessions WHERE session_id=@session_id

DROP PROCEDURE sess_GetValue
CREATE PROCEDURE sess_GetValue
@session_id int,
@item_id int
AS

/* 
   This procedure is part of the AU.Common library, a set of ActiveX
   controls and C++ classes used to aid in COM and Web development.
   Copyright (C) 2002 Adam Milazzo

   This library is free software; you can redistribute it and/or
   modify it under the terms of the GNU Lesser General Public
   License as published by the Free Software Foundation; either
   version 2.1 of the License, or (at your option) any later version.

   This library is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
   Lesser General Public License for more details.

   You should have received a copy of the GNU Lesser General Public
   License along with this library; if not, write to the Free Software
   Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
*/

SET NOCOUNT ON
SELECT bindata,textdata
  FROM session_data WHERE session_id=@session_id AND item_id=@item_id

DROP PROCEDURE sess_PutValue
CREATE PROCEDURE sess_PutValue
@session_id int,
@item_id int,
@bindata binary(16)=NULL,
@textdata text=NULL
AS
SET NOCOUNT ON

/* 
   This procedure is part of the AU.Common library, a set of ActiveX
   controls and C++ classes used to aid in COM and Web development.
   Copyright (C) 2002 Adam Milazzo

   This library is free software; you can redistribute it and/or
   modify it under the terms of the GNU Lesser General Public
   License as published by the Free Software Foundation; either
   version 2.1 of the License, or (at your option) any later version.

   This library is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
   Lesser General Public License for more details.

   You should have received a copy of the GNU Lesser General Public
   License along with this library; if not, write to the Free Software
   Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
*/

UPDATE session_data SET bindata=@bindata,textdata=@textdata
  WHERE session_id=@session_id AND item_id=@item_id

DROP PROCEDURE sess_AddItem
CREATE PROCEDURE sess_AddItem
@session_id int,
@item_name varchar(255),
@bindata binary(16)=NULL,
@textdata text=NULL,
@item_id int=0 OUTPUT
AS
SET NOCOUNT ON

/* 
   This procedure is part of the AU.Common library, a set of ActiveX
   controls and C++ classes used to aid in COM and Web development.
   Copyright (C) 2002 Adam Milazzo

   This library is free software; you can redistribute it and/or
   modify it under the terms of the GNU Lesser General Public
   License as published by the Free Software Foundation; either
   version 2.1 of the License, or (at your option) any later version.

   This library is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
   Lesser General Public License for more details.

   You should have received a copy of the GNU Lesser General Public
   License along with this library; if not, write to the Free Software
   Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
*/

SELECT @item_id=item_id FROM session_data WHERE session_id=@session_id AND item_name=@item_name
IF @item_id>0
  UPDATE session_data SET bindata=@bindata,textdata=@textdata
    WHERE session_id=@session_id AND item_id=@item_id
ELSE
BEGIN
  INSERT INTO session_data (session_id,item_name,bindata,textdata)
    VALUES (@session_id,@item_name,@bindata,@textdata)
  SELECT @item_id=@@IDENTITY
END
