#!perl -w

use strict;
use Test;
use Win32::OLE;

BEGIN { plan tests => 5 }

use vars qw( $cwp );

$cwp = Win32::OLE->new("OpenConf.CodeWorksProject") or die;

ok($cwp->Count(), 0);

$cwp->Add('RemoveProc', 'ProcName');

# просто 
ok($cwp->Count(), 1);
ok($cwp->ActionToHTML(0)."\n", <<EOF);
<a class="action" href="javascript:void(0)" onClick="editAction(0);">Удалить подпрограмму <code>ProcName</code></a>
EOF
;
# prettyprint
ok($cwp->ActionToHTML(0, 1), <<EOF);
<a class="action" href="javascript:void(0)" onClick="editAction(0);">
Удалить подпрограмму <code>ProcName</code>
</a>
EOF
;
# prettyprint + indent
ok($cwp->ActionToHTML(0, 1, "\t"), <<EOF);
	<a class="action" href="javascript:void(0)" onClick="editAction(0);">
	Удалить подпрограмму <code>ProcName</code>
	</a>
EOF
;