/*  POLE - Portable C++ library to access OLE Storage
    Copyright (C) 2002-2007 Ariya Hidayat (ariya@kde.org).

    Redistribution and use in source and binary forms, with or without
    modification, are permitted provided that the following conditions
    are met:

    1. Redistributions of source code must retain the above copyright
      notice, this list of conditions and the following disclaimer.
    2. Redistributions in binary form must reproduce the above copyright
      notice, this list of conditions and the following disclaimer in the
      documentation and/or other materials provided with the distribution.

    THIS SOFTWARE IS PROVIDED BY THE AUTHOR ``AS IS'' AND ANY EXPRESS OR
    IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
    OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
    IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DIRECT, INDIRECT,
    INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
    NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
    DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
    THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
    (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
    THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/


#include "poleview.h"
#include "pole.h"

#include <qaction.h>
#include <qapplication.h>
#include <qdatetime.h>
#include <q3filedialog.h>
#include <qlabel.h>
#include <qlayout.h>
#include <q3listview.h>
#include <q3mainwindow.h>
#include <qmenubar.h>
#include <qmessagebox.h>
#include <q3popupmenu.h>
#include <qstatusbar.h>
#include <qstring.h>
#include <q3textedit.h>
#include <qtimer.h>
//Added by qt3to4:
#include <Q3VBoxLayout>

PoleView::PoleView(): Q3MainWindow()
{
  storage = 0;

  view = new Q3ListView( this );
  view->addColumn( tr("Name" ) );
  view->addColumn( tr("Size" ) );
  view->setColumnAlignment( 1, Qt::AlignRight );
  setCentralWidget( view );

  Q3PopupMenu * file = new Q3PopupMenu( this );
  menuBar()->insertItem( tr("&File"), file );
  file->insertItem( tr("&New Window"), this, SLOT(newWindow()), Qt::CTRL+Qt::Key_N );
  file->insertItem( tr("&Open..."), this, SLOT(choose()), Qt::CTRL+Qt::Key_O );
  file->insertItem( tr("&Close"), this, SLOT(closeFile()) );
  file->insertSeparator();
  file->insertItem( tr("&Quit"), qApp, SLOT( closeAllWindows() ), Qt::CTRL+Qt::Key_Q );

  Q3PopupMenu * streamMenu = new Q3PopupMenu( this );
  menuBar()->insertItem( tr("&Stream"), streamMenu );
  streamMenu->insertItem( tr("&Export..."), this, SLOT(exportStream() ), Qt::CTRL+Qt::Key_E );
  streamMenu->insertItem( tr("&View..."), this, SLOT(viewStream()) );

  Q3PopupMenu * help = new Q3PopupMenu( this );
  menuBar()->insertItem( tr("&Help"), help );
  help->insertItem( tr("&About"), this, SLOT(about()), Qt::Key_F1 );
  help->insertItem( tr("About &Qt"), this, SLOT(aboutQt()) );

  resize( 400, 300 );
  setCaption( tr("POLEView" ) );
  statusBar()->message( tr("Ready"), 5000 );
}

void PoleView::newWindow()
{
  PoleView* v = new PoleView();
  v->show();
}

void PoleView::choose()
{
  QString form("%1 (%2)" );
  QString allTypes( "*.doc *.xls *.xla *.ppt *.dot *.xlt *.pps *.wps" );
  QString msOfficeTypes( "*.doc *.xls *.xla *.ppt *.dot *.xlt *.pps" );

  QString filter0 = QString(form).arg( tr("All OLE Files") ).arg( allTypes );
  QString filter1 = QString(form).arg( tr("Microsoft Office Files") ).arg( msOfficeTypes );
  QString filter2a = QString(form).arg( tr("Microsoft Word Document") ).arg( "*.doc" );
  QString filter2b = QString(form).arg( tr("Microsoft Word Template") ).arg( "*.dot" );
  QString filter3a = QString(form).arg( tr("Microsoft Excel Workbook") ).arg( "*.xls" );
  QString filter3b = QString(form).arg( tr("Microsoft Excel Template") ).arg( "*.xlt" );
  QString filter3c = QString(form).arg( tr("Microsoft Excel Add-In") ).arg( "*.xla" );
  QString filter4a = QString(form).arg( tr("Microsoft PowerPoint Presentation") ).arg( "*.ppt" );
  QString filter4b = QString(form).arg( tr("Microsoft PowerPoint Template") ).arg( "*.pps" );
  QString filter5 = QString(form).arg( tr("Microsoft Works Files") ).arg( "*.wps" );
  QString filter6 = QString(form).arg( tr("All Files") ).arg( "*" );

  QString filter = filter0;
  filter = filter.append( ";;" ).append( filter1 );
  filter = filter.append( ";;" ).append( filter2a );
  filter = filter.append( ";;" ).append( filter2b );
  filter = filter.append( ";;" ).append( filter3a );
  filter = filter.append( ";;" ).append( filter3b );
  filter = filter.append( ";;" ).append( filter3c );
  filter = filter.append( ";;" ).append( filter4a );
  filter = filter.append( ";;" ).append( filter4b );
  filter = filter.append( ";;" ).append( filter5 );
  filter = filter.append( ";;" ).append( filter6 );

  QString fn = Q3FileDialog::getOpenFileName( QString::null, filter, this );

  if ( !fn.isEmpty() ) openFile( fn );
  else
    statusBar()->message( tr("Loading aborted"), 2000 );
}

class StreamItem: public Q3ListViewItem
{
public:
  StreamItem( Q3ListViewItem* parent, const QString& name, POLE::Stream* stream = 0 );
  StreamItem( Q3ListView* parent, const QString& name  );
  QString name;
  POLE::Stream* stream;
};

StreamItem::StreamItem( Q3ListViewItem* parent, const QString& n, POLE::Stream* s ):
Q3ListViewItem( parent, n )
{
  name = n;
  stream = s;
  if( stream )
    setText( 1, QString::number( stream->size() ) );
  setOpen( true );
}

StreamItem::StreamItem( Q3ListView* parent, const QString& n ):
Q3ListViewItem( parent, n )
{
  name = n;
  stream = 0;
  setOpen( true );
}

void visit( Q3ListViewItem* parent, POLE::Storage* storage, const std::string path  )
{
  std::list<std::string> entries;
  entries = storage->entries( path );

  std::list<std::string>::iterator it;
  for( it = entries.begin(); it != entries.end(); ++it )
  {
    std::string name = *it;
    std::string fullname = path + name;

    if( storage->isDirectory( fullname ) )
    {
      StreamItem* item = new StreamItem( parent, QString(name.c_str()) );
      visit( item, storage, fullname + "/" );
    }
    else
      new StreamItem( parent, QString(name.c_str()), new POLE::Stream( storage, fullname ) );
  }
}

void PoleView::openFile( const QString &fileName )
{
  if( storage ) closeFile();

  QTime t; t.start();
  storage = new POLE::Storage( fileName.latin1() );
  storage->open();

  if( storage->result() != POLE::Storage::Ok )
  {
    QString msg = QString( tr("Unable to open file %1\nProbably it is not a compound document.") ).arg(fileName);
    QMessageBox::critical( 0, tr("Error"), msg );
    closeFile();
    return;
  }

  QString msg = QString( tr("Loading %1 (%2 ms)") ).arg( fileName ).arg( t.elapsed() );
  statusBar()->message( msg, 2000 );

  view->clear();
  StreamItem* root = new StreamItem( view, tr("Root") );
  visit( root, storage, "/" );

  setCaption( QString( tr("%1 - POLEView" ).arg( fileName ) ) );
}

void PoleView::closeFile()
{
  if( storage )
  {
    storage->close();
    delete storage;
  }

  storage = 0;
  view->clear();
  setCaption( tr("POLEView" ) );
}

void PoleView::viewStream()
{
  StreamItem* item = (StreamItem*) view->selectedItem();
  if( !item )
  {
    QMessageBox::warning( 0, tr("View Stream"),
      tr("Nothing is selected"),
      QMessageBox::Ok, QMessageBox::NoButton );
    return;
  }

  QString name = item->text( 0 );
  if( !item->stream )
  {
    QMessageBox::warning( 0, tr("View Stream"),
      tr("'%1' is not a stream").arg( name ),
      QMessageBox::Ok, QMessageBox::NoButton );
    return;
  }

  StreamView* sv = new StreamView( item->stream );

  sv->setCaption( name );
  QTimer::singleShot( 200, sv, SLOT( show() ) );
}

void PoleView::exportStream()
{
  StreamItem* item = (StreamItem*) view->selectedItem();
  if( !item )
  {
    QMessageBox::warning( 0, tr("View Stream"),
      tr("Nothing is selected"),
      QMessageBox::Ok, QMessageBox::NoButton );
    return;
  }

  QString name = item->text( 0 );
  if( !item->stream )
  {
    QMessageBox::warning( 0, tr("View Stream"),
      tr("'%1' is not a stream").arg( name ),
      QMessageBox::Ok, QMessageBox::NoButton );
    return;
  }

  QString fn = Q3FileDialog::getSaveFileName( QString::null, QString::null,
    0, 0, tr("Export Stream As") );
  if ( fn.isEmpty() )
  {
    statusBar()->message( tr("Export aborted"), 2000 );
    return;
  }

  unsigned char buffer[16];
  QFile outf( fn );
  if( !outf.open( QIODevice::WriteOnly ) )
  {
    QString msg = QString( tr("Unable to write to file %1\n") ).arg(fn);
    QMessageBox::critical( 0, tr("Error"), msg );
    return;
  }

  statusBar()->message( tr("Exporting... Please wait") );
  for( ;; )
  {
      unsigned read = item->stream->read( buffer, sizeof( buffer ) );
      outf.writeBlock( (const char*)buffer, read  );
      if( read < sizeof( buffer ) ) break;
  }
  outf.close();
  statusBar()->message( tr("Stream is exported."), 2000 );
}

void PoleView::about()
{
  QMessageBox::about( this, tr("About POLEView"),
    tr("Simple structured storage viewer\n"
    "Copyright (C) 2004 Ariya Hidayat (ariya@kde.org)"));
}

void PoleView::aboutQt()
{
  QMessageBox::aboutQt( this, tr("POLEView") );
}

#define STREAM_MAX_SIZE 32  // in KB

StreamView::StreamView( POLE::Stream* s ): QDialog( 0 )
{
  stream = s;
  setModal( false );

  Q3VBoxLayout* layout = new Q3VBoxLayout( this );
  layout->setAutoAdd( true );
  layout->setMargin( 10 );
  layout->setSpacing( 5 );

  infoLabel = new QLabel( this );

  log = new Q3TextEdit( this );
  log->setTextFormat( Qt::LogText );
  log->setFont( QFont("Courier") );
  log->setMinimumSize( 500, 300 );

  QTimer::singleShot( 0, this, SLOT( loadStream() ) );
}

void StreamView::loadStream()
{
  unsigned size = stream->size();

  if( size > STREAM_MAX_SIZE*1024 )
  {
    infoLabel->setText( tr("This stream is too large. "
      "Only the first %1 KB is shown.").arg( STREAM_MAX_SIZE ) );
    size = STREAM_MAX_SIZE*1024;
  }

  unsigned char buffer[16];
  stream->seek( 0 );
  for( unsigned j = 0; j < size; j+= 16 )
  {
    unsigned read = stream->read( buffer, 16 );
    appendData( buffer, read );
    if( read < sizeof( buffer ) ) break;
  }

  QTimer::singleShot( 100, this, SLOT( goTop() ) );
}

void StreamView::goTop()
{
  log->ensureVisible( 0, 0 );
}

void StreamView::appendData( unsigned char* data, unsigned length )
{
  QString msg;
  for( unsigned i = 0; i < length; i++ )
  {
    QString s = QString::number( data[i], 16 );
    while( s.length() < 2 ) s.prepend( '0' );
    msg.append( s );
    msg.append( ' ' );
  }
  msg.append( "   " );
  for( unsigned i = 0; i < length; i++ )
  {
    if( (data[i]>31) && (data[i]<128) )
    {
      if( data[i] == '<' ) msg.append( "&lt;" );
      else if( data[i] == '>' ) msg.append( "&gt;" );
      else if( data[i] == '&' ) msg.append( "&amp;" );
      else msg.append( data[i] );
    }
    else msg.append( '.' );
  }

  log->append( msg );
}

// --- main program

#include <qapplication.h>

int main( int argc, char ** argv )
{
  QApplication a( argc, argv );

  PoleView* v = new PoleView();
  v->show();

  a.connect( &a, SIGNAL(lastWindowClosed()), &a, SLOT(quit()) );
  return a.exec();
}
