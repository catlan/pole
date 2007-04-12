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

#include <QAction>
#include <QApplication>
#include <QFileDialog>
#include <QFileInfo>
#include <QIcon>
#include <QLabel>
#include <QMainWindow>
#include <QMenuBar>
#include <QMenu>
#include <QMessageBox>
#include <QStatusBar>
#include <QTextEdit>
#include <QTime>
#include <QTimer>
#include <QTreeWidget>
#include <QToolBar>
#include <QVBoxLayout>

class ActionPack
{
public:
  QAction* fileNew;
  QAction* fileOpen;
  QAction* fileClose;
  QAction* fileQuit;
  QAction* streamExport;
  QAction* streamView;
  QAction* helpAbout;
  QAction* helpAboutQt;

  ActionPack( QObject* parent )
  {
    fileNew = new QAction( PoleView::tr("&New Window"), parent );
    fileNew->setShortcut( PoleView::tr("Ctrl+N") );
    QObject::connect( fileNew, SIGNAL(triggered()), parent, SLOT(newWindow() ) );

    fileOpen = new QAction( PoleView::tr("&Open..."), parent );
    fileOpen->setShortcut( PoleView::tr("Ctrl+O") );
    fileOpen->setIcon( QIcon( ":/fileopen.png" ) );
    QObject::connect( fileOpen, SIGNAL(triggered()), parent, SLOT(choose() ) );

    fileClose = new QAction( PoleView::tr("&Close"), parent );
    QObject::connect( fileClose, SIGNAL(triggered()), parent, SLOT(closeFile() ) );

    fileQuit = new QAction( PoleView::tr("&Quit"), parent );
    fileQuit->setShortcut( PoleView::tr("Ctrl+Q") );
    QObject::connect( fileQuit, SIGNAL(triggered()), qApp, SLOT(closeAllWindows() ) );

    streamExport = new QAction( PoleView::tr("&Export..."), parent );
    streamExport->setShortcut( PoleView::tr("Ctrl+E") );
    streamExport->setIcon( QIcon( ":/streamexport.png" ) );
    QObject::connect( streamExport, SIGNAL(triggered()), parent, SLOT(exportStream() ) );

    streamView = new QAction( PoleView::tr("&View..."), parent );
    streamView->setIcon( QIcon( ":/streamview.png" ) );
    QObject::connect( streamView, SIGNAL(triggered()), parent, SLOT(viewStream() ) );

    helpAbout = new QAction( PoleView::tr("&About..."), parent );
    helpAbout->setShortcut( PoleView::tr("F1") );
    QObject::connect( helpAbout, SIGNAL(triggered()), parent, SLOT(about() ) );

    helpAboutQt = new QAction( PoleView::tr("About &Qt"), parent );
    QObject::connect( helpAboutQt, SIGNAL(triggered()), parent, SLOT(aboutQt() ) );
  }
};

class PoleView::Private
{
public:
  POLE::Storage* storage;
  QTreeWidget* tree;
  ActionPack* actions;
};

PoleView::PoleView(): QMainWindow()
{
  d = new PoleView::Private;
  d->actions = new ActionPack(this);
  d->storage = 0;

  QStringList headers;
  headers << tr("Name");
  headers << tr("Size");

  d->tree = new QTreeWidget( this );
  d->tree->setColumnCount( 2 );
  d->tree->setHeaderLabels( headers );
  d->tree->setUniformRowHeights( true );
  setCentralWidget( d->tree );

  QMenu* fileMenu = menuBar()->addMenu( tr("&File") );
  fileMenu->addAction( d->actions->fileNew );
  fileMenu->addAction( d->actions->fileOpen );
  fileMenu->addAction( d->actions->fileClose );
  fileMenu->addSeparator();
  fileMenu->addAction( d->actions->fileQuit );

  QMenu* streamMenu = menuBar()->addMenu( tr("&Stream") );
  streamMenu->addAction( d->actions->streamExport );
  streamMenu->addAction( d->actions->streamView );

  QMenu* helpMenu = menuBar()->addMenu( tr("&Help") );
  helpMenu->addAction( d->actions->helpAbout );
  helpMenu->addAction( d->actions->helpAboutQt );

  QToolBar* mainToolBar = addToolBar( tr("&Main") );
  mainToolBar->addAction( d->actions->fileOpen );
  mainToolBar->addSeparator();
  mainToolBar->addAction( d->actions->streamExport );
  mainToolBar->addAction( d->actions->streamView );
  mainToolBar->setToolButtonStyle( Qt::ToolButtonTextBesideIcon );

  resize( 400, 300 );
  setWindowTitle( tr("POLEView" ) );
  statusBar()->showMessage( tr("Ready"), 5000 );
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

  QString fn = QFileDialog::getOpenFileName(this, tr("Open File"), QString(), filter);

  if ( !fn.isEmpty() ) openFile( fn );
  else
    statusBar()->showMessage( tr("Loading aborted"), 2000 );
}

class StreamItem: public QTreeWidgetItem
{
public:
  QString name;
  POLE::Stream* stream;
  StreamItem( QTreeWidgetItem* parent, const QString& n, POLE::Stream* s = 0 ):
  QTreeWidgetItem( parent )
  {
    name = n;
    stream = s;
    if( stream )
      setText( 1, QString::number( stream->size() ) );
    setText( 0, n );
    setTextAlignment( 0, Qt::AlignLeft );
    setTextAlignment( 1, Qt::AlignRight );
    setExpanded( true );
  }

  StreamItem( QTreeWidget* parent, const QString& n ):
  QTreeWidgetItem( parent )
  {
    name = n;
    stream = 0;
    setText( 0, n );
    setTextAlignment( 0, Qt::AlignLeft );
    setTextAlignment( 1, Qt::AlignRight );
    setExpanded( true );
  }
};

void visit( QTreeWidgetItem* parent, POLE::Storage* storage, const std::string path  )
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
  if( d->storage ) closeFile();

  QTime t; t.start();
  d->storage = new POLE::Storage( fileName.toLocal8Bit() );
  d->storage->open();

  if( d->storage->result() != POLE::Storage::Ok )
  {
    QString msg = QString( tr("Unable to open file %1\nProbably it is not a compound document.") ).arg(fileName);
    QMessageBox::critical( 0, tr("Error"), msg );
    closeFile();
    return;
  }

  QString msg = QString( tr("Loading %1 (%2 ms)") ).arg( QFileInfo(fileName).fileName() ).arg( t.elapsed() );
  statusBar()->showMessage( msg, 4000 );

  d->tree->clear();
  StreamItem* root = new StreamItem( d->tree, tr("Root") );
  visit( root, d->storage, "/" );
  d->tree->resizeColumnToContents( 0 );
  d->tree->resizeColumnToContents( 1 );

  setWindowTitle( QString( tr("%1 - POLEView" ).arg( fileName ) ) );
}

void PoleView::closeFile()
{
  if( d->storage )
  {
    d->storage->close();
    delete d->storage;
  }

  d->storage = 0;
  d->tree->clear();
  setWindowTitle( tr("POLEView" ) );
}

void PoleView::viewStream()
{
  QList<QTreeWidgetItem*> items = d->tree->selectedItems();
  StreamItem* item = items.count() ? (StreamItem*)items[0] : 0;
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

  sv->setWindowTitle( name );
  QTimer::singleShot( 200, sv, SLOT( show() ) );
}

void PoleView::exportStream()
{
  QList<QTreeWidgetItem*> items = d->tree->selectedItems();
  StreamItem* item = items.count() ? (StreamItem*)items[0] : 0;
  if( !item )
  {
    QMessageBox::warning( 0, tr("Export Stream"),
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

  QString fn = QFileDialog::getSaveFileName(this, tr("Export Stream As"));

  if ( fn.isEmpty() )
  {
    statusBar()->showMessage( tr("Export aborted"), 2000 );
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

  statusBar()->showMessage( tr("Exporting... Please wait") );
  for( ;; )
  {
      unsigned read = item->stream->read( buffer, sizeof( buffer ) );
      outf.write( (const char*)buffer, read  );
      if( read < sizeof( buffer ) ) break;
  }
  outf.close();
  statusBar()->showMessage( tr("Stream is exported."), 2000 );
}

void PoleView::about()
{
  QMessageBox::about( this, tr("About POLEView"),
    tr("Simple structured storage viewer\n\n"
    "Copyright (C) 2004-2007 Ariya Hidayat (ariya@kde.org)\n\n"
    "Icons are from Eclipse project (http://www.eclipse.org)"));
}

void PoleView::aboutQt()
{
  QMessageBox::aboutQt( this, tr("POLEView") );
}

#define STREAM_MAX_SIZE 32  // in KB

class StreamView::Private
{
public:
  POLE::Stream* stream;
  QLabel* infoLabel;
  QTextEdit* log;
};


StreamView::StreamView( POLE::Stream* s ): QDialog( 0 )
{
  d = new Private;
  d->stream = s;

  setModal( false );

  QVBoxLayout* layout = new QVBoxLayout( this );
  layout->setMargin( 10 );
  layout->setSpacing( 5 );

  d->infoLabel = new QLabel( this );

  d->log = new QTextEdit( this );
  d->log->setReadOnly(true);
  d->log->setFont( QFont("Courier") );
  d->log->setMinimumSize( 500, 300 );

  layout->addWidget( d->infoLabel );
  layout->addWidget( d->log );

  QTimer::singleShot( 0, this, SLOT( loadStream() ) );
}

void StreamView::loadStream()
{
  unsigned size = d->stream->size();

  if( size > STREAM_MAX_SIZE*1024 )
  {
    d->infoLabel->setText( tr("This stream is too large. "
      "Only the first %1 KB is shown.").arg( STREAM_MAX_SIZE ) );
    size = STREAM_MAX_SIZE*1024;
  }
  else
    d->infoLabel->setText( tr("Size: %1 bytes").arg( size ) );

  unsigned char buffer[16];
  d->stream->seek( 0 );
  d->log->append( "<pre>");
  for( unsigned j = 0; j < size; j+= 16 )
  {
    unsigned read = d->stream->read( buffer, 16 );
    appendData( buffer, read );
    if( read < sizeof( buffer ) ) break;
  }

  d->log->append( "</pre>");
  d->log->moveCursor( QTextCursor::Start );
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

  d->log->append( msg );
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
