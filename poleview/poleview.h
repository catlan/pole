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


#ifndef POLEVIEW
#define POLEVIEW

#include "pole.h"

#include <QMainWindow>
#include <QDialog>

class PoleView : public QMainWindow
{
  Q_OBJECT

  public:
    PoleView();

  private slots:

    void newWindow();
    void choose();
    void openFile( const QString &fileName );
    void closeFile();
    void viewStream();
    void exportStream();
    void about();
    void aboutQt();

  private:
    class Private;
    Private* d;
    PoleView( const PoleView& );
    PoleView& operator=( const PoleView& );
};


class StreamView: public QDialog
{
  Q_OBJECT

  public:
    StreamView( POLE::Stream* stream );

  private slots:
    void loadStream();
    void goTop();

  private:
    class Private;
    Private *d;
    StreamView( const StreamView& );
    StreamView& operator=( const StreamView& );
    void appendData( unsigned char* data, unsigned length );
};

#endif // POLEVIEW
