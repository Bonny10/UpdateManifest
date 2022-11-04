#!/usr/bin/env python

from xml.etree import ElementTree

class CommentedTreeBuilder ( ElementTree.XMLTreeBuilder ):
    def __init__ ( self, html = 0, target = None ):
        ElementTree.XMLTreeBuilder.__init__( self, html, target )
        self._parser.CommentHandler = self.handle_comment
    
    def handle_comment ( self, data ):
        self._target.start( ElementTree.Comment, {} )
        self._target.data( data )
        self._target.end( ElementTree.Comment )
