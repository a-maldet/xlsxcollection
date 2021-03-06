% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/xlsxcollection_list.R
\name{new_xlsxcollection_list}
\alias{new_xlsxcollection_list}
\title{Create an \code{xlsxcollection_list} class object.}
\usage{
new_xlsxcollection_list(obj, ...)
}
\arguments{
\item{obj}{One can eather pass a list of items or a singe item.}

\item{...}{additional \link[=new_xlsxcollection_item]{xlsxcollection_item} class objects.}
}
\value{
A \code{xlsxcollection_list} class object holding all items.
}
\description{
Bundles mutliple
\link[=new_xlsxcollection_item]{xlsxcollection_item} into a single
object of class \code{xlsxcollection_list}.
}
