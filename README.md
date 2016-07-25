Hyper
=====

Hyper lets you automatically populate Google Docs with data from arbitrary URLs. It comes with GitHub authorization to allow URLs pointing to (private or public) content hosted on GitHub. A working initial use case is using Jupyter notebooks saved to GitHub as data sources.

Installation
------------

The easiest way to install `Hyper` is [here](https://chrome.google.com/webstore/detail/hyper/lbiblkbfmhelinhcobciimdmcnngcffp?utm_source=permalink) via the Google Docs add-on store. Please reply [here](https://github.com/tavinathanson/hyper/issues/1) to be added as a private tester so that you can install `Hyper` this way.

Usage
-----

* Choose a URL as a source of data. For example, a Jupyter notebook on GitHub.
* In that source of data (e.g. inside the notebook), use the format `{{{<label>:<insert arbitrary text>}}}` to identify data to be extracted by the Google Doc. For example, `{{{mutations:median 234 (range 100 - 423)}}}`.
* Create a hyperlink in your Google Doc with any text and the following URL: `<your source of data URL>?hyper=<label>`. For example, `https://github.com/myorg/myrepo/tree/master/analyses/notebooks/notebook.ipynb?hyper=mutations`. Hyper uses the suffix to sift through the data source and grab the relevant piece of data.
* Go to `Add-ons -> Hyper -> Hyperize Links` in your Google Doc, and watch the text of your hyperlink transform into the data. For example, if using the above hyperlink, into `median 234 (range 100 - 423)`.
* When the data updates, simply Hyperize again.

Icons made by [SimpleIcon](http://www.flaticon.com/authors/simpleicon) from [Flaticon](http://www.flaticon.com), licensed by [CC 3.0 BY](http://creativecommons.org/licenses/by/3.0/).
