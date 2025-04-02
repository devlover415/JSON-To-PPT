from dataclasses import dataclass
from typing import Optional, Dict
from pptx.dml.color import RGBColor
from pptx.util import Pt

@dataclass
class Formatting:
    font_name: Optional[str] = None
    font_size: Optional[int] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    strikethrough: Optional[bool] = None
    font_color: Optional[str] = None
    background_color: Optional[str] = None

    @classmethod
    def from_dict(cls, data: Dict) -> 'Formatting':
        return cls(
            font_name=data.get('fontName'),
            font_size=data.get('fontSize'),
            bold=data.get('bold'),
            italic=data.get('italic'),
            underline=data.get('underline'),
            strikethrough=data.get('strikethrough'),
            font_color=data.get('fontColor'),
            background_color=data.get('backgroundColor')
        )

    def apply_to_run(self, run) -> None:
        if self.font_name is not None:
            run.font.name = self.font_name
        
        if self.font_size is not None:
            run.font.size = Pt(self.font_size)
        
        if self.bold is not None:
            run.font.bold = self.bold
            
        if self.italic is not None:
            run.font.italic = self.italic
            
        if self.underline is not None:
            run.font.underline = self.underline
            
        if self.strikethrough is not None:
            run.font.strikethrough = self.strikethrough
            
        if self.font_color is not None:
            run.font.color.rgb = RGBColor.from_string(self.font_color)
##ppt formating